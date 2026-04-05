"""
台灣股市戰略儀表板 - 資料抓取模組 v10.4
資料來源：
  - FinMind API (免費無需 token, 每小時 300 次)
  - 台灣證券交易所 (TWSE) 開放資料 API + OpenAPI v1
  - 證券櫃檯買賣中心 (TPEx) API
  - Google Finance (VIX, US10Y, USD/JPY)
  - FRED (VIX歷史, US10Y歷史)
  - Stooq (VIX歷史, US10Y歷史 備援)
  - 鉅亨網 (cnyes) API (美元指數)
  - 台灣期貨交易所 (TAIFEX) CSV/HTML
  - KGI 凱基證券
"""

import requests
import json
import time
import os
import re
from datetime import datetime, timedelta
from bs4 import BeautifulSoup
import math

# ============================================================
# 歷史數據快取機制 — 跨次執行累積，確保圖表不會只有1個點
# ============================================================
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
HISTORY_CACHE_PATH = os.path.join(SCRIPT_DIR, "history_cache.json")


def _load_history_cache():
    """載入歷史數據快取"""
    try:
        if os.path.exists(HISTORY_CACHE_PATH):
            with open(HISTORY_CACHE_PATH, 'r', encoding='utf-8') as f:
                cache = json.load(f)
                print(f"  [快取] ✅ 載入成功，含 {len(cache.get('vix', []))} VIX / {len(cache.get('us10y', []))} US10Y / {len(cache.get('usd_index', []))} DXY / {len(cache.get('jpy_rate', []))} JPY 筆資料")
                return cache
    except Exception as e:
        print(f"  [快取] ⚠️ 載入失敗: {e}")
    return {}


def _save_history_cache(cache):
    """儲存歷史數據快取"""
    try:
        cache["last_updated"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with open(HISTORY_CACHE_PATH, 'w', encoding='utf-8') as f:
            json.dump(cache, f, ensure_ascii=False, indent=2)
        print(f"  [快取] ✅ 已儲存 history_cache.json")
    except Exception as e:
        print(f"  [快取] ❌ 儲存失敗: {e}")


def _merge_and_trim(cached_data, fresh_data, max_days=35):
    """
    合併快取與新抓取的數據，去重並保留最近 max_days 筆
    fresh_data 會覆蓋同日期的 cached_data
    """
    merged = {}
    for p in (cached_data or []):
        if isinstance(p, dict) and "date" in p and "close" in p:
            merged[p["date"]] = p["close"]
    for p in (fresh_data or []):
        if isinstance(p, dict) and "date" in p and "close" in p:
            merged[p["date"]] = p["close"]  # 新資料覆蓋舊的
    # 按日期排序
    sorted_dates = sorted(merged.keys())
    result = [{"date": d, "close": merged[d]} for d in sorted_dates]
    return result[-max_days:] if len(result) > max_days else result


# 共用 headers
HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
    "Accept-Language": "zh-TW,zh;q=0.9,en-US;q=0.8,en;q=0.7",
}

def _safe_get(url, params=None, timeout=15):
    """安全的 GET 請求，含重試機制"""
    for attempt in range(3):
        try:
            resp = requests.get(url, params=params, headers=HEADERS, timeout=timeout)
            resp.raise_for_status()
            return resp
        except Exception as e:
            if attempt == 2:
                print(f"  [警告] 抓取失敗 {url}: {e}")
                return None
            time.sleep(2)


def _parse_number(s):
    """將含有逗號的數字字串轉為 float"""
    if not s or s == "--" or s == "N/A":
        return None
    try:
        return float(str(s).replace(",", "").replace(" ", ""))
    except (ValueError, TypeError):
        return None


def _fetch_google_finance(symbol):
    """
    從 Google Finance 頁面抓取即時報價 (已驗證可用)
    symbol: 如 "VIX:INDEXCBOE", "TNX:INDEXCBOE", "USD-JPY"
    回傳: {"price": float, "change_pct": float} 或 None
    """
    print(f"  [Google Finance] 取得 {symbol}...")
    url = f"https://www.google.com/finance/quote/{symbol}"
    try:
        resp = requests.get(url, headers={
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
            "Accept-Language": "en-US,en;q=0.9",
        }, timeout=12)
        if resp.status_code == 200:
            # 找 data-last-price
            m = re.search(r'data-last-price="([\d.]+)"', resp.text)
            if m:
                price = float(m.group(1))
                # 找變動百分比
                m2 = re.search(r'data-last-normal-market-change-percent="([+-]?[\d.]+)"', resp.text)
                change_pct = float(m2.group(1)) if m2 else None
                print(f"  [Google Finance] ✅ {symbol} = {price}")
                return {"price": price, "change_pct": change_pct}
            else:
                print(f"  [Google Finance] ⚠️ {symbol} data-last-price not found")
    except Exception as e:
        print(f"  [Google Finance] ❌ {symbol}: {e}")
    return None


def _finmind_fetch(dataset, data_id, start_date, end_date):
    """
    從 FinMind API 取得資料
    ★ 不需要 token 也能用（每小時 300 次請求），有 token 提升至 600 次

    FinMind 資料集對照表:
    ┌─────────────────────────────────────────────────┬──────────────┬───────────────────────────────┐
    │ dataset                                         │ data_id      │ 說明                          │
    ├─────────────────────────────────────────────────┼──────────────┼───────────────────────────────┤
    │ TaiwanStockTotalMarginPurchaseShortSale          │ (不需要)     │ 整體融資融券餘額              │
    │   → name, TodayBalance, YesBalance, buy, sell, Return                                        │
    │ TaiwanStockMarginPurchaseShortSale               │ 股票代號     │ 個股融資融券                  │
    │ TaiwanFuturesInstitutionalInvestors              │ TX/MXF/MTX   │ 期貨三大法人                  │
    │   → institutional_investors, long/short_open_interest_balance_volume                          │
    │ TaiwanOptionPutCallRatio                         │ (不需要)     │ Put/Call Ratio               │
    │   → PutCallRatio                                                                             │
    │ TaiwanFuturesDaily                               │ TX/MXF/MTX   │ 期貨每日行情                  │
    │ TaiwanStockPrice                                 │ 股票代號     │ 股票每日行情                  │
    └─────────────────────────────────────────────────┴──────────────┴───────────────────────────────┘
    """
    try:
        url = "https://api.finmindtrade.com/api/v4/data"
        # 將 YYYYMMDD 轉換為 YYYY-MM-DD
        start_formatted = f"{start_date[:4]}-{start_date[4:6]}-{start_date[6:8]}"
        end_formatted = f"{end_date[:4]}-{end_date[4:6]}-{end_date[6:8]}"

        params = {
            "dataset": dataset,
            "start_date": start_formatted,
            "end_date": end_formatted,
        }
        if data_id:
            params["data_id"] = data_id

        # 有 token 就加上 (提高頻率限制)，沒有也能用
        finmind_token = os.environ.get("FINMIND_TOKEN", "")
        if finmind_token:
            params["token"] = finmind_token

        resp = _safe_get(url, params)
        if resp:
            data = resp.json()
            if data.get("status") == 200 and data.get("data"):
                print(f"    [FinMind] ✅ {dataset} (id={data_id or '-'}) 共{len(data['data'])}筆")
                retur