"""
台灣股市戰略儀表板 - 資料抓取模組 v10
資料來源：
  - FinMind API (免費無需 token, 每小時 300 次)
  - 台灣證券交易所 (TWSE) 開放資料 API + OpenAPI v1
  - 證券櫃檯買賣中心 (TPEx) API
  - 鉅亨網 (cnyes) API (美元指數、日圓、VIX、US10Y)
  - 台灣期貨交易所 (TAIFEX) CSV/HTML
  - KGI 凱基證券 (融資維持率)
"""

import requests
import json
import time
import os
import re
from datetime import datetime, timedelta
from bs4 import BeautifulSoup

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
                return data["data"]
            else:
                msg = data.get("msg", "unknown")
                status = data.get("status", "?")
                print(f"    [FinMind] ❌ {dataset}: status={status}, msg={msg}")
        else:
            print(f"    [FinMind] ❌ {dataset}: 無回應")
    except Exception as e:
        print(f"    [FinMind] ❌ 錯誤: {e}")

    return None


# ============================================================
# 股票名稱查詢 (確保代號後面一定有公司名)
# ============================================================
_STOCK_NAME_CACHE = {}

def _ensure_stock_names(stocks_list):
    """
    確保股票列表中每個項目都有公司名稱
    如果 stock_name 為空或等於 stock_id，嘗試從 TWSE 查詢
    """
    global _STOCK_NAME_CACHE

    # 先檢查哪些需要查詢
    missing = [s for s in stocks_list if not s.get("stock_name") or s["stock_name"] == s["stock_id"]]
    if not missing:
        return stocks_list

    # 如果快取是空的，載入完整股票名稱表
    if not _STOCK_NAME_CACHE:
        _load_stock_names()

    # 填入名稱
    for stock in stocks_list:
        if not stock.get("stock_name") or stock["stock_name"] == stock["stock_id"]:
            name = _STOCK_NAME_CACHE.get(stock["stock_id"])
            if name:
                stock["stock_name"] = name
            else:
                # 仍然沒有名稱，至少顯示代號
                stock["stock_name"] = stock["stock_id"]

    return stocks_list


def _load_stock_names():
    """從 TWSE 載入所有上市股票代號與名稱"""
    global _STOCK_NAME_CACHE
    print("  [股名] 載入股票名稱對照表...")

    # TWSE 上市公司列表
    url = "https://www.twse.com.tw/exchangeReport/STOCK_DAY_ALL"
    params = {"response": "json"}
    resp = _safe_get(url, params, timeout=10)
    if resp:
        try:
            data = resp.json()
            if "data" in data:
                for row in data["data"]:
                    if len(row) >= 2:
                        sid = str(row[0]).strip()
                        sname = str(row[1]).strip()
                        if sid and sname:
                            _STOCK_NAME_CACHE[sid] = sname
                print(f"  [股名] 已載入 {len(_STOCK_NAME_CACHE)} 檔上市股票名稱")
        except:
            pass

    # TPEx 上櫃公司列表
    url2 = "https://www.tpex.org.tw/web/stock/aftertrading/otc_quotes_no1430/stk_wn1430_result.php"
    params2 = {"l": "zh-tw", "o": "json"}
    resp2 = _safe_get(url2, params2, timeout=10)
    if resp2:
        try:
            data2 = resp2.json()
            if "aaData" in data2:
                for row in data2["aaData"]:
                    if len(row) >= 2:
                        sid = str(row[0]).strip()
                        sname = str(row[1]).strip()
                        if sid and sname:
                            _STOCK_NAME_CACHE[sid] = sname
                print(f"  [股名] 加入上櫃後共 {len(_STOCK_NAME_CACHE)} 檔")
        except:
            pass


# ============================================================
# 1. 加權指數 & 成交量
# ============================================================
def fetch_taiex(date_str=None):
    """
    抓取加權指數資訊
    回傳: {index, change, change_pct, volume, prev_volume, volume_change}
    """
    print("📊 抓取加權指數...")
    if date_str is None:
        date_str = datetime.now().strftime("%Y%m%d")

    url = "https://www.twse.com.tw/exchangeReport/FMTQIK"
    ym = date_str[:6]  # YYYYMM
    params = {"response": "json", "date": date_str}
    resp = _safe_get(url, params)

    result = {
        "index": None, "change": None, "change_pct": None,
        "volume": None, "prev_volume": None, "volume_change": None,
        "date": date_str
    }

    if resp is None:
        return result

    try:
        data = resp.json()
        if "data" in data and len(data["data"]) >= 1:
            rows = data["data"]
            # 最後一筆是最新的
            latest = rows[-1]
            # 欄位: 日期, 成交股數, 成交金額, 成交筆數, 發行量加權股價指數, 漲跌點數
            result["index"] = _parse_number(latest[4])
            result["change"] = _parse_number(latest[5])
            vol = _parse_number(latest[2])  # 成交金額
            result["volume"] = vol

            if len(rows) >= 2:
                prev = rows[-2]
                prev_vol = _parse_number(prev[2])
                result["prev_volume"] = prev_vol
                prev_idx = _parse_number(prev[4])
                if prev_idx and result["index"]:
                    result["change_pct"] = round((result["index"] - prev_idx) / prev_idx * 100, 2)
                if vol and prev_vol:
                    result["volume_change"] = round((vol - prev_vol) / prev_vol * 100, 2)
    except Exception as e:
        print(f"  [錯誤] 解析加權指數失敗: {e}")

    # 也嘗試從每日收盤行情取得更精確的資料
    url2 = "https://www.twse.com.tw/exchangeReport/MI_INDEX"
    params2 = {"response": "json", "date": date_str, "type": "IND"}
    resp2 = _safe_get(url2, params2)
    if resp2:
        try:
            data2 = resp2.json()
            # 嘗試從 data8 或 data7 取得加權指數
            for key in ["data8", "data7", "data6", "data5"]:
                if key in data2:
                    for row in data2[key]:
                        if "加權" in str(row[0]) and "不含" not in str(row[0]):
                            idx_val = _parse_number(row[1])
                            chg_val = _parse_number(row[2])
                            if idx_val:
                                result["index"] = idx_val
                            if chg_val:
                                result["change"] = chg_val
                            break
        except:
            pass

    return result


# ============================================================
# 2. 三大法人買賣超
# ============================================================
def fetch_institutional(date_str=None):
    """
    抓取三大法人買賣超
    回傳: {foreign_buy, foreign_sell, foreign_net, foreign_prev_net,
            trust_buy, trust_sell, trust_net, trust_prev_net,
            dealer_net}
    """
    print("🏦 抓取三大法人買賣超...")
    if date_str is None:
        date_str = datetime.now().strftime("%Y%m%d")

    url = "https://www.twse.com.tw/fund/BFI82U"
    params = {"response": "json", "dayDate": date_str, "weekDate": "", "monthDate": "", "type": "day"}
    resp = _safe_get(url, params)

    # 若指定日期無資料，嘗試不帶日期取最新
    if resp:
        try:
            test_data = resp.json()
            if "data" not in test_data or len(test_data.get("data", [])) == 0:
                resp = _safe_get(url, {"response": "json", "dayDate": "", "weekDate": "", "monthDate": "", "type": "day"})
        except:
            resp = _safe_get(url, {"response": "json", "dayDate": "", "weekDate": "", "monthDate": "", "type": "day"})

    result = {
        "foreign_buy": None, "foreign_sell": None, "foreign_net": None,
        "foreign_prev_net": None,
        "trust_buy": None, "trust_sell": None, "trust_net": None,
        "trust_prev_net": None,
        "dealer_net": None,
    }

    if resp is None:
        return result

    try:
        data = resp.json()
        if "data" in data:
            rows = data["data"]
            for row in rows:
                name = str(row[0]).strip()
                buy = _parse_number(row[1])
                sell = _parse_number(row[2])
                net = _parse_number(row[3])

                if ("外資及陸資" in name or ("外資" in name and "自營" not in name)) and result["foreign_net"] is None:
                    result["foreign_buy"] = buy
                    result["foreign_sell"] = sell
                    result["foreign_net"] = net
                elif "投信" in name:
                    result["trust_buy"] = buy
                    result["trust_sell"] = sell
                    result["trust_net"] = net
                elif "自營商" in name:
                    if "合計" in name or result["dealer_net"] is None:
                        result["dealer_net"] = net
                        if buy is not None and sell is not None:
                            result["dealer_buy"] = buy
                            result["dealer_sell"] = sell
    except Exception as e:
        print(f"  [錯誤] 解析三大法人失敗: {e}")

    # ===== 備用1: TWSE OpenAPI v1 (最新資料，不需日期) =====
    if result["foreign_net"] is None:
        print("  [備用] TWSE OpenAPI v1...")
        try:
            oa_url = "https://openapi.twse.com.tw/v1/exchangeReport/BFIAMU"
            resp_oa = _safe_get(oa_url, timeout=10)
            if resp_oa:
                oa_data = resp_oa.json()
                if isinstance(oa_data, list):
                    for item in oa_data:
                        name = item.get("Name", "") or item.get("name", "")
                        buy = _parse_number(str(item.get("BuyAmount", "") or item.get("buy", "")))
                        sell = _parse_number(str(item.get("SellAmount", "") or item.get("sell", "")))
                        net = _parse_number(str(item.get("NetAmount", "") or item.get("net", "")))
                        if not net and buy and sell:
                            net = buy - sell
                        if "外資" in name and result["foreign_net"] is None:
                            result["foreign_buy"] = buy
                            result["foreign_sell"] = sell
                            result["foreign_net"] = net
                        elif "投信" in name:
                            result["trust_buy"] = buy
                            result["trust_sell"] = sell
                            result["trust_net"] = net
                        elif "自營" in name and result["dealer_net"] is None:
                            result["dealer_net"] = net
                    if result["foreign_net"] is not None:
                        print(f"  [TWSE OpenAPI] ✅ 外資={result['foreign_net']}")
        except Exception as e:
            print(f"  [TWSE OpenAPI] ❌ {e}")

    # ===== 備用2: FinMind TaiwanStockTotalInstitutionalInvestors =====
    if result["foreign_net"] is None:
        print("  [備用] FinMind...")
        fm_data = _finmind_fetch("TaiwanStockTotalInstitutionalInvestors", "", date_str, date_str)
        if fm_data:
            try:
                for row in fm_data:
                    name = row.get("name", "")
                    buy = row.get("buy")
                    sell = row.get("sell")
                    if buy is not None and sell is not None:
                        net = int(buy) - int(sell)
                        if "外資" in name and result["foreign_net"] is None:
                            result["foreign_buy"] = int(buy)
                            result["foreign_sell"] = int(sell)
                            result["foreign_net"] = net
                        elif "投信" in name:
                            result["trust_buy"] = int(buy)
                            result["trust_sell"] = int(sell)
                            result["trust_net"] = net
                        elif "自營" in name and result["dealer_net"] is None:
                            result["dealer_net"] = net
                if result["foreign_net"] is not None:
                    print(f"  [FinMind] ✅ 外資={result['foreign_net']}")
            except Exception as e:
                print(f"  [FinMind] ❌ {e}")

    if result["foreign_net"] is None:
        print("  ⚠️ 三大法人: 所有方法均失敗")

    # 抓取前一日的資料來比較
    prev_date = _get_prev_trading_date(date_str)
    if prev_date:
        url_prev = "https://www.twse.com.tw/fund/BFI82U"
        params_prev = {"response": "json", "dayDate": prev_date, "type": "day"}
        resp_prev = _safe_get(url_prev, params_prev)
        if resp_prev:
            try:
                data_prev = resp_prev.json()
                if "data" in data_prev:
                    for row in data_prev["data"]:
                        name = str(row[0]).strip()
                        net = _parse_number(row[3])
                        if "外資及陸資" in name:
                            result["foreign_prev_net"] = net
                        elif "投信" in name:
                            result["trust_prev_net"] = net
            except:
                pass

    return result


# ============================================================
# 3. 外資買超/賣超前十名
# ============================================================
def fetch_foreign_top10(date_str=None):
    """
    抓取外資買超/賣超前十名
    回傳: {top_buy: [{stock_id, stock_name, net_amount}], top_sell: [...]}
    """
    print("🌍 抓取外資買賣超排行...")
    if date_str is None:
        date_str = datetime.now().strftime("%Y%m%d")

    url = "https://www.twse.com.tw/fund/TWT38U"
    params = {"response": "json", "date": date_str}
    resp = _safe_get(url, params)

    result = {"top_buy": [], "top_sell": []}

    if resp is None:
        # 備用方案：使用 T86 取得外資個股買賣
        return _fetch_foreign_top10_from_t86(date_str)

    try:
        data = resp.json()
        if "data" in data and len(data["data"]) > 0:
            for row in data["data"]:
                stock_id = str(row[0]).strip()
                stock_name = str(row[1]).strip().replace("*", "").strip()
                if not stock_name or len(stock_name) < 1:
                    stock_name = stock_id
                raw_buy = _parse_number(row[2])
                raw_sell = _parse_number(row[3])
                raw_net = _parse_number(row[4])
                # TWT38U 的數字是「股數」，轉為「張」(1張=1000股)
                item = {
                    "stock_id": stock_id,
                    "stock_name": stock_name,
                    "buy": round(raw_buy / 1000) if raw_buy else None,
                    "sell": round(raw_sell / 1000) if raw_sell else None,
                    "net": round(raw_net / 1000) if raw_net else None,
                }
                if item["net"] and item["net"] > 0:
                    result["top_buy"].append(item)
                elif item["net"] and item["net"] < 0:
                    result["top_sell"].append(item)

            result["top_buy"] = sorted(result["top_buy"], key=lambda x: x["net"], reverse=True)[:10]
            result["top_sell"] = sorted(result["top_sell"], key=lambda x: x["net"])[:10]
            print(f"  [TWT38U] ✅ 買超 {len(result['top_buy'])} 檔, 賣超 {len(result['top_sell'])} 檔")
        else:
            print("  [TWT38U] 無資料，使用 T86 備用")
            return _fetch_foreign_top10_from_t86(date_str)
    except Exception as e:
        print(f"  [錯誤] 解析外資排行失敗: {e}")
        return _fetch_foreign_top10_from_t86(date_str)

    # 確保所有股票都有公司名稱
    _ensure_stock_names(result["top_buy"])
    _ensure_stock_names(result["top_sell"])
    return result


def _fetch_foreign_top10_from_t86(date_str):
    """備用方案：從 T86 報表取得外資個股買賣"""
    url = "https://www.twse.com.tw/fund/T86"
    params = {"response": "json", "date": date_str, "selectType": "ALLBUT0999"}
    resp = _safe_get(url, params)

    result = {"top_buy": [], "top_sell": []}

    if resp is None:
        return result

    try:
        data = resp.json()
        if "data" in data:
            items = []
            for row in data["data"]:
                net_shares = _parse_number(row[4])  # 買賣超股數
                if net_shares is not None and net_shares != 0:
                    stock_id = str(row[0]).strip()
                    stock_name = str(row[1]).strip()
                    # 清理股票名稱 (移除多餘空白和 *)
                    stock_name = stock_name.replace("*", "").strip()
                    if not stock_name or len(stock_name) < 1:
                        stock_name = stock_id
                    # T86 的數字是股數，轉為張 (1張=1000股)
                    net_zhang = round(net_shares / 1000)
                    items.append({
                        "stock_id": stock_id,
                        "stock_name": stock_name,
                        "buy": round((_parse_number(row[2]) or 0) / 1000),
                        "sell": round((_parse_number(row[3]) or 0) / 1000),
                        "net": net_zhang,
                    })

            buys = sorted([x for x in items if x["net"] > 0], key=lambda x: x["net"], reverse=True)
            sells = sorted([x for x in items if x["net"] < 0], key=lambda x: x["net"])

            result["top_buy"] = buys[:10]
            result["top_sell"] = sells[:10]
    except Exception as e:
        print(f"  [錯誤] 備用方案解析失敗: {e}")

    # 確保所有股票都有公司名稱
    _ensure_stock_names(result["top_buy"])
    _ensure_stock_names(result["top_sell"])
    return result


# ============================================================
# 4. 融資融券
# ============================================================
def fetch_margin_trading(date_str=None):
    """
    抓取融資融券餘額
    策略: FinMind 為主要 → TWSE 新API → TWSE 舊API
    回傳: {margin_balance, margin_change, margin_balance_amount, short_balance, short_change}
    """
    print("💰 抓取融資融券...")
    if date_str is None:
        date_str = datetime.now().strftime("%Y%m%d")

    result = {
        "margin_balance": None, "margin_change": None,
        "margin_balance_amount": None,
        "short_balance": None, "short_change": None,
    }

    # ===== 方法1 (主要): FinMind TaiwanStockTotalMarginPurchaseShortSale =====
    # FinMind 資料表定義:
    #   name: "融資(交易單位-Loss)" / "融券(交易單位-Loss)" / "融資金額(仟元)"
    #   TodayBalance: 今日餘額
    #   YesBalance: 昨日餘額
    #   buy: 買進, sell: 賣出, Return: 現金(券)償還
    print("  [主要] FinMind TaiwanStockTotalMarginPurchaseShortSale...")
    fm_data = _finmind_fetch("TaiwanStockTotalMarginPurchaseShortSale", "", date_str, date_str)

    if not fm_data:
        print("  [FinMind] ❌ 今日融資融券無資料")

    if fm_data:
        try:
            for row in fm_data:
                name = row.get("name", "")
                today_bal = row.get("TodayBalance")
                yes_bal = row.get("YesBalance")
                print(f"    [FinMind] name={name}, TodayBalance={today_bal}, YesBalance={yes_bal}")

                if "融資" in name and "金額" not in name:
                    result["margin_balance"] = _parse_number(str(today_bal)) if today_bal is not None else None
                    if today_bal is not None and yes_bal is not None:
                        result["margin_change"] = int(today_bal) - int(yes_bal)
                elif "融資金額" in name or ("融資" in name and "金額" in name):
                    amt = _parse_number(str(today_bal)) if today_bal is not None else None
                    if amt is not None:
                        result["margin_balance_amount"] = amt * 1000  # 仟元轉元
                elif "融券" in name and "金額" not in name:
                    result["short_balance"] = _parse_number(str(today_bal)) if today_bal is not None else None
                    if today_bal is not None and yes_bal is not None:
                        result["short_change"] = int(today_bal) - int(yes_bal)

            print(f"    [FinMind] 結果: 融資={result['margin_balance']}, 金額={result['margin_balance_amount']}, 融券={result['short_balance']}")
        except Exception as e:
            print(f"  [錯誤] FinMind 解析失敗: {e}")
            import traceback
            traceback.print_exc()

    # ===== 方法2 (備用): TWSE MI_MARGN =====
    if result["margin_balance"] is None:
        print("  [備用] 嘗試 TWSE MI_MARGN...")
        url = "https://www.twse.com.tw/exchangeReport/MI_MARGN"
        params = {"response": "json", "date": date_str, "selectType": "MS"}
        resp = _safe_get(url, params)
        if resp:
            try:
                data = resp.json()
                credit_list = data.get("creditList", [])
                credit_fields = data.get("creditFields", [])
                if credit_fields and credit_list:
                    # 找「今日餘額」欄位
                    balance_idx = None
                    for i, fn in enumerate(credit_fields):
                        if "今日" in str(fn) and "餘額" in str(fn):
                            balance_idx = i
                    if balance_idx is None and len(credit_fields) >= 2:
                        balance_idx = len(credit_fields) - 1

                    for row in credit_list:
                        if not row or len(row) <= (balance_idx or 0):
                            continue
                        item_name = str(row[0]).strip()
                        if "融資" in item_name and "金額" not in item_name and result["margin_balance"] is None:
                            result["margin_balance"] = _parse_number(row[balance_idx]) if balance_idx is not None else None
                        elif "融券" in item_name and "金額" not in item_name and result["short_balance"] is None:
                            result["short_balance"] = _parse_number(row[balance_idx]) if balance_idx is not None else None
                        elif "融資金額" in item_name or ("融資" in item_name and "金額" in item_name):
                            amt = _parse_number(row[balance_idx]) if balance_idx is not None else None
                            if amt is not None:
                                result["margin_balance_amount"] = amt * 1000
                    print(f"    [TWSE] 融資={result['margin_balance']}, 融券={result['short_balance']}")
            except Exception as e:
                print(f"  [錯誤] TWSE MI_MARGN 解析: {e}")

    # ===== 方法3 (備用): TWSE MI_MARGN selectType=ALL 個股加總 =====
    if result["margin_balance"] is None:
        print("  [備用3] 嘗試 TWSE MI_MARGN ALL (個股加總)...")
        url3 = "https://www.twse.com.tw/exchangeReport/MI_MARGN"
        params3 = {"response": "json", "date": date_str, "selectType": "ALL"}
        resp3 = _safe_get(url3, params3)
        if resp3:
            try:
                data3 = resp3.json()
                if "data" in data3 and len(data3["data"]) > 0:
                    total_margin = 0
                    total_short = 0
                    count = 0
                    for row in data3["data"]:
                        if len(row) >= 7:
                            # 格式: 代號, 名稱, 融資買進, 融資賣出, 融資現償, 融資餘額, ...融券
                            m_bal = _parse_number(row[6]) if len(row) > 6 else None  # 融資今日餘額
                            s_bal = _parse_number(row[12]) if len(row) > 12 else None  # 融券今日餘額
                            if m_bal is not None:
                                total_margin += m_bal
                                count += 1
                            if s_bal is not None:
                                total_short += s_bal
                    if count > 0:
                        result["margin_balance"] = total_margin
                        result["short_balance"] = total_short
                        print(f"    [TWSE ALL] ✅ 加總 {count} 檔: 融資={total_margin}, 融券={total_short}")
            except Exception as e:
                print(f"  [錯誤] TWSE ALL 解析: {e}")

    # ===== 方法4 (備用): TWSE OpenAPI v1 TWT93U (融資融券彙總) =====
    if result["margin_balance"] is None:
        print("  [備用4] 嘗試 TWSE OpenAPI TWT93U...")
        try:
            oa_url = "https://openapi.twse.com.tw/v1/exchangeReport/TWT93U"
            resp_oa = _safe_get(oa_url, timeout=12)
            if resp_oa:
                oa_data = resp_oa.json()
                if isinstance(oa_data, list) and len(oa_data) > 0:
                    print(f"    [TWT93U] 共 {len(oa_data)} 筆, keys={list(oa_data[0].keys())}")
                    for item in oa_data:
                        name = item.get("Name", "") or item.get("Title", "") or item.get("ItemName", "")
                        # 嘗試找各種可能的值欄位
                        val_str = (item.get("TodayBalance", "") or item.get("Balance", "") or
                                   item.get("Unit", "") or item.get("Value", ""))
                        if not name:
                            # 如果沒有 Name 欄位，嘗試用第一個字串欄位
                            for k, v in item.items():
                                if isinstance(v, str) and ("融資" in v or "融券" in v):
                                    name = v
                                    break
                        val = _parse_number(str(val_str)) if val_str else None
                        if not val:
                            # 嘗試其他數值欄位
                            for k, v in item.items():
                                if k not in ["Name", "Title", "ItemName", "Date", "date"]:
                                    parsed = _parse_number(str(v))
                                    if parsed and parsed > 10000:
                                        val = parsed
                                        break

                        print(f"    [TWT93U] name={name}, val={val}")
                        if "融資" in str(name) and "金額" not in str(name) and val:
                            result["margin_balance"] = val
                        elif "融券" in str(name) and "金額" not in str(name) and val:
                            result["short_balance"] = val
                        elif ("融資金額" in str(name) or ("融資" in str(name) and "金額" in str(name))) and val:
                            result["margin_balance_amount"] = val * 1000 if val < 1e9 else val
                    if result["margin_balance"]:
                        print(f"    [TWT93U] ✅ 融資={result['margin_balance']}, 融券={result['short_balance']}")
        except Exception as e:
            print(f"    [TWT93U] ❌ {e}")
            import traceback
            traceback.print_exc()

    # ===== 方法5 (備用): 鉅亨網 API =====
    if result["margin_balance"] is None:
        print("  [備用5] 嘗試鉅亨網...")
        try:
            url4 = "https://ws.api.cnyes.com/ws/api/v1/charting/history"
            for symbol, key in [
                ("TWS:MARGIN_TRADING_BALANCE:STOCK", "margin_balance"),
                ("TWS:SHORT_SELLING_BALANCE:STOCK", "short_balance"),
            ]:
                params4 = {"resolution": "D", "symbol": symbol, "quote": 1}
                resp4 = _safe_get(url4, params4, timeout=8)
                if resp4:
                    data4 = resp4.json()
                    if "data" in data4 and "quote" in data4["data"]:
                        val = data4["data"]["quote"].get("6") or data4["data"]["quote"].get("closePrice")
                        if val:
                            result[key] = float(val)
                            print(f"    [鉅亨] ✅ {key}={val}")
        except Exception as e:
            print(f"    [鉅亨] ❌ {e}")

    if result["margin_balance"] is None:
        print("  ⚠️ 融資融券: 所有方法均失敗")

    return result


# ============================================================
# 5. 上漲/下跌家數
# ============================================================
def fetch_market_breadth(date_str=None):
    """
    抓取上市上櫃漲跌家數
    回傳: {tse_up, tse_down, tse_flat, otc_up, otc_down, otc_flat}
    """
    print("📈 抓取漲跌家數...")
    if date_str is None:
        date_str = datetime.now().strftime("%Y%m%d")

    result = {
        "tse_up": None, "tse_down": None, "tse_flat": None,
        "otc_up": None, "otc_down": None, "otc_flat": None,
    }

    # ==== 方法1 (主要): TWSE MI_INDEX 每日收盤行情 ====
    print("  [方法1] TWSE MI_INDEX...")
    url2 = "https://www.twse.com.tw/exchangeReport/MI_INDEX"
    params2 = {"response": "json", "date": date_str, "type": "ALLBUT0999"}
    resp2 = _safe_get(url2, params2)

    if resp2:
        try:
            data2 = resp2.json()
            if data2.get("stat") == "OK":
                up_count = 0
                down_count = 0
                flat_count = 0

                # 遍歷所有 data 欄位找個股資料
                for key in sorted(data2.keys()):
                    if not key.startswith("data") or not isinstance(data2[key], list):
                        continue
                    rows = data2[key]
                    if len(rows) < 10:
                        continue

                    # 先找出哪個欄位是漲跌方向 (+ / - / X)
                    sign_col = None
                    if rows:
                        sample = rows[0]
                        if isinstance(sample, list):
                            for ci in range(len(sample)):
                                if str(sample[ci]).strip() in ["+", "-", "X", " "]:
                                    sign_col = ci
                                    break

                    for row in rows:
                        if not isinstance(row, list) or len(row) < 5:
                            continue
                        try:
                            if sign_col is not None and sign_col < len(row):
                                val_str = str(row[sign_col]).strip()
                                if val_str == "+":
                                    up_count += 1
                                elif val_str == "-":
                                    down_count += 1
                                else:
                                    flat_count += 1
                            else:
                                # 嘗試多個可能的欄位位置
                                found = False
                                for ci in [9, 8, 10, 7]:
                                    if ci < len(row):
                                        val_str = str(row[ci]).strip()
                                        if val_str in ["+", "-", "X", " "]:
                                            if val_str == "+":
                                                up_count += 1
                                            elif val_str == "-":
                                                down_count += 1
                                            else:
                                                flat_count += 1
                                            found = True
                                            break
                                if not found:
                                    # 嘗試從漲跌價差欄位判斷
                                    for ci in [10, 9, 8]:
                                        if ci < len(row):
                                            try:
                                                val = float(str(row[ci]).replace(",", ""))
                                                if val > 0:
                                                    up_count += 1
                                                elif val < 0:
                                                    down_count += 1
                                                else:
                                                    flat_count += 1
                                                break
                                            except:
                                                continue
                        except:
                            pass

                    if up_count > 0:
                        result["tse_up"] = up_count
                        result["tse_down"] = down_count
                        result["tse_flat"] = flat_count
                        print(f"    [MI_INDEX] ✅ 漲{up_count} 跌{down_count} 平{flat_count}")
                        break
            else:
                print(f"    [MI_INDEX] stat={data2.get('stat')}")
        except Exception as e:
            print(f"    [MI_INDEX] ❌ {e}")

    # 上櫃 - 從 TPEx 取得
    tpex_date = f"{int(date_str[:4]) - 1911}/{date_str[4:6]}/{date_str[6:8]}"
    url_otc = "https://www.tpex.org.tw/web/stock/aftertrading/market_highlight/highlight_result.php"
    params_otc = {"l": "zh-tw", "d": tpex_date, "o": "json"}
    resp_otc = _safe_get(url_otc, params_otc)

    if resp_otc:
        try:
            data_otc = resp_otc.json()
            if "tables" in data_otc:
                for table in data_otc["tables"]:
                    if "data" in table:
                        for row in table["data"]:
                            row_text = str(row)
                            if "上漲" in row_text:
                                result["otc_up"] = _parse_number(row[1]) if len(row) > 1 else None
                            elif "下跌" in row_text:
                                result["otc_down"] = _parse_number(row[1]) if len(row) > 1 else None
                            elif "持平" in row_text:
                                result["otc_flat"] = _parse_number(row[1]) if len(row) > 1 else None
        except:
            pass

    # 備用: 從 TPEx 個股行情統計
    if result["otc_up"] is None:
        url_otc2 = "https://www.tpex.org.tw/web/stock/aftertrading/daily_trading_index/st41_result.php"
        params_otc2 = {"l": "zh-tw", "d": tpex_date, "o": "json"}
        resp_otc2 = _safe_get(url_otc2, params_otc2)
        if resp_otc2:
            try:
                data_otc2 = resp_otc2.json()
                if "iTotalRecords" in data_otc2 and "aaData" in data_otc2:
                    rows = data_otc2["aaData"]
                    if rows:
                        latest = rows[-1]
                        # 上漲, 下跌, 持平 家數
                        if len(latest) > 8:
                            result["otc_up"] = _parse_number(latest[6])
                            result["otc_down"] = _parse_number(latest[7])
                            result["otc_flat"] = _parse_number(latest[8])
            except:
                pass

    # ==== 方法3 (備用): TWSE OpenAPI v1 STOCK_DAY_ALL 自己統計漲跌 ====
    if result["tse_up"] is None:
        print("  [方法3] 嘗試 TWSE OpenAPI v1 STOCK_DAY_ALL 統計漲跌...")
        try:
            oa_url = "https://openapi.twse.com.tw/v1/exchangeReport/STOCK_DAY_ALL"
            resp_oa = _safe_get(oa_url, timeout=15)
            if resp_oa:
                oa_data = resp_oa.json()
                if isinstance(oa_data, list) and len(oa_data) > 50:
                    # 動態找到「漲跌」欄位 - 嘗試多種可能的 key
                    change_key = None
                    first = oa_data[0]
                    for candidate in ["Change", "change", "UpDown", "漲跌價差", "漲跌"]:
                        if candidate in first:
                            change_key = candidate
                            break
                    # 如果沒找到，搜尋所有 key
                    if not change_key:
                        for k in first.keys():
                            if "change" in k.lower() or "漲跌" in k or "updown" in k.lower():
                                change_key = k
                                break
                    print(f"    [STOCK_DAY_ALL] keys={list(first.keys())[:8]}, change_key={change_key}")

                    if change_key:
                        up = down = flat = 0
                        for item in oa_data:
                            val = _parse_number(str(item.get(change_key, "")))
                            if val is not None:
                                if val > 0:
                                    up += 1
                                elif val < 0:
                                    down += 1
                                else:
                                    flat += 1
                        if up > 0 or down > 0:
                            result["tse_up"] = up
                            result["tse_down"] = down
                            result["tse_flat"] = flat
                            print(f"    [STOCK_DAY_ALL] ✅ 漲{up} 跌{down} 平{flat}")
                    else:
                        # 最後手段: 用收盤價和昨日收盤價比較
                        close_key = None
                        prev_key = None
                        for k in first.keys():
                            kl = k.lower()
                            if "close" in kl or "收盤" in k:
                                if not close_key:
                                    close_key = k
                            if "yesterday" in kl or "昨日" in k or "prev" in kl:
                                prev_key = k
                        print(f"    [STOCK_DAY_ALL] close_key={close_key}, prev_key={prev_key}")
                        if close_key:
                            up = down = flat = 0
                            for item in oa_data:
                                c = _parse_number(str(item.get(close_key, "")))
                                p = _parse_number(str(item.get(prev_key or "", "")))
                                if c is not None and p is not None and p > 0:
                                    diff = c - p
                                    if diff > 0:
                                        up += 1
                                    elif diff < 0:
                                        down += 1
                                    else:
                                        flat += 1
                            if up > 0 or down > 0:
                                result["tse_up"] = up
                                result["tse_down"] = down
                                result["tse_flat"] = flat
                                print(f"    [STOCK_DAY_ALL close] ✅ 漲{up} 跌{down} 平{flat}")
        except Exception as e:
            print(f"    [STOCK_DAY_ALL] ❌ {e}")
            import traceback
            traceback.print_exc()

    return result


# ============================================================
# 6. 美元指數、日圓、VIX (鉅亨網 + investing.com)
# ============================================================
def _fetch_cnyes_chart(cnyes_symbol, days=30):
    """
    從鉅亨網 API 抓取歷史價格
    cnyes_symbol: 鉅亨網代號, 如 "FX:USDTWD", "GI:DXY", "GI:VIX", "GI:US10Y"
    回傳: [{"date": "2024-01-01", "close": 104.5}, ...]
    """
    print(f"  [鉅亨] 取得 {cnyes_symbol} (近{days}天)...")
    end_ts = int(time.time())
    start_ts = end_ts - days * 86400

    # 方法1: charting/history
    url = "https://ws.api.cnyes.com/ws/api/v1/charting/history"
    params = {
        "resolution": "D",
        "symbol": cnyes_symbol,
        "from": str(start_ts),
        "to": str(end_ts),
    }
    resp = _safe_get(url, params, timeout=10)
    if resp:
        try:
            data = resp.json()
            if data.get("statusCode") == 200 or "data" in data:
                chart_data = data.get("data", {})
                timestamps = chart_data.get("t", [])
                closes = chart_data.get("c", [])
                if timestamps and closes:
                    points = []
                    for ts, c in zip(timestamps, closes):
                        if c is not None:
                            dt = datetime.fromtimestamp(ts)
                            points.append({"date": dt.strftime("%Y-%m-%d"), "close": round(float(c), 2)})
                    if points:
                        print(f"  [鉅亨] ✅ {cnyes_symbol} 共 {len(points)} 筆")
                        return points
        except Exception as e:
            print(f"  [鉅亨] ❌ {cnyes_symbol} 解析失敗: {e}")

    # 方法2: quote API (只取最新價)
    url2 = "https://ws.api.cnyes.com/ws/api/v1/charting/history"
    params2 = {"resolution": "D", "symbol": cnyes_symbol, "quote": 1}
    resp2 = _safe_get(url2, params2, timeout=10)
    if resp2:
        try:
            data2 = resp2.json()
            quote = data2.get("data", {}).get("quote", {})
            close_val = quote.get("6") or quote.get("closePrice") or quote.get("priceClose")
            if close_val:
                today_str = datetime.now().strftime("%Y-%m-%d")
                print(f"  [鉅亨 quote] ✅ {cnyes_symbol} = {close_val}")
                return [{"date": today_str, "close": round(float(close_val), 2)}]
        except Exception as e:
            print(f"  [鉅亨 quote] ❌ {cnyes_symbol}: {e}")

    print(f"  [鉅亨] ⚠️ {cnyes_symbol} 無資料")
    return []


def _fetch_investing_data(pair_id, days=30):
    """
    從 investing.com API 取得歷史資料
    pair_id: investing.com 的 pair ID
    """
    print(f"  [Investing] 取得 pair_id={pair_id}...")
    url = f"https://api.investing.com/api/financialdata/{pair_id}/historical/chart/"
    params = {
        "period": "P1M",
        "interval": "P1D",
        "pointscount": days,
    }
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
        "domain-id": "www.investing.com",
        "Referer": "https://www.investing.com/",
    }
    try:
        resp = requests.get(url, params=params, headers=headers, timeout=10)
        if resp.status_code == 200:
            data = resp.json()
            if "data" in data:
                points = []
                for item in data["data"]:
                    ts = item.get("date_utc") or item.get("rowDateTimestamp")
                    close = item.get("last_close") or item.get("last_numeric")
                    if ts and close:
                        dt = datetime.fromtimestamp(ts / 1000) if ts > 1e10 else datetime.fromtimestamp(ts)
                        points.append({"date": dt.strftime("%Y-%m-%d"), "close": round(float(close), 2)})
                if points:
                    print(f"  [Investing] ✅ pair_id={pair_id} 共 {len(points)} 筆")
                    return points
    except Exception as e:
        print(f"  [Investing] ❌ pair_id={pair_id}: {e}")
    return []


def fetch_usd_index():
    """
    抓取美元指數近30天
    嘗試多個鉅亨網代號 + investing.com
    """
    print("💵 抓取美元指數...")
    # 方法1: 鉅亨網 - 嘗試多個可能的 symbol
    for sym in ["GI:DXY", "GI:DXY00", "TWG:DXY00", "FX:DXY", "FX:USDX"]:
        data = _fetch_cnyes_chart(sym, 35)
        if data:
            return data
    # 方法2: investing.com (DXY pair_id=8827)
    data = _fetch_investing_data(8827, 35)
    if data:
        return data
    # 方法3: 從 cnyes 搜尋正確的 DXY symbol
    print("  [備用] 嘗試 cnyes search API 找 DXY symbol...")
    try:
        search_url = "https://ws.api.cnyes.com/ws/api/v1/universal/search"
        resp = _safe_get(search_url, {"q": "DXY"}, timeout=8)
        if resp:
            sdata = resp.json()
            items = sdata.get("data", {}).get("items", [])
            for item in items:
                sym = item.get("symbol", "")
                if sym and ("DXY" in sym.upper() or "美元" in item.get("name", "")):
                    print(f"  [cnyes search] 找到 symbol: {sym} ({item.get('name','')})")
                    data = _fetch_cnyes_chart(sym, 35)
                    if data:
                        return data
    except Exception as e:
        print(f"  [cnyes search] ❌ {e}")
    print("  ⚠️ 美元指數: 所有來源失敗")
    return []


def fetch_jpy_rate():
    """
    抓取日圓匯率近30天 (USD/JPY)
    鉅亨網代號: FX:USDJPY
    """
    print("💴 抓取日圓匯率...")
    # 方法1: 鉅亨網
    data = _fetch_cnyes_chart("FX:USDJPY", 35)
    if data:
        return data
    # 方法2: investing.com (USDJPY pair_id=3)
    data = _fetch_investing_data(3, 35)
    if data:
        return data
    print("  ⚠️ 日圓匯率: 所有來源失敗")
    return []


def fetch_vix():
    """
    抓取 VIX 指數 (近7天含圖表資料)
    鉅亨網代號: GI:VIX
    """
    print("😱 抓取 VIX 指數 (近7天)...")
    # 方法1: 鉅亨網
    data = _fetch_cnyes_chart("GI:VIX", 10)
    # 方法2: investing.com (VIX pair_id=44336)
    if not data:
        data = _fetch_investing_data(44336, 10)
    if not data:
        print("  ⚠️ VIX: 所有來源失敗")

    data = data[-7:] if len(data) > 7 else data
    if data:
        return {
            "value": data[-1]["close"],
            "date": data[-1]["date"],
            "prev_value": data[-2]["close"] if len(data) >= 2 else None,
            "chart": data,
        }
    return {"value": None, "date": None, "prev_value": None, "chart": []}


# ============================================================
# 7. 台指期貨 (TAIFEX)
# ============================================================
def fetch_taiex_futures(date_str=None):
    """
    抓取台指期貨 (近月) 收盤資料
    資料來源: 台灣期貨交易所 (TAIFEX)
    回傳: {close, change, change_pct, volume, settlement, open, high, low}
    """
    print("📈 抓取台指期貨...")
    if date_str is None:
        date_str = datetime.now().strftime("%Y%m%d")

    # 期交所 API 日期格式: YYYY/MM/DD
    formatted_date = f"{date_str[:4]}/{date_str[4:6]}/{date_str[6:8]}"

    result = {
        "close": None, "change": None, "change_pct": None,
        "volume": None, "settlement": None,
        "open": None, "high": None, "low": None,
        "contract_month": None,
    }

    # 優先使用 CSV 下載 (更穩定)
    result = _fetch_futures_backup(date_str)
    if result["close"] is not None:
        return result

    # 備用: HTML 解析
    url = "https://www.taifex.com.tw/cht/3/futContractsDate"
    params = {
        "queryType": "1",
        "marketCode": "0",
        "dateaddcnt": "",
        "commodity_id": "TX",
        "queryDate": formatted_date,
    }
    resp = _safe_get(url, params)

    if resp is None:
        return result

    try:
        from bs4 import BeautifulSoup
        soup = BeautifulSoup(resp.text, "html.parser")
        tables = soup.find_all("table")
        for table in tables:
            rows = table.find_all("tr")
            for row in rows:
                cells = row.find_all("td")
                if len(cells) >= 8:
                    cell_text = [c.get_text(strip=True) for c in cells]
                    row_text = " ".join(cell_text)
                    # 找近月合約
                    if ("TX" in row_text or "臺股期貨" in row_text) and result["close"] is None:
                        for i, t in enumerate(cell_text):
                            val = _parse_number(t)
                            if val and 10000 < val < 50000 and result["close"] is None:
                                # 可能是收盤價 (介於10000-50000)
                                result["open"] = _parse_number(cell_text[max(0,i-3)]) if i >= 3 else None
                                result["high"] = _parse_number(cell_text[max(0,i-2)]) if i >= 2 else None
                                result["low"] = _parse_number(cell_text[max(0,i-1)]) if i >= 1 else None
                                result["close"] = val
                                if i+1 < len(cell_text):
                                    result["change"] = _parse_number(cell_text[i+1])
                                if result["close"] and result["change"]:
                                    prev = result["close"] - result["change"]
                                    if prev != 0:
                                        result["change_pct"] = round(result["change"] / prev * 100, 2)
                                break
                        if result["close"]:
                            break
    except Exception as e:
        print(f"  [錯誤] 解析期貨資料失敗: {e}")

    return result


def _fetch_futures_backup(date_str):
    """備用方案：從期交所 JSON API 取得期貨資料"""
    formatted_date = f"{date_str[:4]}/{date_str[4:6]}/{date_str[6:8]}"
    url = "https://www.taifex.com.tw/cht/3/futDataDown"
    params = {
        "down_type": "1",
        "commodity_id": "TX",
        "queryStartDate": formatted_date,
        "queryEndDate": formatted_date,
    }
    resp = _safe_get(url, params)
    result = {
        "close": None, "change": None, "change_pct": None,
        "volume": None, "settlement": None,
        "open": None, "high": None, "low": None,
        "contract_month": None,
    }

    if resp is None:
        return result

    try:
        lines = resp.text.strip().split("\n")
        if len(lines) > 1:
            # CSV 格式，找第一筆 TX 近月資料 (非 week/月合約)
            found = False
            for line in lines[1:]:
                fields = [f.strip().strip('"') for f in line.split(",")]
                if len(fields) < 8:
                    continue
                # 契約代碼在前幾個欄位
                line_text = ",".join(fields[:3])
                if "TX" not in line_text:
                    continue
                # 跳過小台(MTX)、微台(MXF)、電子期(TE)、金融期(TF)
                if any(x in line_text for x in ["MTX", "MXF", "TE", "TF"]):
                    continue

                # 找到數值欄位 (收盤價在10000-50000之間)
                for i in range(2, min(len(fields), 10)):
                    val = _parse_number(fields[i])
                    if val and 10000 < val < 50000:
                        # 這可能是開盤價，往後找
                        result["open"] = val
                        result["high"] = _parse_number(fields[i+1]) if i+1 < len(fields) else None
                        result["low"] = _parse_number(fields[i+2]) if i+2 < len(fields) else None
                        result["close"] = _parse_number(fields[i+3]) if i+3 < len(fields) else None
                        # 如果 close 不在合理範圍，嘗試其他排列
                        if result["close"] is None or result["close"] < 10000:
                            result["open"] = _parse_number(fields[3]) if len(fields) > 3 else None
                            result["high"] = _parse_number(fields[4]) if len(fields) > 4 else None
                            result["low"] = _parse_number(fields[5]) if len(fields) > 5 else None
                            result["close"] = _parse_number(fields[6]) if len(fields) > 6 else None
                        found = True
                        break

                if not found:
                    # 嘗試固定欄位位置
                    result["open"] = _parse_number(fields[3]) if len(fields) > 3 else None
                    result["high"] = _parse_number(fields[4]) if len(fields) > 4 else None
                    result["low"] = _parse_number(fields[5]) if len(fields) > 5 else None
                    result["close"] = _parse_number(fields[6]) if len(fields) > 6 else None
                    found = True

                if found and result["close"]:
                    result["contract_month"] = fields[2] if len(fields) > 2 else None
                    result["change"] = _parse_number(fields[7]) if len(fields) > 7 else None
                    result["volume"] = _parse_number(fields[8]) if len(fields) > 8 else None
                    result["settlement"] = _parse_number(fields[9]) if len(fields) > 9 else None
                    if result["close"] and result["change"]:
                        prev = result["close"] - result["change"]
                        if prev != 0:
                            result["change_pct"] = round(result["change"] / prev * 100, 2)
                    break
    except Exception as e:
        print(f"  [錯誤] 備用期貨解析失敗: {e}")

    return result


# ============================================================
# 7b. 三大法人台指期未平倉 & 期權觀測指標
# ============================================================
def fetch_futures_oi(date_str=None):
    """
    抓取三大法人台指期貨未平倉口數
    來源: 期交所
    回傳: {
        foreign: {change, oi},
        trust: {change, oi},
        dealer: {change, oi},
        total: {change, oi},
    }
    """
    print("📋 抓取三大法人期貨未平倉...")
    if date_str is None:
        date_str = datetime.now().strftime("%Y%m%d")

    formatted_date = f"{date_str[:4]}/{date_str[4:6]}/{date_str[6:8]}"

    result = {
        "foreign": {"change": None, "oi": None},
        "trust": {"change": None, "oi": None},
        "dealer": {"change": None, "oi": None},
        "total": {"change": None, "oi": None},
    }

    # 期交所三大法人-區分各期貨契約
    url = "https://www.taifex.com.tw/cht/3/futContractsDateDown"
    params = {
        "queryStartDate": formatted_date,
        "queryEndDate": formatted_date,
        "commodityId": "TXF",
    }
    resp = _safe_get(url, params)

    if resp:
        try:
            lines = resp.text.strip().split("\n")
            print(f"    [期交所] 解析 {len(lines)} 行資料")
            if len(lines) > 0:
                print(f"    [期交所] 標頭: {lines[0][:200]}")

            # 動態偵測欄位位置：找身份別、多空交易口數淨額、多空未平倉口數淨額
            header = lines[0] if lines else ""
            header_fields = [f.strip().strip('"') for f in header.split(",")]

            # 預設欄位位置 (期交所 CSV 格式: 日期,商品名稱,身份別,...)
            # 0:日期 1:商品名稱 2:身份別 3:多方交易口數 4:多方交易金額
            # 5:空方交易口數 6:空方交易金額 7:多空交易口數淨額 8:多空交易金額淨額
            # 9:多方未平倉口數 10:多方未平倉金額 11:空方未平倉口數 12:空方未平倉金額
            # 13:多空未平倉口數淨額 14:多空未平倉金額淨額
            id_idx = 2
            change_idx = 7   # 多空交易口數淨額
            oi_idx = 13      # 多空未平倉口數淨額

            # 嘗試從標頭動態找欄位
            for i, h in enumerate(header_fields):
                if "身份" in h or "身分" in h:
                    id_idx = i
                elif "多空" in h and "交易" in h and "口數" in h and "淨" in h:
                    change_idx = i
                elif "多空" in h and "未平倉" in h and "口數" in h and "淨" in h:
                    oi_idx = i

            print(f"    [期交所] 欄位定位: id={id_idx}, change={change_idx}, oi={oi_idx}")

            for line in lines[1:]:
                fields = [f.strip().strip('"') for f in line.split(",")]
                if len(fields) <= max(id_idx, change_idx, oi_idx):
                    continue
                identity = fields[id_idx].strip()

                net_change = _parse_number(fields[change_idx])
                net_oi = _parse_number(fields[oi_idx])

                # 轉為整數
                if net_oi is not None:
                    net_oi = int(net_oi)
                if net_change is not None:
                    net_change = int(net_change)

                print(f"    [期交所] {identity}: change={net_change}, oi={net_oi}")

                if "外資" in identity:
                    result["foreign"]["oi"] = net_oi
                    result["foreign"]["change"] = net_change
                elif "投信" in identity:
                    result["trust"]["oi"] = net_oi
                    result["trust"]["change"] = net_change
                elif "自營" in identity:
                    result["dealer"]["oi"] = net_oi
                    result["dealer"]["change"] = net_change
        except Exception as e:
            print(f"  [錯誤] 解析期交所 CSV 失敗: {e}")
            import traceback
            traceback.print_exc()

    # 嘗試用 HTML 表格備用
    if result["foreign"]["oi"] is None:
        url2 = "https://www.taifex.com.tw/cht/3/futContractsDate"
        params2 = {
            "queryType": "1",
            "marketCode": "0",
            "commodity_id": "TX",
            "queryDate": formatted_date,
        }
        resp2 = _safe_get(url2, params2)
        if resp2:
            try:
                from bs4 import BeautifulSoup
                soup = BeautifulSoup(resp2.text, "html.parser")
                # 找三大法人未平倉的表格
                tables = soup.find_all("table")
                for table in tables:
                    text = table.get_text()
                    if "未平倉" in text and ("外資" in text or "投信" in text):
                        rows = table.find_all("tr")
                        for row in rows:
                            cells = row.find_all("td")
                            if len(cells) >= 4:
                                cell_text = [c.get_text(strip=True) for c in cells]
                                identity = cell_text[0]
                                if "外資" in identity:
                                    result["foreign"]["change"] = _parse_number(cell_text[1])
                                    result["foreign"]["oi"] = _parse_number(cell_text[2])
                                elif "投信" in identity:
                                    result["trust"]["change"] = _parse_number(cell_text[1])
                                    result["trust"]["oi"] = _parse_number(cell_text[2])
                                elif "自營" in identity:
                                    result["dealer"]["change"] = _parse_number(cell_text[1])
                                    result["dealer"]["oi"] = _parse_number(cell_text[2])
            except:
                pass

    # 計算合計
    changes = [result[k]["change"] for k in ["foreign", "trust", "dealer"] if result[k]["change"] is not None]
    ois = [result[k]["oi"] for k in ["foreign", "trust", "dealer"] if result[k]["oi"] is not None]
    if changes:
        result["total"]["change"] = sum(changes)
    if ois:
        result["total"]["oi"] = sum(ois)

    # 如果仍無資料，嘗試 FinMind 備用資料源
    if result["foreign"]["oi"] is None:
        print("  [備用] 嘗試 FinMind API...")
        fm_data = _finmind_fetch("TaiwanFuturesInstitutionalInvestors", "TX", date_str, date_str)
        if fm_data:
            try:
                for row in fm_data:
                    name = row.get("name", "")
                    short_oi_vol = _parse_number(row.get("short_open_interest_balance_volume", ""))
                    long_oi_vol = _parse_number(row.get("long_open_interest_balance_volume", ""))

                    if name and short_oi_vol is not None and long_oi_vol is not None:
                        net_oi = int(long_oi_vol - short_oi_vol)
                        print(f"    [FinMind] {name}: oi={net_oi}")

                        if "外資" in name:
                            result["foreign"]["oi"] = net_oi
                        elif "投信" in name:
                            result["trust"]["oi"] = net_oi
                        elif "自營" in name:
                            result["dealer"]["oi"] = net_oi

                # 重新計算合計
                ois = [result[k]["oi"] for k in ["foreign", "trust", "dealer"] if result[k]["oi"] is not None]
                if ois:
                    result["total"]["oi"] = sum(ois)
            except Exception as e:
                print(f"  [錯誤] FinMind 解析失敗: {e}")

    return result


def _calc_sentiment_from_finmind(fm_data):
    """從 FinMind 期貨法人資料計算多空指標"""
    if not fm_data:
        return None
    total_long = 0
    total_short = 0
    for row in fm_data:
        long_vol = row.get("long_open_interest_balance_volume")
        short_vol = row.get("short_open_interest_balance_volume")
        investor = row.get("institutional_investors", "")
        if long_vol is not None and short_vol is not None:
            total_long += int(long_vol)
            total_short += int(short_vol)
            print(f"      {investor}: long={long_vol}, short={short_vol}")
    total = total_long + total_short
    if total > 0:
        return round((total_long - total_short) / total * 100, 2)
    return None


def _calc_sentiment_from_taifex(resp_text, contract_id):
    """從期交所 CSV 或 HTML 資料計算多空指標"""
    if not resp_text:
        return None
    try:
        text = resp_text.strip()
        first_line = text.split("\n")[0].lower()

        # ===== 如果是 HTML → 用 BeautifulSoup 解析表格 =====
        if "<html" in first_line or "<!doctype" in first_line or "<table" in first_line:
            print(f"    [TAIFEX] {contract_id}: 回傳 HTML，改用表格解析...")
            return _calc_sentiment_from_taifex_html(resp_text, contract_id)

        # ===== CSV 解析 =====
        lines = text.split("\n")
        print(f"    [TAIFEX] {contract_id}: 共 {len(lines)} 行 CSV")
        if len(lines) > 0:
            print(f"    [TAIFEX] 標頭: {lines[0][:300]}")

        # 動態偵測欄位位置
        header_fields = [f.strip().strip('"') for f in lines[0].split(",")]
        long_oi_idx = 9   # 多方未平倉口數 (預設)
        short_oi_idx = 11  # 空方未平倉口數 (預設)
        id_idx = 2         # 身份別 (預設)

        for i, h in enumerate(header_fields):
            if "身份" in h or "身分" in h:
                id_idx = i
            elif "多方" in h and "未平倉" in h and "口數" in h:
                long_oi_idx = i
            elif "空方" in h and "未平倉" in h and "口數" in h:
                short_oi_idx = i

        print(f"    [TAIFEX] 欄位定位: id={id_idx}, long_oi={long_oi_idx}, short_oi={short_oi_idx}")

        total_long_oi = 0
        total_short_oi = 0
        found_data = False
        for line in lines[1:]:
            fields = [f.strip().strip('"') for f in line.split(",")]
            if len(fields) > max(long_oi_idx, short_oi_idx):
                long_oi = _parse_number(fields[long_oi_idx])
                short_oi = _parse_number(fields[short_oi_idx])
                identity = fields[id_idx].strip() if len(fields) > id_idx else ""
                if long_oi is not None and long_oi > 0:
                    total_long_oi += long_oi
                    found_data = True
                if short_oi is not None and short_oi > 0:
                    total_short_oi += short_oi
                    found_data = True
                if identity:
                    print(f"      {identity}: long_oi={long_oi}, short_oi={short_oi}")

        total = total_long_oi + total_short_oi
        if total > 0:
            sentiment = round((total_long_oi - total_short_oi) / total * 100, 2)
            print(f"    [TAIFEX] {contract_id}: {sentiment}% (long={total_long_oi}, short={total_short_oi})")
            return sentiment
        elif not found_data:
            print(f"    [TAIFEX] {contract_id}: CSV 無有效資料")
    except Exception as e:
        print(f"  [錯誤] 解析{contract_id}多空指標失敗: {e}")
    return None


def _calc_sentiment_from_taifex_html(html_text, contract_id):
    """從期交所 HTML 表格解析多空指標 (MTX 等會回傳 HTML)"""
    try:
        soup = BeautifulSoup(html_text, "html.parser")
        tables = soup.find_all("table")
        total_long_oi = 0
        total_short_oi = 0
        found_data = False

        for table in tables:
            rows = table.find_all("tr")
            # 先找表頭，定位欄位
            header_row = None
            for row in rows:
                ths = row.find_all("th")
                if ths:
                    header_text = [th.get_text(strip=True) for th in ths]
                    if any("未平倉" in h for h in header_text):
                        header_row = header_text
                        break

            # 解析資料行
            for row in rows:
                cells = row.find_all("td")
                if len(cells) < 5:
                    continue
                cell_text = [c.get_text(strip=True).replace(",", "") for c in cells]

                # 找身份別 (外資/投信/自營商)
                identity = ""
                for ct in cell_text:
                    if "外資" in ct or "投信" in ct or "自營" in ct:
                        identity = ct
                        break
                if not identity:
                    continue

                # 在這一行找多方/空方未平倉口數
                # 通常表格結構: 身份別 | 多方交易口數 | ... | 多方未平倉口數 | 多方未平倉金額 | 空方未平倉口數 | ...
                nums = []
                for ct in cell_text:
                    v = _parse_number(ct)
                    if v is not None:
                        nums.append(v)

                # 期交所表格: 通常有 12+ 數字欄位
                # 多方未平倉口數 在位置 6 (0-indexed)，空方在位置 8
                if len(nums) >= 10:
                    long_oi = int(nums[6])   # 多方未平倉口數
                    short_oi = int(nums[8])  # 空方未平倉口數
                    total_long_oi += long_oi
                    total_short_oi += short_oi
                    found_data = True
                    print(f"      [HTML] {identity}: long_oi={long_oi}, short_oi={short_oi}")
                elif len(nums) >= 4:
                    # 較短的表格格式
                    long_oi = int(nums[-4]) if nums[-4] > 0 else 0
                    short_oi = int(nums[-2]) if nums[-2] > 0 else 0
                    if long_oi > 0 or short_oi > 0:
                        total_long_oi += long_oi
                        total_short_oi += short_oi
                        found_data = True
                        print(f"      [HTML] {identity}: long_oi={long_oi}, short_oi={short_oi}")

        total = total_long_oi + total_short_oi
        if total > 0:
            sentiment = round((total_long_oi - total_short_oi) / total * 100, 2)
            print(f"    [TAIFEX HTML] {contract_id}: {sentiment}% (long={total_long_oi}, short={total_short_oi})")
            return sentiment
        elif not found_data:
            print(f"    [TAIFEX HTML] {contract_id}: 表格無有效資料")
    except Exception as e:
        print(f"  [錯誤] HTML解析{contract_id}失敗: {e}")
        import traceback
        traceback.print_exc()
    return None


def fetch_sentiment_indicators(date_str=None):
    """
    抓取期權觀測指標：微台多空指標、小台多空指標、PCR（含前日對比）
    策略：FinMind 為主要資料源，TAIFEX CSV 為備用
    """
    print("🔮 抓取期權觀測指標...")
    if date_str is None:
        date_str = datetime.now().strftime("%Y%m%d")

    formatted_date = f"{date_str[:4]}/{date_str[4:6]}/{date_str[6:8]}"

    result = {
        "micro_sentiment": None,       # 微台多空指標
        "micro_sentiment_prev": None,
        "mini_sentiment": None,         # 小台多空指標
        "mini_sentiment_prev": None,
        "pcr_today": None,              # PCR 今日
        "pcr_prev": None,               # PCR 前一交易日
    }

    prev_date = _get_prev_trading_date(date_str)

    # ===== 方法1 (主要): FinMind TaiwanFuturesInstitutionalInvestors =====
    for data_id, key_prefix, label in [("MXF", "micro", "微台"), ("MTX", "mini", "小台")]:
        print(f"  [FinMind] 取得{label}多空指標 (data_id={data_id})...")
        # 今日 (不用前一交易日替代)
        fm_data = _finmind_fetch("TaiwanFuturesInstitutionalInvestors", data_id, date_str, date_str)
        if not fm_data:
            print(f"    [FinMind] ❌ {label}今日無資料")
        if fm_data:
            sentiment = _calc_sentiment_from_finmind(fm_data)
            if sentiment is not None:
                result[f"{key_prefix}_sentiment"] = sentiment
                print(f"    [FinMind] {label}多空: {sentiment}%")

        # 前日
        if prev_date:
            fm_prev = _finmind_fetch("TaiwanFuturesInstitutionalInvestors", data_id, prev_date, prev_date)
            if fm_prev:
                prev_sentiment = _calc_sentiment_from_finmind(fm_prev)
                if prev_sentiment is not None:
                    result[f"{key_prefix}_sentiment_prev"] = prev_sentiment
                    print(f"    [FinMind] {label}前日多空: {prev_sentiment}%")

    # ===== 方法2 (備用): TAIFEX 期交所 CSV =====
    for contract_id, key_prefix, label in [("MXF", "micro", "微台"), ("MTX", "mini", "小台")]:
        if result[f"{key_prefix}_sentiment"] is not None:
            continue  # FinMind 已成功取得，跳過
        print(f"  [TAIFEX備用] 取得{label}多空指標 (commodityId={contract_id})...")
        url3 = "https://www.taifex.com.tw/cht/3/futContractsDateDown"
        params3 = {
            "queryStartDate": formatted_date,
            "queryEndDate": formatted_date,
            "commodityId": contract_id,
        }
        resp3 = _safe_get(url3, params3)
        if resp3:
            sentiment = _calc_sentiment_from_taifex(resp3.text, contract_id)
            if sentiment is not None:
                result[f"{key_prefix}_sentiment"] = sentiment

        # 方法2 仍失敗 → 方法3: TAIFEX HTML 頁面 (非下載)
        if result[f"{key_prefix}_sentiment"] is None:
            print(f"  [TAIFEX HTML備用] 取得{label}多空 (commodity_id={contract_id})...")
            url_html = "https://www.taifex.com.tw/cht/3/futContractsDate"
            params_html = {
                "queryType": "1",
                "marketCode": "0",
                "commodity_id": contract_id,
                "queryDate": formatted_date,
            }
            resp_html = _safe_get(url_html, params_html)
            if resp_html:
                sentiment = _calc_sentiment_from_taifex_html(resp_html.text, contract_id)
                if sentiment is not None:
                    result[f"{key_prefix}_sentiment"] = sentiment
                    print(f"    [TAIFEX HTML] ✅ {label}多空: {sentiment}%")

        # 前日
        if prev_date and result[f"{key_prefix}_sentiment_prev"] is None:
            prev_formatted = f"{prev_date[:4]}/{prev_date[4:6]}/{prev_date[6:8]}"
            params4 = {
                "queryStartDate": prev_formatted,
                "queryEndDate": prev_formatted,
                "commodityId": contract_id,
            }
            resp4 = _safe_get(url3, params4)
            if resp4:
                prev_sentiment = _calc_sentiment_from_taifex(resp4.text, contract_id)
                if prev_sentiment is not None:
                    result[f"{key_prefix}_sentiment_prev"] = prev_sentiment

    # ===== PCR (Put/Call Ratio) - FinMind 為主要 =====
    print("  [FinMind] 取得 PCR...")
    fm_pcr = _finmind_fetch("TaiwanOptionPutCallRatio", "", date_str, date_str)
    if not fm_pcr:
        print("    [FinMind] ❌ PCR 今日無資料")
    if fm_pcr:
        for row in fm_pcr:
            pcr_val = row.get("PutCallRatio")
            if pcr_val is not None:
                result["pcr_today"] = round(float(pcr_val), 2)
                print(f"    [FinMind] ✅ PCR: {result['pcr_today']}%")
                break

    if prev_date:
        fm_pcr_prev = _finmind_fetch("TaiwanOptionPutCallRatio", "", prev_date, prev_date)
        if fm_pcr_prev:
            for row in fm_pcr_prev:
                pcr_val = row.get("PutCallRatio")
                if pcr_val is not None:
                    result["pcr_prev"] = round(float(pcr_val), 2)
                    break

    # ===== PCR 備用: TAIFEX HTML =====
    if result["pcr_today"] is None:
        print("  [TAIFEX備用] 取得 PCR...")
        url_pcr = "https://www.taifex.com.tw/cht/3/pcRatio"
        for target_date, key in [(formatted_date, "pcr_today"), (None, "pcr_prev")]:
            if target_date is None:
                if prev_date:
                    target_date = f"{prev_date[:4]}/{prev_date[4:6]}/{prev_date[6:8]}"
                else:
                    continue
            if result[key] is not None:
                continue
            params_pcr = {"queryStartDate": target_date, "queryEndDate": target_date}
            resp_pcr = _safe_get(url_pcr, params_pcr)
            if resp_pcr:
                try:
                    soup = BeautifulSoup(resp_pcr.text, "html.parser")
                    for table in soup.find_all("table"):
                        for row in table.find_all("tr"):
                            cells = row.find_all("td")
                            if len(cells) >= 3:
                                cell_text = [c.get_text(strip=True) for c in cells]
                                for t in cell_text:
                                    val = _parse_number(t.replace("%", ""))
                                    if val and 30 < val < 300 and result[key] is None:
                                        result[key] = val
                except:
                    pass

    print(f"  [結果] 微台={result['micro_sentiment']}, 小台={result['mini_sentiment']}, PCR={result['pcr_today']}")
    return result


# ============================================================
# 8. 融資維持率
# ============================================================
def fetch_margin_maintenance_ratio(date_str=None):
    """
    抓取整體融資維持率
    融資維持率 = 擔保品市值 / 融資金額 × 100%
    低於 130% 會被追繳，低於 120% 強制斷頭

    策略：
    1. FinMind TaiwanStockTotalMarginPurchaseShortSale 取得融資金額
       配合 FinMind TaiwanStockPrice 或加權指數估算擔保品市值
    2. TWSE MI_MARGN (creditFields 動態解析)
    3. 合理估算 (台股一般 150-170% 之間)
    """
    print("📉 抓取融資維持率...")
    if date_str is None:
        date_str = datetime.now().strftime("%Y%m%d")

    result = {"ratio": None, "date": date_str}

    # ===== 方法1: TWSE MI_MARGN =====
    print("  [方法1] TWSE MI_MARGN...")
    url = "https://www.twse.com.tw/exchangeReport/MI_MARGN"
    params = {"response": "json", "date": date_str, "selectType": "MS"}
    resp = _safe_get(url, params)

    margin_amount = None
    collateral_value = None

    if resp:
        try:
            data = resp.json()
            if data.get("stat") != "OK":
                print(f"    [TWSE MI_MARGN] stat={data.get('stat')}")
            else:
                credit_fields = data.get("creditFields", [])
                credit_list = data.get("creditList", [])
                print(f"    [TWSE] creditFields={credit_fields}")
                print(f"    [TWSE] creditList ({len(credit_list)} 行):")
                for i, row in enumerate(credit_list):
                    print(f"      [{i}] {row}")

                if credit_fields and credit_list:
                    # 動態找到各欄位索引
                    balance_idx = None
                    for i, fn in enumerate(credit_fields):
                        fn_str = str(fn)
                        if "今日" in fn_str and "餘額" in fn_str:
                            balance_idx = i
                            break
                    # 如果沒找到「今日餘額」，找最後一個數值欄位
                    if balance_idx is None:
                        for i in range(len(credit_fields) - 1, 0, -1):
                            fn_str = str(credit_fields[i])
                            if "餘額" in fn_str or "合計" in fn_str:
                                balance_idx = i
                                break
                    if balance_idx is None and len(credit_fields) >= 2:
                        balance_idx = len(credit_fields) - 1
                    print(f"    [TWSE] 使用 balance_idx={balance_idx} ({credit_fields[balance_idx] if balance_idx and balance_idx < len(credit_fields) else 'N/A'})")

                    for row in credit_list:
                        if not row or len(row) <= (balance_idx or 0):
                            continue
                        item_name = str(row[0]).strip()

                        # 嘗試所有可能的值欄位
                        if balance_idx is not None:
                            raw_val = _parse_number(row[balance_idx])
                        else:
                            # 從後面找第一個數值
                            raw_val = None
                            for ci in range(len(row) - 1, 0, -1):
                                raw_val = _parse_number(row[ci])
                                if raw_val is not None:
                                    break

                        if "融資金額" in item_name or ("融資" in item_name and "金額" in item_name):
                            if raw_val is not None:
                                margin_amount = raw_val * 1000 if raw_val < 1e9 else raw_val
                                print(f"    [TWSE] 融資金額: {margin_amount}")
                        elif "擔保" in item_name or "市值" in item_name:
                            if raw_val is not None:
                                collateral_value = raw_val * 1000 if raw_val < 1e9 else raw_val
                                print(f"    [TWSE] 擔保品市值: {collateral_value}")

                    if margin_amount and collateral_value and margin_amount > 0:
                        result["ratio"] = round(collateral_value / margin_amount * 100, 1)
                        print(f"    [TWSE] ✅ 融資維持率: {result['ratio']}%")
                    else:
                        print(f"    [TWSE] 無法計算: margin_amount={margin_amount}, collateral={collateral_value}")
        except Exception as e:
            print(f"  [錯誤] TWSE MI_MARGN: {e}")
            import traceback
            traceback.print_exc()

    # ===== 方法2: FinMind 取得融資金額 + 擔保品估算 =====
    if result["ratio"] is None:
        print("  [方法2] FinMind 融資金額估算...")
        target_date = date_str
        fm_data = _finmind_fetch("TaiwanStockTotalMarginPurchaseShortSale", "", target_date, target_date)
        if not fm_data:
            print("    [FinMind] ❌ 今日融資金額無資料")

        if fm_data:
            fm_margin_amt = None
            fm_margin_bal = None
            for row in fm_data:
                name = row.get("name", "")
                today_bal = row.get("TodayBalance")
                if "融資金額" in name or ("融資" in name and "金額" in name):
                    if today_bal is not None:
                        fm_margin_amt = float(today_bal) * 1000  # 仟元轉元
                        print(f"    [FinMind] 融資金額: {fm_margin_amt}")
                elif "融資" in name and "金額" not in name:
                    if today_bal is not None:
                        fm_margin_bal = float(today_bal)
                        print(f"    [FinMind] 融資餘額(張): {fm_margin_bal}")

            # 融資維持率一般在 150%-170% 之間
            # 如果有融資金額，可以搭配加權指數大致估算
            # 但更準確的做法是使用合理的歷史中位數
            if fm_margin_amt and fm_margin_amt > 0:
                # 嘗試從加權指數的變化推估維持率
                # 台股長期融資維持率中位數約 160%
                # 如果加權指數近日下跌則維持率下降
                # 這裡先用經驗估算 (之後可從個股精算)
                print(f"    [FinMind] 有融資金額 {fm_margin_amt}，需擔保品市值來計算維持率")
                print(f"    [FinMind] 嘗試用加權指數推估...")
                # 暫不估算，留給方法3

    # ===== 方法3: Goodinfo 或其他來源 =====
    if result["ratio"] is None:
        print("  [方法3] 嘗試 Goodinfo...")
        url3 = "https://goodinfo.tw/tw/StockMarginList.asp"
        headers_gi = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
            "Accept-Language": "zh-TW,zh;q=0.9",
            "Referer": "https://goodinfo.tw/tw/index.asp",
        }
        try:
            resp3 = requests.get(url3, headers=headers_gi, timeout=10)
            if resp3.status_code == 200:
                soup = BeautifulSoup(resp3.text, "html.parser")
                text = soup.get_text()
                import re
                patterns = [
                    r'整體.*?維持率.*?(\d{2,3}\.?\d*)\s*%',
                    r'維持率.*?(\d{2,3}\.?\d*)\s*%',
                    r'融資維持率.*?(\d{2,3}\.?\d*)',
                ]
                for pattern in patterns:
                    match = re.search(pattern, text)
                    if match:
                        val = float(match.group(1))
                        if 100 < val < 300:
                            result["ratio"] = round(val, 1)
                            print(f"    [Goodinfo] ✅ 融資維持率: {result['ratio']}%")
                            break
        except Exception as e:
            print(f"    [Goodinfo] ❌ {e}")

    # ===== 方法4: KGI 凱基證券頁面爬蟲 =====
    if result["ratio"] is None:
        print("  [方法4] 嘗試 KGI 凱基證券...")
        try:
            kgi_url = "https://www.kgi.com.tw/zh-tw/product-market/stock-market-overview/tw-stock-market/tw-stock-market-detail"
            kgi_params = {
                "a": "B658010E71E243C4A1D6B5F7BE914BDC",
                "b": "5D48401A7CE148CD8ABAC965F9B5AFBF",
            }
            resp_kgi = requests.get(kgi_url, params=kgi_params, headers=HEADERS, timeout=12)
            if resp_kgi.status_code == 200:
                soup = BeautifulSoup(resp_kgi.text, "html.parser")
                page_text = soup.get_text()
                # 找維持率數字 (通常在 130-200 之間)
                patterns = [
                    r'(?:整體)?融資維持率[^\d]*?(\d{2,3}\.?\d*)\s*%?',
                    r'維持率[^\d]*?(\d{2,3}\.?\d*)\s*%?',
                    r'(\d{2,3}\.\d{1,2})\s*%',  # 找任何百分比
                ]
                for pattern in patterns:
                    matches = re.findall(pattern, page_text)
                    for m in matches:
                        val = float(m)
                        if 100 < val < 300:
                            result["ratio"] = round(val, 1)
                            print(f"    [KGI] ✅ 融資維持率: {result['ratio']}%")
                            break
                    if result["ratio"]:
                        break
                # 也嘗試找 JSON-LD 或 script 中的數據
                if result["ratio"] is None:
                    scripts = soup.find_all("script")
                    for s in scripts:
                        if s.string and ("維持率" in s.string or "maintenance" in s.string.lower()):
                            nums = re.findall(r'(\d{2,3}\.\d{1,2})', s.string)
                            for n in nums:
                                val = float(n)
                                if 100 < val < 300:
                                    result["ratio"] = round(val, 1)
                                    print(f"    [KGI script] ✅ 融資維持率: {result['ratio']}%")
                                    break
                            if result["ratio"]:
                                break
        except Exception as e:
            print(f"    [KGI] ❌ {e}")

    # ===== 方法5: 鉅亨網 API =====
    if result["ratio"] is None:
        print("  [方法5] 嘗試鉅亨網 API...")
        try:
            url5 = "https://ws.api.cnyes.com/ws/api/v1/charting/history"
            params5 = {
                "resolution": "D",
                "symbol": "TWS:MARGIN_MAINTENANCE_RATIO:STOCK",
                "quote": 1,
            }
            resp5 = _safe_get(url5, params5, timeout=8)
            if resp5:
                data5 = resp5.json()
                if "data" in data5 and "quote" in data5["data"]:
                    quote = data5["data"]["quote"]
                    val = quote.get("6") or quote.get("closePrice")
                    if val and 100 < float(val) < 300:
                        result["ratio"] = round(float(val), 1)
                        print(f"    [鉅亨] ✅ 融資維持率: {result['ratio']}%")
        except Exception as e:
            print(f"    [鉅亨] ❌ {e}")

    # ===== 方法6: MacroMicro 圖表 API =====
    if result["ratio"] is None:
        print("  [方法6] 嘗試 MacroMicro...")
        try:
            mm_url = "https://www.macromicro.me/charts/data/53117"
            resp_mm = requests.get(mm_url, headers={
                "User-Agent": "Mozilla/5.0",
                "Accept": "application/json",
                "Referer": "https://www.macromicro.me/charts/53117/taiwan-taiex-maintenance-margin",
            }, timeout=10)
            if resp_mm.status_code == 200:
                mm_data = resp_mm.json()
                # MacroMicro 回傳格式: {"data": {"series": [{"data": [[timestamp, value], ...]}]}}
                series = mm_data.get("data", {}).get("series", [])
                if series:
                    points = series[0].get("data", [])
                    if points:
                        latest = points[-1]
                        val = float(latest[-1]) if isinstance(latest, list) else float(latest)
                        if 100 < val < 300:
                            result["ratio"] = round(val, 1)
                            print(f"    [MacroMicro] ✅ 融資維持率: {result['ratio']}%")
        except Exception as e:
            print(f"    [MacroMicro] ❌ {e}")

    # ===== 方法7: TWSE OpenAPI v1 (最新資料) =====
    if result["ratio"] is None:
        print("  [方法7] 嘗試 TWSE OpenAPI v1...")
        try:
            oa_url = "https://openapi.twse.com.tw/v1/exchangeReport/MI_MARGN"
            resp_oa = _safe_get(oa_url, timeout=10)
            if resp_oa:
                oa_data = resp_oa.json()
                if isinstance(oa_data, list) and len(oa_data) > 0:
                    # OpenAPI v1 回傳 JSON array，找融資金額和擔保品
                    oa_margin = None
                    oa_collateral = None
                    for item in oa_data:
                        name = item.get("Name", "") or item.get("ItemName", "")
                        val_str = item.get("TodayBalance", "") or item.get("Balance", "")
                        val = _parse_number(str(val_str))
                        if "融資金額" in name and val:
                            oa_margin = val * 1000 if val < 1e9 else val
                        elif ("擔保" in name or "市值" in name) and val:
                            oa_collateral = val * 1000 if val < 1e9 else val
                    if oa_margin and oa_collateral and oa_margin > 0:
                        result["ratio"] = round(oa_collateral / oa_margin * 100, 1)
                        print(f"    [TWSE OpenAPI] ✅ 融資維持率: {result['ratio']}%")
        except Exception as e:
            print(f"    [TWSE OpenAPI] ❌ {e}")

    if result["ratio"] is None:
        print("  ⚠️ 融資維持率: 所有 7 種方法均失敗")

    return result


# ============================================================
# 9. CNN Fear & Greed Index
# ============================================================
def fetch_cnn_fear_greed():
    """
    抓取 CNN Fear & Greed Index
    使用 CNN 的公開 API
    回傳: {value, label, prev_value, prev_label}
    """
    print("🎭 抓取 CNN Fear & Greed Index...")

    result = {
        "value": None, "label": None,
        "prev_value": None, "prev_label": None,
    }

    url = "https://production.dataviz.cnn.io/index/fearandgreed/graphdata"
    resp = _safe_get(url)

    if resp:
        try:
            data = resp.json()
            if "fear_and_greed" in data:
                fg = data["fear_and_greed"]
                result["value"] = round(fg.get("score", 0), 1) if fg.get("score") else None
                result["label"] = fg.get("rating", "")
                if "previous_close" in fg:
                    result["prev_value"] = round(fg["previous_close"], 1)
                if "previous_1_week" in fg:
                    result["week_ago_value"] = round(fg["previous_1_week"], 1)
        except:
            pass

    return result


# ============================================================
# 10. 比特幣恐慌與貪婪指數 (Alternative.me)
# ============================================================
def fetch_crypto_fear_greed():
    """
    抓取比特幣恐慌與貪婪指數
    來源: Alternative.me Crypto Fear & Greed Index
    0-24: Extreme Fear, 25-49: Fear, 50-74: Greed, 75-100: Extreme Greed
    回傳: {value, label, prev_value, prev_label}
    """
    print("₿ 抓取比特幣恐慌貪婪指數...")

    result = {
        "value": None, "label": None,
        "prev_value": None, "prev_label": None,
    }

    url = "https://api.alternative.me/fng/"
    params = {"limit": 2, "format": "json"}
    resp = _safe_get(url, params)

    if resp:
        try:
            data = resp.json()
            if "data" in data and len(data["data"]) > 0:
                today = data["data"][0]
                result["value"] = int(today["value"])
                result["label"] = today["value_classification"]
                if len(data["data"]) > 1:
                    prev = data["data"][1]
                    result["prev_value"] = int(prev["value"])
                    result["prev_label"] = prev["value_classification"]
        except:
            pass

    return result


# ============================================================
# 11. Put/Call Ratio (台指選擇權)
# ============================================================
def fetch_put_call_ratio(date_str=None):
    """
    抓取台指選擇權 Put/Call Ratio (未平倉量比)
    來源: 台灣期貨交易所
    PCR > 100% 偏多 (賣權未平倉多，莊家看多)
    PCR < 100% 偏空
    回傳: {ratio, put_oi, call_oi}
    """
    print("⚖️ 抓取 Put/Call Ratio...")
    if date_str is None:
        date_str = datetime.now().strftime("%Y%m%d")

    formatted_date = f"{date_str[:4]}/{date_str[4:6]}/{date_str[6:8]}"

    result = {"ratio": None, "put_oi": None, "call_oi": None}

    # 期交所選擇權每日交易量及未平倉量
    url = "https://www.taifex.com.tw/cht/3/pcRatio"
    params = {
        "queryStartDate": formatted_date,
        "queryEndDate": formatted_date,
    }
    resp = _safe_get(url, params)

    if resp:
        try:
            from bs4 import BeautifulSoup
            soup = BeautifulSoup(resp.text, "html.parser")
            tables = soup.find_all("table")
            for table in tables:
                rows = table.find_all("tr")
                for row in rows:
                    cells = row.find_all("td")
                    if len(cells) >= 4:
                        cell_text = [c.get_text(strip=True) for c in cells]
                        # 找含有比率的行
                        ratio_val = _parse_number(cell_text[-1]) if cell_text[-1] else None
                        if ratio_val and 30 < ratio_val < 300:
                            result["ratio"] = ratio_val
                            # 嘗試取得 Put OI / Call OI
                            put_oi = _parse_number(cell_text[1]) if len(cell_text) > 1 else None
                            call_oi = _parse_number(cell_text[2]) if len(cell_text) > 2 else None
                            if put_oi:
                                result["put_oi"] = put_oi
                            if call_oi:
                                result["call_oi"] = call_oi
                            break
        except:
            pass

    # 備用: 從期交所 CSV 下載
    if result["ratio"] is None:
        url2 = "https://www.taifex.com.tw/cht/3/callsAndPutsDate"
        params2 = {
            "queryType": "1",
            "commodity_id": "TXO",
            "queryDate": formatted_date,
        }
        resp2 = _safe_get(url2, params2)
        if resp2:
            try:
                from bs4 import BeautifulSoup
                soup2 = BeautifulSoup(resp2.text, "html.parser")
                tables = soup2.find_all("table")
                total_put_oi = 0
                total_call_oi = 0
                for table in tables:
                    for row in table.find_all("tr"):
                        cells = row.find_all("td")
                        cell_text = [c.get_text(strip=True) for c in cells]
                        # 找合計列
                        if any("合計" in t or "Total" in t for t in cell_text):
                            for i, t in enumerate(cell_text):
                                if "買權" in t or "Call" in t:
                                    call_oi_val = _parse_number(cell_text[i+2]) if i+2 < len(cell_text) else None
                                    if call_oi_val:
                                        total_call_oi = call_oi_val
                                elif "賣權" in t or "Put" in t:
                                    put_oi_val = _parse_number(cell_text[i+2]) if i+2 < len(cell_text) else None
                                    if put_oi_val:
                                        total_put_oi = put_oi_val
                if total_call_oi > 0 and total_put_oi > 0:
                    result["ratio"] = round(total_put_oi / total_call_oi * 100, 1)
                    result["put_oi"] = total_put_oi
                    result["call_oi"] = total_call_oi
            except:
                pass

    return result


# ============================================================
# 12. 美國 10 年期公債殖利率 (US10Y)
# ============================================================
def fetch_us10y():
    """
    抓取美國 10 年期公債殖利率
    鉅亨網代號: GI:US10Y
    """
    print("🏛️ 抓取美國10年期公債殖利率...")
    # 方法1: 鉅亨網
    data = _fetch_cnyes_chart("GI:US10Y", 10)
    # 方法2: investing.com (US10Y pair_id=23705)
    if not data:
        data = _fetch_investing_data(23705, 10)
    if not data:
        print("  ⚠️ US10Y: 所有來源失敗")

    data = data[-7:] if len(data) > 7 else data

    if data:
        return {
            "value": data[-1]["close"],
            "date": data[-1]["date"],
            "prev_value": data[-2]["close"] if len(data) >= 2 else None,
            "chart": data,
        }
    return {"value": None, "date": None, "prev_value": None, "chart": []}


# ============================================================
# 工具函數
# ============================================================
def _get_prev_trading_date(date_str):
    """取得前一個交易日 (簡單版，跳過週末)"""
    dt = datetime.strptime(date_str, "%Y%m%d")
    for i in range(1, 10):
        prev = dt - timedelta(days=i)
        if prev.weekday() < 5:  # 週一到週五
            return prev.strftime("%Y%m%d")
    return None


def fetch_all_data(date_str=None):
    """
    一次抓取所有資料
    回傳完整的 dict
    """
    if date_str is None:
        date_str = datetime.now().strftime("%Y%m%d")

    print(f"\n{'='*50}")
    print(f"  台灣股市戰略儀表板 - 資料抓取")
    print(f"  日期: {date_str}")
    print(f"{'='*50}\n")

    data = {
        "date": date_str,
        "generated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "taiex": fetch_taiex(date_str),
        "institutional": fetch_institutional(date_str),
        "foreign_top10": fetch_foreign_top10(date_str),
        "margin": fetch_margin_trading(date_str),
        "breadth": fetch_market_breadth(date_str),
        "usd_index": fetch_usd_index(),
        "jpy_rate": fetch_jpy_rate(),
        "vix": fetch_vix(),
        "futures": fetch_taiex_futures(date_str),
        "margin_ratio": fetch_margin_maintenance_ratio(date_str),
        "cnn_fg": fetch_cnn_fear_greed(),
        "crypto_fg": fetch_crypto_fear_greed(),
        "pcr": fetch_put_call_ratio(date_str),
        "us10y": fetch_us10y(),
        "futures_oi": fetch_futures_oi(date_str),
        "sentiment": fetch_sentiment_indicators(date_str),
    }

    print(f"\n✅ 資料抓取完成！")
    return data


if __name__ == "__main__":
    data = fetch_all_data()
    print(json.dumps(data, ensure_ascii=False, indent=2, default=str))
