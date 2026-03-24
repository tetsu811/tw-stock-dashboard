"""
台灣股市戰略儀表板 - 資料抓取模組
資料來源：
  - 台灣證券交易所 (TWSE) 開放資料 API
  - 證券櫃檯買賣中心 (TPEx) API
  - Yahoo Finance API (美元指數、日圓、VIX)
"""

import requests
import json
import time
import os
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
    從 FinMind API 取得資料 (備用資料來源)
    dataset: 資料集名稱
    data_id: 資料ID (例如 TX, MXF, MTX)
    start_date: YYYYMMDD 格式
    end_date: YYYYMMDD 格式
    """
    finmind_token = os.environ.get("FINMIND_TOKEN", "")
    if not finmind_token:
        print(f"    [跳過] FinMind token 未設置，無法使用備用資料源")
        return None

    try:
        url = "https://api.finmindtrade.com/api/v4/data"
        # 將 YYYYMMDD 轉換為 YYYY-MM-DD
        start_formatted = f"{start_date[:4]}-{start_date[4:6]}-{start_date[6:8]}"
        end_formatted = f"{end_date[:4]}-{end_date[4:6]}-{end_date[6:8]}"

        params = {
            "dataset": dataset,
            "data_id": data_id,
            "start_date": start_formatted,
            "end_date": end_formatted,
            "token": finmind_token,
        }
        resp = _safe_get(url, params)
        if resp:
            data = resp.json()
            if data.get("status") == 200 and data.get("data"):
                print(f"    [FinMind] 成功取得 {dataset} (data_id={data_id})")
                return data["data"]
    except Exception as e:
        print(f"    [FinMind] 錯誤: {e}")

    return None


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
        if "data" in data:
            for row in data["data"]:
                item = {
                    "stock_id": str(row[0]).strip(),
                    "stock_name": str(row[1]).strip(),
                    "buy": _parse_number(row[2]),
                    "sell": _parse_number(row[3]),
                    "net": _parse_number(row[4]),
                }
                if item["net"] and item["net"] > 0:
                    result["top_buy"].append(item)
                elif item["net"] and item["net"] < 0:
                    result["top_sell"].append(item)

            result["top_buy"] = sorted(result["top_buy"], key=lambda x: x["net"], reverse=True)[:10]
            result["top_sell"] = sorted(result["top_sell"], key=lambda x: x["net"])[:10]
    except Exception as e:
        print(f"  [錯誤] 解析外資排行失敗: {e}")
        return _fetch_foreign_top10_from_t86(date_str)

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

    return result


# ============================================================
# 4. 融資融券
# ============================================================
def fetch_margin_trading(date_str=None):
    """
    抓取融資融券餘額
    回傳: {margin_balance, margin_change, short_balance, short_change}
    """
    print("💰 抓取融資融券...")
    if date_str is None:
        date_str = datetime.now().strftime("%Y%m%d")

    url = "https://www.twse.com.tw/exchangeReport/MI_MARGN"
    params = {"response": "json", "date": date_str, "selectType": "MS"}
    resp = _safe_get(url, params)

    result = {
        "margin_balance": None, "margin_change": None,
        "margin_balance_amount": None,
        "short_balance": None, "short_change": None,
    }

    if resp is None:
        return result

    try:
        data = resp.json()
        # creditList 包含融資融券彙總
        if "creditList" in data and len(data["creditList"]) > 0:
            summary = data["creditList"][-1]  # 最後一列是合計
            print(f"    [TWSE] creditList 欄位數: {len(summary)}")
            print(f"    [TWSE] 行資料: {summary}")

            # 嘗試多種欄位位置，TWSE API 可能有不同版本
            # 通常格式: [日期, 融資買進, 融資賣出, 融資餘額(張), 融資餘額(金額), 融資餘額變化, 融券買進, 融券賣出, 融券餘額(張), 融券餘額(金額), ...]

            # 優先嘗試常見欄位位置
            if len(summary) > 3:
                result["margin_balance"] = _parse_number(summary[3])  # 融資餘額張數
            if len(summary) > 4:
                result["margin_balance_amount"] = _parse_number(summary[4])  # 融資餘額金額
            if len(summary) > 5:
                result["margin_change"] = _parse_number(summary[5])  # 融資餘額變化

            # 嘗試取得融券，可能在 field[8] 或 field[9]
            if len(summary) > 8:
                result["short_balance"] = _parse_number(summary[8])
            elif len(summary) > 9:
                result["short_balance"] = _parse_number(summary[9])

            print(f"    [TWSE] 融資餘額: {result['margin_balance']}, 融券餘額: {result['short_balance']}")
    except Exception as e:
        print(f"  [錯誤] 解析融資融券失敗: {e}")

    # 如果仍無資料，嘗試 FinMind 備用資料源
    if result["margin_balance"] is None:
        print("  [備用] 嘗試 FinMind API 取得融資融券...")
        fm_data = _finmind_fetch("TaiwanStockTotalMarginPurchaseShortSale", "", date_str, date_str)
        if fm_data:
            try:
                for row in fm_data:
                    margin_today = _parse_number(row.get("MarginPurchaseTodayBalance", ""))
                    short_today = _parse_number(row.get("ShortSaleTodayBalance", ""))
                    if margin_today is not None:
                        result["margin_balance"] = margin_today
                    if short_today is not None:
                        result["short_balance"] = short_today
                    print(f"    [FinMind] 融資餘額: {result['margin_balance']}, 融券餘額: {result['short_balance']}")
            except Exception as e:
                print(f"  [錯誤] FinMind 解析失敗: {e}")

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

    # 上市 - 從 TWSE 每日收盤行情取得漲跌家數
    # 方法1: 從 TAIEX 指數頁面直接取得統計
    url_stat = "https://www.twse.com.tw/exchangeReport/BWIBBU_d"
    params_stat = {"response": "json", "date": date_str, "selectType": "ALL"}

    # 方法2: 從個股行情計算漲跌家數
    url2 = "https://www.twse.com.tw/exchangeReport/MI_INDEX"
    params2 = {"response": "json", "date": date_str, "type": "ALLBUT0999"}
    resp2 = _safe_get(url2, params2)

    if resp2:
        try:
            data2 = resp2.json()
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

                for row in rows:
                    if not isinstance(row, list) or len(row) < 10:
                        continue
                    try:
                        # 漲跌方向通常在第9或第10欄
                        for ci in [9, 8, 10]:
                            if ci < len(row):
                                val_str = str(row[ci]).strip()
                                if val_str in ["+", "-", "X", " "]:
                                    if val_str == "+":
                                        up_count += 1
                                    elif val_str == "-":
                                        down_count += 1
                                    elif val_str == "X" or val_str == " ":
                                        flat_count += 1
                                    break
                                # 也可能是帶符號的數字
                                elif val_str.startswith("+") or "▲" in val_str:
                                    up_count += 1
                                    break
                                elif val_str.startswith("-") or "▼" in val_str:
                                    down_count += 1
                                    break
                    except:
                        pass

                if up_count > 0:
                    result["tse_up"] = up_count
                    result["tse_down"] = down_count
                    result["tse_flat"] = flat_count
                    break
        except:
            pass

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

    return result


# ============================================================
# 6. 美元指數、日圓、VIX (Yahoo Finance)
# ============================================================
def fetch_yahoo_chart(symbol, period="30d", interval="1d"):
    """
    從 Yahoo Finance API 抓取歷史價格
    symbol: "DX=F" (美元指數), "JPY=X" (USD/JPY), "^VIX"
    回傳: [{"date": "2024-01-01", "close": 104.5}, ...]
    """
    url = "https://query1.finance.yahoo.com/v8/finance/chart/" + symbol
    params = {
        "range": period,
        "interval": interval,
        "includePrePost": "false",
    }
    resp = _safe_get(url, params)

    if resp is None:
        return []

    try:
        data = resp.json()
        result_data = data["chart"]["result"][0]
        timestamps = result_data["timestamp"]
        closes = result_data["indicators"]["quote"][0]["close"]

        points = []
        for ts, close in zip(timestamps, closes):
            if close is not None:
                dt = datetime.fromtimestamp(ts)
                points.append({
                    "date": dt.strftime("%Y-%m-%d"),
                    "close": round(close, 2)
                })
        return points
    except Exception as e:
        print(f"  [錯誤] Yahoo Finance {symbol} 解析失敗: {e}")
        return []


def fetch_usd_index():
    """抓取美元指數近30天"""
    print("💵 抓取美元指數...")
    return fetch_yahoo_chart("DX=F", "35d", "1d")


def fetch_jpy_rate():
    """抓取日圓匯率近30天 (USD/JPY)"""
    print("💴 抓取日圓匯率...")
    return fetch_yahoo_chart("JPY=X", "35d", "1d")


def fetch_vix():
    """抓取 VIX 指數 (近7天含圖表資料)"""
    print("😱 抓取 VIX 指數 (近7天)...")
    data = fetch_yahoo_chart("^VIX", "10d", "1d")
    # 只取最近7筆
    data = data[-7:] if len(data) > 7 else data
    if data:
        return {
            "value": data[-1]["close"],
            "date": data[-1]["date"],
            "prev_value": data[-2]["close"] if len(data) >= 2 else None,
            "chart": data,  # 近7天完整資料供畫圖
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
            for line in lines[1:]:
                fields = [f.strip().strip('"') for f in line.split(",")]
                if len(fields) < 13:
                    continue
                identity = fields[1].strip() if len(fields) > 1 else ""

                # 正確的欄位位置 (0-indexed):
                # 6: 多空交易口數淨額 (增減口數) ← NET TRADE CHANGE
                # 12: 多空未平倉口數淨額 ← NET OI
                net_change = _parse_number(fields[6]) if len(fields) > 6 else None
                net_oi = _parse_number(fields[12]) if len(fields) > 12 else None

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


def fetch_sentiment_indicators(date_str=None):
    """
    抓取期權觀測指標：微台多空指標、小台多空指標、PCR（含前日對比）
    來源: 期交所
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

    # 嘗試從期交所 PCR 頁面取得
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
                    if len(cells) >= 3:
                        cell_text = [c.get_text(strip=True) for c in cells]
                        # PCR 數據
                        for i, t in enumerate(cell_text):
                            val = _parse_number(t.replace("%", ""))
                            if val and 30 < val < 300:
                                if result["pcr_today"] is None:
                                    result["pcr_today"] = val
        except:
            pass

    # 微台 & 小台多空指標
    # 計算方式：(多方未平倉 - 空方未平倉) / 總未平倉 * 100%
    for contract_id, key_prefix in [("MXF", "micro"), ("MTX", "mini")]:
        print(f"    [情緒指標] 計算{contract_id}多空指標...")
        url3 = "https://www.taifex.com.tw/cht/3/futContractsDateDown"
        params3 = {
            "queryStartDate": formatted_date,
            "queryEndDate": formatted_date,
            "commodityId": contract_id,
        }
        resp3 = _safe_get(url3, params3)
        if resp3:
            try:
                lines = resp3.text.strip().split("\n")
                total_long_oi = 0
                total_short_oi = 0
                for line in lines[1:]:
                    fields = [f.strip().strip('"') for f in line.split(",")]
                    if len(fields) > 10:
                        # 正確的欄位位置: 8=多方未平倉口數, 10=空方未平倉口數
                        long_oi = _parse_number(fields[8])
                        short_oi = _parse_number(fields[10])
                        if long_oi:
                            total_long_oi += long_oi
                        if short_oi:
                            total_short_oi += short_oi
                total = total_long_oi + total_short_oi
                if total > 0:
                    sentiment = round((total_long_oi - total_short_oi) / total * 100, 2)
                    result[f"{key_prefix}_sentiment"] = sentiment
                    print(f"      {key_prefix}: {sentiment}% (long={total_long_oi}, short={total_short_oi})")
            except Exception as e:
                print(f"  [錯誤] 解析{contract_id}多空指標失敗: {e}")

    # 抓取前一交易日的數據做比較
    prev_date = _get_prev_trading_date(date_str)
    if prev_date:
        prev_formatted = f"{prev_date[:4]}/{prev_date[4:6]}/{prev_date[6:8]}"

        # 前日 PCR
        url_prev = "https://www.taifex.com.tw/cht/3/pcRatio"
        params_prev = {"queryStartDate": prev_formatted, "queryEndDate": prev_formatted}
        resp_prev = _safe_get(url_prev, params_prev)
        if resp_prev:
            try:
                from bs4 import BeautifulSoup
                soup_prev = BeautifulSoup(resp_prev.text, "html.parser")
                for table in soup_prev.find_all("table"):
                    for row in table.find_all("tr"):
                        cells = row.find_all("td")
                        if len(cells) >= 3:
                            for c in cells:
                                val = _parse_number(c.get_text(strip=True).replace("%", ""))
                                if val and 30 < val < 300:
                                    if result["pcr_prev"] is None:
                                        result["pcr_prev"] = val
            except:
                pass

        # 前日微台、小台
        for contract_id, key_prefix in [("MXF", "micro"), ("MTX", "mini")]:
            url4 = "https://www.taifex.com.tw/cht/3/futContractsDateDown"
            params4 = {
                "queryStartDate": prev_formatted,
                "queryEndDate": prev_formatted,
                "commodityId": contract_id,
            }
            resp4 = _safe_get(url4, params4)
            if resp4:
                try:
                    lines = resp4.text.strip().split("\n")
                    total_long = 0
                    total_short = 0
                    for line in lines[1:]:
                        fields = [f.strip().strip('"') for f in line.split(",")]
                        if len(fields) > 10:
                            # 正確的欄位位置: 8=多方未平倉口數, 10=空方未平倉口數
                            lo = _parse_number(fields[8])
                            so = _parse_number(fields[10])
                            if lo: total_long += lo
                            if so: total_short += so
                    total = total_long + total_short
                    if total > 0:
                        result[f"{key_prefix}_sentiment_prev"] = round((total_long - total_short) / total * 100, 2)
                except Exception as e:
                    print(f"  [錯誤] 解析前日{contract_id}多空指標失敗: {e}")

    return result


# ============================================================
# 8. 融資維持率
# ============================================================
def fetch_margin_maintenance_ratio(date_str=None):
    """
    抓取整體融資維持率
    融資維持率 = 擔保品市值 / 融資金額 × 100%
    低於 130% 會被追繳，低於 120% 強制斷頭
    """
    print("📉 抓取融資維持率...")
    if date_str is None:
        date_str = datetime.now().strftime("%Y%m%d")

    result = {"ratio": None, "date": date_str}

    # 從 TWSE 融資融券彙總表計算
    url = "https://www.twse.com.tw/exchangeReport/MI_MARGN"
    params = {"response": "json", "date": date_str, "selectType": "MS"}
    resp = _safe_get(url, params)

    if resp:
        try:
            data = resp.json()
            if "creditList" in data and len(data["creditList"]) > 0:
                summary = data["creditList"][-1]
                print(f"    [TWSE] creditList 欄位數: {len(summary)}")
                print(f"    [TWSE] 完整行資料: {summary}")

                # 嘗試各種欄位位置組合
                # creditList 格式可能是: [日期, 融資買進, 融資賣出, 融資餘額(張), 融資餘額(金額), ...]
                # 或更多欄位包含擔保品市值
                margin_amt = None
                collateral = None
                # 遍歷所有數值欄位，找融資金額和擔保品
                values = []
                for i, cell in enumerate(summary):
                    val = _parse_number(cell)
                    if val and val > 1e9:  # 大於10億的值
                        values.append((i, val))
                        print(f"    [TWSE] field[{i}] (超過10億): {val}")

                # 嘗試直接計算: 如果有多個大數值，擔保品通常 > 融資金額
                if len(values) >= 2:
                    # 按數值排序，最大的可能是擔保品市值，次大的是融資金額
                    values_sorted = sorted(values, key=lambda x: x[1], reverse=True)
                    # 融資維持率 = 擔保品 / 融資 * 100，通常在 130-200 之間
                    for i in range(len(values_sorted)):
                        for j in range(i+1, len(values_sorted)):
                            ratio_test = values_sorted[i][1] / values_sorted[j][1] * 100
                            if 100 < ratio_test < 300:
                                result["ratio"] = round(ratio_test, 1)
                                print(f"    [TWSE] 融資維持率: {result['ratio']}% (擔保品={values_sorted[i][1]}, 融資={values_sorted[j][1]})")
                                break
                        if result["ratio"]:
                            break
        except Exception as e:
            print(f"  [錯誤] 解析融資維持率失敗: {e}")

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
    來源: Yahoo Finance (^TNX)
    回傳: {value, change, chart (近7天)}
    """
    print("🏛️ 抓取美國10年期公債殖利率...")
    data = fetch_yahoo_chart("%5ETNX", "10d", "1d")
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
