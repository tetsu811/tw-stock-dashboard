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
    data_id: 資料ID (例如 TX, MXF, MTX)，可為空字串
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
            "start_date": start_formatted,
            "end_date": end_formatted,
            "token": finmind_token,
        }
        if data_id:
            params["data_id"] = data_id

        resp = _safe_get(url, params)
        if resp:
            data = resp.json()
            if data.get("status") == 200 and data.get("data"):
                print(f"    [FinMind] 成功取得 {dataset} (data_id={data_id or 'N/A'}) 共{len(data['data'])}筆")
                return data["data"]
            else:
                msg = data.get("msg", "unknown")
                print(f"    [FinMind] {dataset} 回應: status={data.get('status')}, msg={msg}")
    except Exception as e:
        print(f"    [FinMind] 錯誤: {e}")

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

    # ===== 方法1: TWSE MI_MARGN (用 creditFields 動態對應欄位) =====
    # 嘗試今天和前幾個交易日 (資料可能延遲)
    url = "https://www.twse.com.tw/exchangeReport/MI_MARGN"
    resp = None
    tried_dates = [date_str]
    prev = _get_prev_trading_date(date_str)
    if prev:
        tried_dates.append(prev)

    for try_date in tried_dates:
        params = {"response": "json", "date": try_date, "selectType": "MS"}
        resp = _safe_get(url, params)
        if resp:
            try:
                test_data = resp.json()
                if test_data.get("stat") == "OK" and test_data.get("creditList"):
                    print(f"    [TWSE] MI_MARGN 使用日期: {try_date}")
                    break
            except:
                pass
        resp = None  # 重置，繼續嘗試下一個日期
        print(f"    [TWSE] MI_MARGN 日期 {try_date} 無資料，嘗試前一天...")

    if resp:
        try:
            data = resp.json()
            print(f"    [TWSE] MI_MARGN stat={data.get('stat')}")

            # 使用 creditFields 取得欄位名稱，creditList 取得資料
            credit_fields = data.get("creditFields", [])
            credit_list = data.get("creditList", [])
            print(f"    [TWSE] creditFields: {credit_fields}")
            print(f"    [TWSE] creditList 共 {len(credit_list)} 行")

            # 找「今日餘額」欄位的位置
            balance_idx = None
            prev_balance_idx = None
            for i, field_name in enumerate(credit_fields):
                if "今日" in str(field_name) and "餘額" in str(field_name):
                    balance_idx = i
                elif "前日" in str(field_name) and "餘額" in str(field_name):
                    prev_balance_idx = i

            # 如果找不到精確匹配，嘗試最後一個欄位 (通常是今日餘額)
            if balance_idx is None and len(credit_fields) >= 2:
                balance_idx = len(credit_fields) - 1
                prev_balance_idx = len(credit_fields) - 2
                print(f"    [TWSE] 使用推測: balance_idx={balance_idx}, prev_idx={prev_balance_idx}")
            else:
                print(f"    [TWSE] 欄位定位: balance_idx={balance_idx}, prev_idx={prev_balance_idx}")

            # 遍歷每一行，根據「項目」名稱匹配
            for row in credit_list:
                if not row or len(row) <= (balance_idx or 0):
                    continue
                item_name = str(row[0]).strip()
                print(f"    [TWSE] 行: '{item_name}' => {row}")

                if "融資" in item_name and "金額" not in item_name and "張" not in item_name:
                    # 融資(交易單位) - 張數
                    if balance_idx is not None:
                        result["margin_balance"] = _parse_number(row[balance_idx])
                    if prev_balance_idx is not None:
                        prev_val = _parse_number(row[prev_balance_idx])
                        if prev_val is not None and result["margin_balance"] is not None:
                            result["margin_change"] = result["margin_balance"] - prev_val
                elif "融資金額" in item_name or ("融資" in item_name and "金額" in item_name):
                    # 融資金額(仟元)
                    if balance_idx is not None:
                        amt = _parse_number(row[balance_idx])
                        if amt is not None:
                            # 融資金額單位是仟元，轉為元
                            if "仟" in item_name or amt < 1e9:
                                result["margin_balance_amount"] = amt * 1000
                            else:
                                result["margin_balance_amount"] = amt
                elif "融券" in item_name and "金額" not in item_name:
                    # 融券(交易單位) - 張數
                    if balance_idx is not None:
                        result["short_balance"] = _parse_number(row[balance_idx])
                    if prev_balance_idx is not None:
                        prev_val = _parse_number(row[prev_balance_idx])
                        if prev_val is not None and result["short_balance"] is not None:
                            result["short_change"] = result["short_balance"] - prev_val

            print(f"    [TWSE] 結果: 融資餘額={result['margin_balance']}, 金額={result['margin_balance_amount']}, 融券={result['short_balance']}")
        except Exception as e:
            print(f"  [錯誤] 解析融資融券失敗: {e}")
            import traceback
            traceback.print_exc()

    # ===== 方法2: FinMind 備用 =====
    if result["margin_balance"] is None:
        print("  [備用] 嘗試 FinMind TaiwanStockTotalMarginPurchaseShortSale...")
        fm_data = _finmind_fetch("TaiwanStockTotalMarginPurchaseShortSale", "", date_str, date_str)
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
    # 嘗試多個 API 端點 (v8 已停用，v11 有 crumb 限制)
    urls = [
        f"https://query2.finance.yahoo.com/v8/finance/chart/{symbol}",
        f"https://query1.finance.yahoo.com/v8/finance/chart/{symbol}",
    ]

    params = {
        "range": period,
        "interval": interval,
        "includePrePost": "false",
    }

    resp = None
    for url in urls:
        resp = _safe_get(url, params)
        if resp and resp.status_code == 200:
            break
        resp = None

    if resp is None:
        # 備用: 嘗試 yfinance 風格的下載端點
        print(f"  [Yahoo] chart API 失敗，嘗試 download 端點...")
        from datetime import timedelta
        end_dt = datetime.now()
        days = int(period.replace("d", "")) if "d" in period else 30
        start_dt = end_dt - timedelta(days=days)
        dl_url = f"https://query1.finance.yahoo.com/v7/finance/download/{symbol}"
        dl_params = {
            "period1": str(int(start_dt.timestamp())),
            "period2": str(int(end_dt.timestamp())),
            "interval": interval,
            "events": "history",
        }
        resp_dl = _safe_get(dl_url, dl_params)
        if resp_dl and resp_dl.status_code == 200:
            try:
                lines = resp_dl.text.strip().split("\n")
                points = []
                for line in lines[1:]:
                    parts = line.split(",")
                    if len(parts) >= 5 and parts[4] != "null":
                        close_val = float(parts[4])
                        points.append({
                            "date": parts[0],
                            "close": round(close_val, 2)
                        })
                if points:
                    print(f"  [Yahoo] download 成功: {symbol} 共 {len(points)} 筆")
                    return points
            except Exception as e:
                print(f"  [Yahoo] download 解析失敗: {e}")
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
    """從期交所 CSV 資料計算多空指標"""
    if not resp_text:
        return None
    try:
        lines = resp_text.strip().split("\n")
        print(f"    [TAIFEX] {contract_id}: 共 {len(lines)} 行")
        if len(lines) > 0:
            print(f"    [TAIFEX] 標頭: {lines[0][:300]}")

        # 檢查是否是 HTML 而非 CSV
        if "<html" in lines[0].lower() or "<table" in lines[0].lower():
            print(f"    [TAIFEX] {contract_id}: 回傳 HTML 而非 CSV，跳過")
            return None

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
            print(f"    [TAIFEX] {contract_id}: 無有效資料")
    except Exception as e:
        print(f"  [錯誤] 解析{contract_id}多空指標失敗: {e}")
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
        # 今日
        fm_data = _finmind_fetch("TaiwanFuturesInstitutionalInvestors", data_id, date_str, date_str)
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

    # ===== PCR (Put/Call Ratio) =====
    # 嘗試從期交所 PCR 頁面取得
    url_pcr = "https://www.taifex.com.tw/cht/3/pcRatio"
    for target_date, key in [(formatted_date, "pcr_today"), (None, "pcr_prev")]:
        if target_date is None:
            if prev_date:
                target_date = f"{prev_date[:4]}/{prev_date[4:6]}/{prev_date[6:8]}"
            else:
                continue
        params_pcr = {"queryStartDate": target_date, "queryEndDate": target_date}
        resp_pcr = _safe_get(url_pcr, params_pcr)
        if resp_pcr:
            try:
                from bs4 import BeautifulSoup
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

    # ===== PCR 備用: FinMind =====
    if result["pcr_today"] is None:
        print("  [FinMind] 嘗試取得 PCR...")
        fm_pcr = _finmind_fetch("TaiwanOptionPutCallRatio", "", date_str, date_str)
        if fm_pcr:
            for row in fm_pcr:
                pcr_val = row.get("PutCallRatio")
                if pcr_val is not None:
                    result["pcr_today"] = round(float(pcr_val), 2)
                    print(f"    [FinMind] PCR: {result['pcr_today']}%")
                    break

    if result["pcr_prev"] is None and prev_date:
        fm_pcr_prev = _finmind_fetch("TaiwanOptionPutCallRatio", "", prev_date, prev_date)
        if fm_pcr_prev:
            for row in fm_pcr_prev:
                pcr_val = row.get("PutCallRatio")
                if pcr_val is not None:
                    result["pcr_prev"] = round(float(pcr_val), 2)
                    break

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
    1. 從 TWSE MI_MARGN 用 creditFields 動態解析
    2. 從 TWSE 融資融券統計頁面爬取
    3. 從 FinMind 計算 (個股加總)
    """
    print("📉 抓取融資維持率...")
    if date_str is None:
        date_str = datetime.now().strftime("%Y%m%d")

    result = {"ratio": None, "date": date_str}

    # ===== 方法1: TWSE MI_MARGN 用 creditFields 動態解析 =====
    url = "https://www.twse.com.tw/exchangeReport/MI_MARGN"
    params = {"response": "json", "date": date_str, "selectType": "MS"}
    resp = _safe_get(url, params)

    margin_amount = None  # 融資金額
    collateral_value = None  # 擔保品市值

    if resp:
        try:
            data = resp.json()
            credit_fields = data.get("creditFields", [])
            credit_list = data.get("creditList", [])
            print(f"    [TWSE] 融資維持率: creditFields={credit_fields}")

            # 找「今日餘額」欄位位置
            balance_idx = None
            for i, field_name in enumerate(credit_fields):
                if "今日" in str(field_name) and "餘額" in str(field_name):
                    balance_idx = i
            if balance_idx is None and len(credit_fields) >= 2:
                balance_idx = len(credit_fields) - 1

            # 遍歷所有行，找融資金額和擔保品
            for row in credit_list:
                if not row or len(row) <= (balance_idx or 0):
                    continue
                item_name = str(row[0]).strip()

                if "融資金額" in item_name or ("融資" in item_name and "金額" in item_name):
                    if balance_idx is not None:
                        amt = _parse_number(row[balance_idx])
                        if amt is not None:
                            # 仟元轉元
                            margin_amount = amt * 1000 if "仟" in item_name or amt < 1e9 else amt
                            print(f"    [TWSE] 融資金額: {margin_amount}")
                elif "擔保" in item_name or "市值" in item_name:
                    if balance_idx is not None:
                        val = _parse_number(row[balance_idx])
                        if val is not None:
                            collateral_value = val * 1000 if "仟" in item_name or val < 1e9 else val
                            print(f"    [TWSE] 擔保品市值: {collateral_value}")

            if margin_amount and collateral_value and margin_amount > 0:
                result["ratio"] = round(collateral_value / margin_amount * 100, 1)
                print(f"    [TWSE] 融資維持率: {result['ratio']}%")
        except Exception as e:
            print(f"  [錯誤] 解析融資維持率失敗: {e}")

    # ===== 方法2: 如果 MI_MARGN 沒有擔保品數據，嘗試爬取 TWSE 統計頁面 =====
    if result["ratio"] is None:
        print("  [備用] 嘗試 TWSE 融資融券統計頁面...")
        url2 = "https://www.twse.com.tw/exchangeReport/MI_MARGN"
        params2 = {"response": "json", "date": date_str, "selectType": "ALL"}
        resp2 = _safe_get(url2, params2)

        if resp2:
            try:
                data2 = resp2.json()
                # selectType=ALL 回傳個股融資融券明細
                # 可能包含 maintainRatio 或在 data 中有維持率欄位
                # 嘗試從整體統計取得
                if "creditList" in data2:
                    all_margin_amt = 0
                    all_collateral = 0
                    for row in data2["creditList"]:
                        if len(row) >= 6:
                            # 嘗試找個股的融資金額和擔保品
                            pass  # 個股資料太多，跳過

                # 嘗試其他可用欄位
                if "marginNote" in data2:
                    note = str(data2["marginNote"])
                    print(f"    [TWSE] marginNote: {note[:200]}")
            except Exception as e:
                print(f"  [錯誤] 備用融資維持率解析: {e}")

    # ===== 方法3: 嘗試從新聞/金融網站取得 =====
    if result["ratio"] is None:
        print("  [備用] 嘗試從 Goodinfo 取得融資維持率...")
        url3 = "https://goodinfo.tw/tw/StockMarginList.asp"
        resp3 = _safe_get(url3, timeout=10)
        if resp3:
            try:
                from bs4 import BeautifulSoup
                soup = BeautifulSoup(resp3.text, "html.parser")
                # 找整體維持率
                text = soup.get_text()
                import re
                # 搜尋「整體融資維持率」或類似文字後面跟著百分比
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
                            print(f"    [Goodinfo] 融資維持率: {result['ratio']}%")
                            break
            except Exception as e:
                print(f"  [錯誤] Goodinfo 解析: {e}")

    if result["ratio"] is None:
        print("  ⚠️ 無法取得融資維持率")

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
