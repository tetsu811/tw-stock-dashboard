#!/usr/bin/env python3
"""
台灣股市儀表板 - API 來源測試工具
在本機或 GitHub Actions 執行，驗證每個 API 是否能正常回傳資料。

用法:
  python test_sources.py
  python test_sources.py --section dxy        # 只測 DXY
  python test_sources.py --section margin     # 只測融資
  python test_sources.py --section breadth    # 只測漲跌家數
  python test_sources.py --section ratio      # 只測融資維持率
  python test_sources.py --section foreign    # 只測外資排行
  python test_sources.py --section all        # 全部測 (預設)
"""

import requests
import json
import time
import sys
import os
import re
from datetime import datetime, timedelta

today = datetime.now().strftime('%Y-%m-%d')
dow = datetime.now().weekday()
if dow == 5:
    trade_date = (datetime.now() - timedelta(days=1)).strftime('%Y-%m-%d')
elif dow == 6:
    trade_date = (datetime.now() - timedelta(days=2)).strftime('%Y-%m-%d')
else:
    trade_date = today

twse_date = trade_date.replace('-', '')
twse_roc_date = f"{int(twse_date[:4])-1911}/{twse_date[4:6]}/{twse_date[6:8]}"

HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
    'Accept-Language': 'zh-TW,zh;q=0.9',
}

passed = 0
failed = 0

def test(name, func):
    global passed, failed
    try:
        result = func()
        if result:
            print(f"  ✅ {name}")
            print(f"     → {str(result)[:300]}")
            passed += 1
            return result
        else:
            print(f"  ❌ {name}: returned None/empty")
            failed += 1
    except Exception as e:
        print(f"  ❌ {name}: {type(e).__name__}: {str(e)[:200]}")
        failed += 1
    return None


# ============================================================
def test_dxy():
    """測試 DXY (美元指數) 所有來源"""
    print("\n" + "="*60)
    print("DXY (美元指數)")
    print("="*60)

    # 1. cnyes GI:DXY (charting/history with from/to)
    def _cnyes_gi_dxy():
        end_ts = int(time.time())
        start_ts = end_ts - 30 * 86400
        url = "https://ws.api.cnyes.com/ws/api/v1/charting/history"
        params = {"resolution": "D", "symbol": "GI:DXY", "from": str(start_ts), "to": str(end_ts)}
        r = requests.get(url, params=params, headers=HEADERS, timeout=10)
        data = r.json()
        if data.get("data", {}).get("c"):
            closes = data["data"]["c"]
            return f"GI:DXY works! {len(closes)} points, latest={closes[-1]}"
        return None
    test("cnyes charting GI:DXY (from/to)", _cnyes_gi_dxy)

    # 2. cnyes GI:DXY (quote mode)
    def _cnyes_gi_dxy_quote():
        url = "https://ws.api.cnyes.com/ws/api/v1/charting/history"
        params = {"resolution": "D", "symbol": "GI:DXY", "quote": 1}
        r = requests.get(url, params=params, headers=HEADERS, timeout=10)
        data = r.json()
        print(f"     [DEBUG] statusCode={data.get('statusCode')}, keys={list(data.get('data',{}).keys()) if 'data' in data else 'N/A'}")
        quote = data.get("data", {}).get("quote", {})
        if quote:
            print(f"     [DEBUG] quote keys={list(quote.keys())}")
            val = quote.get("6") or quote.get("closePrice") or quote.get("priceClose")
            if val:
                return f"GI:DXY quote={val}"
        return None
    test("cnyes quote GI:DXY", _cnyes_gi_dxy_quote)

    # 3. Try alternative symbols
    for sym in ["GI:DXY00", "TWG:DXY00", "FX:DXY", "GI:DX1!", "TWF:DX"]:
        def _try_sym(s=sym):
            url = "https://ws.api.cnyes.com/ws/api/v1/charting/history"
            params = {"resolution": "D", "symbol": s, "quote": 1}
            r = requests.get(url, params=params, headers=HEADERS, timeout=10)
            data = r.json()
            if data.get("data", {}).get("c"):
                return f"{s} works! latest={data['data']['c'][-1]}"
            if data.get("data", {}).get("quote"):
                q = data["data"]["quote"]
                val = q.get("6") or q.get("closePrice")
                if val:
                    return f"{s} quote={val}"
            return None
        test(f"cnyes {sym}", _try_sym)

    # 4. cnyes search for DXY
    def _cnyes_search():
        url = "https://ws.api.cnyes.com/ws/api/v1/universal/search"
        params = {"q": "DXY"}
        r = requests.get(url, params=params, headers=HEADERS, timeout=10)
        data = r.json()
        if data.get("statusCode") == 200:
            items = data.get("data", {}).get("items", [])
            results = [(i.get("symbol","?"), i.get("name","?")) for i in items[:15]]
            return f"Search 'DXY': {results}"
        return None
    test("cnyes search 'DXY'", _cnyes_search)

    def _cnyes_search_usd():
        url = "https://ws.api.cnyes.com/ws/api/v1/universal/search"
        params = {"q": "美元指數"}
        r = requests.get(url, params=params, headers=HEADERS, timeout=10)
        data = r.json()
        if data.get("statusCode") == 200:
            items = data.get("data", {}).get("items", [])
            results = [(i.get("symbol","?"), i.get("name","?")) for i in items[:15]]
            return f"Search '美元指數': {results}"
        return None
    test("cnyes search '美元指數'", _cnyes_search_usd)

    # 5. investing.com
    def _investing_dxy():
        end_ts = int(time.time())
        start_ts = end_ts - 30 * 86400
        url = "https://api.investing.com/api/financialdata/8827/historical/chart/"
        params = {"period": "P1M", "interval": "P1D", "pointscount": 30}
        h = {**HEADERS, 'domain-id': 'www.investing.com'}
        r = requests.get(url, params=params, headers=h, timeout=10)
        data = r.json()
        if isinstance(data, dict) and data.get("data"):
            last = data["data"][-1]
            return f"investing.com DXY: {len(data['data'])} points, last={last}"
        return f"response={str(data)[:200]}"
    test("investing.com DXY (pair_id=8827)", _investing_dxy)

    # 6. FRED API (free, no key for limited use)
    def _fred_dxy():
        url = "https://api.stlouisfed.org/fred/series/observations"
        params = {
            "series_id": "DTWEXBGS",
            "api_key": "DEMO_KEY",
            "file_type": "json",
            "sort_order": "desc",
            "limit": 30,
        }
        r = requests.get(url, params=params, headers=HEADERS, timeout=10)
        data = r.json()
        if "observations" in data and data["observations"]:
            obs = data["observations"]
            return f"FRED DTWEXBGS: {len(obs)} points, latest={obs[0]}"
        return None
    test("FRED API DTWEXBGS", _fred_dxy)


# ============================================================
def test_margin():
    """測試融資餘額所有來源"""
    print("\n" + "="*60)
    print("融資餘額 (Margin Trading)")
    print("="*60)

    # 1. FinMind
    def _finmind():
        url = "https://api.finmindtrade.com/api/v4/data"
        params = {"dataset": "TaiwanStockTotalMarginPurchaseShortSale",
                  "start_date": trade_date, "end_date": trade_date}
        token = os.environ.get("FINMIND_TOKEN", "")
        if token:
            params["token"] = token
        r = requests.get(url, params=params, headers=HEADERS, timeout=15)
        data = r.json()
        if data.get("status") == 200 and data.get("data"):
            rows = data["data"]
            print(f"     [DEBUG] {len(rows)} rows")
            for row in rows:
                print(f"     [DEBUG] name={row.get('name')}, TodayBalance={row.get('TodayBalance')}, YesBalance={row.get('YesBalance')}")
            return f"FinMind: {len(rows)} rows, names={[r.get('name') for r in rows]}"
        return f"status={data.get('status')}, msg={data.get('msg','')}"
    test("FinMind TaiwanStockTotalMarginPurchaseShortSale", _finmind)

    # 2. TWSE MI_MARGN selectType=MS
    def _twse_ms():
        url = "https://www.twse.com.tw/exchangeReport/MI_MARGN"
        params = {"response": "json", "date": twse_date, "selectType": "MS"}
        r = requests.get(url, params=params, headers=HEADERS, timeout=15)
        data = r.json()
        if data.get("stat") == "OK":
            cf = data.get("creditFields", [])
            cl = data.get("creditList", [])
            print(f"     [DEBUG] creditFields={cf}")
            for row in cl:
                print(f"     [DEBUG] creditList row={row}")
            return f"TWSE MI_MARGN MS: creditFields={cf}, creditList={len(cl)} rows"
        return f"stat={data.get('stat')}"
    test("TWSE MI_MARGN selectType=MS", _twse_ms)

    # 3. TWSE OpenAPI TWT93U
    def _twse_twt93u():
        url = "https://openapi.twse.com.tw/v1/exchangeReport/TWT93U"
        r = requests.get(url, headers=HEADERS, timeout=15)
        data = r.json()
        if data and isinstance(data, list):
            print(f"     [DEBUG] first record keys={list(data[0].keys())}")
            print(f"     [DEBUG] first record={json.dumps(data[0], ensure_ascii=False)[:300]}")
            return f"TWT93U: {len(data)} records"
        return None
    test("TWSE OpenAPI TWT93U", _twse_twt93u)


# ============================================================
def test_breadth():
    """測試漲跌家數所有來源"""
    print("\n" + "="*60)
    print("漲跌家數 (Market Breadth)")
    print("="*60)

    # 1. TWSE MI_INDEX type=ALLBUT0999
    def _mi_index():
        url = "https://www.twse.com.tw/exchangeReport/MI_INDEX"
        params = {"response": "json", "date": twse_date, "type": "ALLBUT0999"}
        r = requests.get(url, params=params, headers=HEADERS, timeout=15)
        data = r.json()
        if data.get("stat") == "OK":
            data_keys = [k for k in data.keys() if k.startswith("data")]
            groups = data.get("groups", [])
            # Print sample of each data table
            for dk in data_keys:
                rows = data[dk]
                if rows and len(rows) > 0:
                    print(f"     [DEBUG] {dk}: {len(rows)} rows, sample[0]={rows[0][:5] if isinstance(rows[0], list) else rows[0]}")
            print(f"     [DEBUG] groups={groups[:5]}")
            # Try to find up/down markers
            sample_signs = []
            for dk in data_keys:
                for row in data[dk][:3]:
                    if isinstance(row, list):
                        for i, cell in enumerate(row):
                            if str(cell).strip() in ["+", "-", "X", " "]:
                                sample_signs.append(f"{dk}[row][{i}]={cell}")
            return f"MI_INDEX: {data_keys}, groups={len(groups)}, signs={sample_signs[:10]}"
        return f"stat={data.get('stat')}"
    test("TWSE MI_INDEX ALLBUT0999", _mi_index)

    # 2. TWSE FMTQIK (每日市場成交資訊)
    def _fmtqik():
        url = "https://openapi.twse.com.tw/v1/exchangeReport/FMTQIK"
        r = requests.get(url, headers=HEADERS, timeout=15)
        data = r.json()
        if data and isinstance(data, list):
            print(f"     [DEBUG] {len(data)} records, keys={list(data[0].keys())}")
            print(f"     [DEBUG] last record={json.dumps(data[-1], ensure_ascii=False)}")
            return f"FMTQIK: {len(data)} records"
        return None
    test("TWSE OpenAPI FMTQIK", _fmtqik)

    # 3. TWSE STOCK_DAY_ALL (count changes ourselves)
    def _stock_day_all():
        url = "https://openapi.twse.com.tw/v1/exchangeReport/STOCK_DAY_ALL"
        r = requests.get(url, headers=HEADERS, timeout=15)
        data = r.json()
        if data and isinstance(data, list) and len(data) > 50:
            first = data[0]
            print(f"     [DEBUG] {len(data)} stocks, keys={list(first.keys())}")
            print(f"     [DEBUG] sample={json.dumps(first, ensure_ascii=False)}")
            # Try to count up/down
            up = down = flat = 0
            change_key = None
            for k in first.keys():
                if "change" in k.lower() or "漲跌" in k:
                    change_key = k
                    break
            if change_key:
                for item in data:
                    try:
                        val = float(str(item.get(change_key, "0")).replace(",", ""))
                        if val > 0: up += 1
                        elif val < 0: down += 1
                        else: flat += 1
                    except:
                        pass
                return f"STOCK_DAY_ALL: {len(data)} stocks, change_key='{change_key}', up={up} down={down} flat={flat}"
            return f"STOCK_DAY_ALL: {len(data)} stocks, no change key found in {list(first.keys())}"
        return None
    test("TWSE OpenAPI STOCK_DAY_ALL", _stock_day_all)

    # 4. TPEx breadth
    def _tpex():
        url = "https://www.tpex.org.tw/web/stock/aftertrading/market_highlight/highlight_result.php"
        params = {"l": "zh-tw", "d": twse_roc_date, "o": "json"}
        r = requests.get(url, params=params, headers=HEADERS, timeout=15)
        data = r.json()
        print(f"     [DEBUG] keys={list(data.keys())}")
        if "reportList" in data:
            print(f"     [DEBUG] reportList={data['reportList'][:5]}")
            return f"TPEx reportList: {len(data['reportList'])} items"
        elif "tables" in data:
            for t in data["tables"][:3]:
                print(f"     [DEBUG] table title={t.get('title')}, data[:3]={t.get('data', [])[:3]}")
            return f"TPEx tables: {len(data['tables'])} tables"
        elif "aaData" in data:
            return f"TPEx aaData: {len(data['aaData'])} rows"
        return f"TPEx: keys={list(data.keys())}, raw={str(data)[:200]}"
    test("TPEx market_highlight", _tpex)

    # 5. TPEx daily trading index
    def _tpex_st41():
        url = "https://www.tpex.org.tw/web/stock/aftertrading/daily_trading_index/st41_result.php"
        params = {"l": "zh-tw", "d": twse_roc_date, "o": "json"}
        r = requests.get(url, params=params, headers=HEADERS, timeout=15)
        data = r.json()
        if "aaData" in data and data["aaData"]:
            rows = data["aaData"]
            last = rows[-1]
            print(f"     [DEBUG] {len(rows)} rows, last={last}")
            return f"TPEx st41: {len(rows)} rows, last_row_len={len(last)}"
        return f"keys={list(data.keys())}"
    test("TPEx daily_trading_index st41", _tpex_st41)


# ============================================================
def test_ratio():
    """測試融資維持率所有來源"""
    print("\n" + "="*60)
    print("融資維持率 (Margin Maintenance Ratio)")
    print("="*60)

    # 1. TWSE MI_MARGN creditList (should have 融資金額 and 擔保品市值)
    def _mi_margn_credit():
        url = "https://www.twse.com.tw/exchangeReport/MI_MARGN"
        params = {"response": "json", "date": twse_date, "selectType": "MS"}
        r = requests.get(url, params=params, headers=HEADERS, timeout=15)
        data = r.json()
        if data.get("stat") == "OK":
            cf = data.get("creditFields", [])
            cl = data.get("creditList", [])
            print(f"     [DEBUG] creditFields={cf}")
            for row in cl:
                print(f"     [DEBUG] row={row}")
            # Try to calculate ratio
            margin_amt = None
            collateral = None
            balance_idx = None
            for i, fn in enumerate(cf):
                if "今日" in str(fn) and "餘額" in str(fn):
                    balance_idx = i
            for row in cl:
                name = str(row[0]).strip()
                if "融資金額" in name and balance_idx:
                    margin_amt = float(str(row[balance_idx]).replace(",", ""))
                elif "擔保" in name and balance_idx:
                    collateral = float(str(row[balance_idx]).replace(",", ""))
            if margin_amt and collateral:
                ratio = round(collateral / margin_amt * 100, 1)
                return f"MI_MARGN: ratio={ratio}%, margin_amt={margin_amt}, collateral={collateral}"
            return f"MI_MARGN: fields={cf}, rows={len(cl)}, margin_amt={margin_amt}, collateral={collateral}"
        return f"stat={data.get('stat')}"
    test("TWSE MI_MARGN creditList", _mi_margn_credit)

    # 2. FinMind margin data
    def _finmind_margin():
        url = "https://api.finmindtrade.com/api/v4/data"
        params = {"dataset": "TaiwanStockTotalMarginPurchaseShortSale",
                  "start_date": trade_date, "end_date": trade_date}
        token = os.environ.get("FINMIND_TOKEN", "")
        if token:
            params["token"] = token
        r = requests.get(url, params=params, headers=HEADERS, timeout=15)
        data = r.json()
        if data.get("status") == 200 and data.get("data"):
            rows = data["data"]
            margin_amt = None
            margin_bal = None
            for row in rows:
                name = row.get("name", "")
                if "融資金額" in name:
                    margin_amt = row.get("TodayBalance")
                elif "融資" in name and "金額" not in name:
                    margin_bal = row.get("TodayBalance")
            return f"FinMind: margin_amt={margin_amt} (仟元), margin_bal={margin_bal} (張)"
        return None
    test("FinMind margin data for ratio calc", _finmind_margin)

    # 3. Goodinfo
    def _goodinfo():
        url = "https://goodinfo.tw/tw/StockMarginList.asp"
        h = {**HEADERS, 'Referer': 'https://goodinfo.tw/tw/index.asp'}
        r = requests.get(url, headers=h, timeout=10)
        if r.status_code == 200:
            text = r.text
            patterns = [
                r'整體.*?維持率.*?(\d{2,3}\.?\d*)\s*%',
                r'融資維持率.*?(\d{2,3}\.?\d*)',
            ]
            for p in patterns:
                match = re.search(p, text)
                if match:
                    return f"Goodinfo: ratio={match.group(1)}%"
            # Check if we got blocked or redirected
            if len(text) < 1000:
                return f"Goodinfo: page too short ({len(text)} chars), possibly blocked"
            return f"Goodinfo: no ratio found in page ({len(text)} chars)"
        return f"Goodinfo: status={r.status_code}"
    test("Goodinfo margin ratio", _goodinfo)

    # 4. MacroMicro
    def _macromicro():
        url = "https://www.macromicro.me/charts/data/53117"
        h = {**HEADERS, 'Accept': 'application/json',
             'Referer': 'https://www.macromicro.me/charts/53117/taiwan-taiex-maintenance-margin'}
        r = requests.get(url, headers=h, timeout=10)
        if r.status_code == 200:
            data = r.json()
            series = data.get("data", {}).get("series", [])
            if series:
                points = series[0].get("data", [])
                if points:
                    latest = points[-1]
                    return f"MacroMicro: latest={latest}"
            return f"MacroMicro: response keys={list(data.keys())}"
        return f"MacroMicro: status={r.status_code}"
    test("MacroMicro chart 53117", _macromicro)

    # 5. cnyes margin maintenance
    def _cnyes_margin():
        url = "https://ws.api.cnyes.com/ws/api/v1/charting/history"
        params = {"resolution": "D", "symbol": "TWS:MARGIN_MAINTENANCE_RATIO:STOCK", "quote": 1}
        r = requests.get(url, params=params, headers=HEADERS, timeout=10)
        data = r.json()
        if data.get("data", {}).get("quote"):
            q = data["data"]["quote"]
            val = q.get("6") or q.get("closePrice")
            return f"cnyes margin ratio: val={val}, quote_keys={list(q.keys())}"
        if data.get("data", {}).get("c"):
            closes = data["data"]["c"]
            return f"cnyes margin ratio: latest={closes[-1]}"
        return None
    test("cnyes TWS:MARGIN_MAINTENANCE_RATIO:STOCK", _cnyes_margin)


# ============================================================
def test_foreign():
    """測試外資排行所有來源"""
    print("\n" + "="*60)
    print("外資排行 (Foreign Top 10)")
    print("="*60)

    # 1. TWSE TWT38U
    def _twt38u():
        url = "https://www.twse.com.tw/fund/TWT38U"
        params = {"response": "json", "date": twse_date}
        r = requests.get(url, params=params, headers=HEADERS, timeout=15)
        data = r.json()
        if data.get("stat") == "OK" and data.get("data"):
            rows = data["data"]
            fields = data.get("fields", [])
            print(f"     [DEBUG] fields={fields}")
            print(f"     [DEBUG] first 3 rows:")
            for row in rows[:3]:
                print(f"       {row}")
            return f"TWT38U: {len(rows)} rows"
        return f"stat={data.get('stat')}"
    test("TWSE TWT38U", _twt38u)

    # 2. TWSE T86
    def _t86():
        url = "https://www.twse.com.tw/fund/T86"
        params = {"response": "json", "date": twse_date, "selectType": "ALLBUT0999"}
        r = requests.get(url, params=params, headers=HEADERS, timeout=15)
        data = r.json()
        if data.get("stat") == "OK" and data.get("data"):
            rows = data["data"]
            fields = data.get("fields", [])
            print(f"     [DEBUG] fields={fields}")
            print(f"     [DEBUG] first 3 rows:")
            for row in rows[:3]:
                print(f"       {row}")
            return f"T86: {len(rows)} rows"
        return f"stat={data.get('stat')}"
    test("TWSE T86", _t86)

    # 3. OpenAPI TWT38U
    def _openapi_twt38u():
        url = "https://openapi.twse.com.tw/v1/fund/TWT38U"
        r = requests.get(url, headers=HEADERS, timeout=15)
        data = r.json()
        if data and isinstance(data, list):
            print(f"     [DEBUG] {len(data)} records, keys={list(data[0].keys())}")
            print(f"     [DEBUG] first={json.dumps(data[0], ensure_ascii=False)[:200]}")
            return f"OpenAPI TWT38U: {len(data)} records"
        return None
    test("TWSE OpenAPI TWT38U", _openapi_twt38u)


# ============================================================
if __name__ == "__main__":
    section = "all"
    if len(sys.argv) > 1:
        if sys.argv[-1] in ["dxy", "margin", "breadth", "ratio", "foreign", "all"]:
            section = sys.argv[-1]
        elif "--section" in sys.argv:
            idx = sys.argv.index("--section")
            if idx + 1 < len(sys.argv):
                section = sys.argv[idx + 1]

    print(f"🔍 API 來源測試工具")
    print(f"今日: {today}, 交易日: {trade_date}, TWSE格式: {twse_date}")
    print(f"測試區段: {section}")

    if section in ["all", "dxy"]:
        test_dxy()
    if section in ["all", "margin"]:
        test_margin()
    if section in ["all", "breadth"]:
        test_breadth()
    if section in ["all", "ratio"]:
        test_ratio()
    if section in ["all", "foreign"]:
        test_foreign()

    print(f"\n{'='*60}")
    print(f"📊 結果: ✅ {passed} 通過, ❌ {failed} 失敗")
    print(f"{'='*60}")
