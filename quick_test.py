#!/usr/bin/env python3
"""
快速 API 測試腳本 - 在 GitHub Actions 上跑，驗證每個來源
用法: python quick_test.py
"""
import requests
import re
import time
from datetime import datetime, timedelta

UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
today = datetime.now().strftime("%Y%m%d")
passed = 0
failed = 0

def test(name, func):
    global passed, failed
    try:
        r = func()
        if r:
            print(f"  ✅ {name}: {str(r)[:200]}")
            passed += 1
        else:
            print(f"  ❌ {name}: 無資料")
            failed += 1
    except Exception as e:
        print(f"  ❌ {name}: {e}")
        failed += 1


# ========== VIX ==========
print("\n😱 VIX 指數")
print("=" * 50)

def t_fred_vix():
    end = datetime.now().strftime("%Y-%m-%d")
    start = (datetime.now() - timedelta(days=30)).strftime("%Y-%m-%d")
    r = requests.get(f"https://fred.stlouisfed.org/graph/fredgraph.csv?id=VIXCLS&cosd={start}&coed={end}",
                     headers={"User-Agent": UA}, timeout=15)
    lines = r.text.strip().split("\n")
    valid = [l for l in lines[1:] if len(l.split(",")) >= 2 and l.split(",")[1].strip() not in ["", "."]]
    if valid:
        last = valid[-1].split(",")
        return f"Date={last[0]} Value={last[1]} ({len(valid)} valid rows)"
    return None
test("FRED VIXCLS", t_fred_vix)

def t_stooq_vix():
    r = requests.get("https://stooq.com/q/d/l/?s=^vix&i=d", headers={"User-Agent": UA}, timeout=15)
    lines = r.text.strip().split("\n")
    if len(lines) > 1 and "Date" in lines[0]:
        last = lines[-1].split(",")
        return f"Date={last[0]} Close={last[4]} ({len(lines)-1} rows)"
    return None
test("Stooq ^vix", t_stooq_vix)

def t_google_finance_vix():
    r = requests.get("https://www.google.com/finance/quote/VIX:INDEXCBOE",
                     headers={"User-Agent": UA, "Accept-Language": "en-US,en;q=0.9"}, timeout=12)
    if r.status_code == 200:
        m = re.search(r'data-last-price="([\d.]+)"', r.text)
        if m:
            return f"VIX={m.group(1)}"
    return None
test("Google Finance VIX:INDEXCBOE", t_google_finance_vix)


# ========== US10Y ==========
print("\n🏛️ 美國10年期公債殖利率")
print("=" * 50)

def t_fred_10y():
    end = datetime.now().strftime("%Y-%m-%d")
    start = (datetime.now() - timedelta(days=30)).strftime("%Y-%m-%d")
    r = requests.get(f"https://fred.stlouisfed.org/graph/fredgraph.csv?id=DGS10&cosd={start}&coed={end}",
                     headers={"User-Agent": UA}, timeout=15)
    lines = r.text.strip().split("\n")
    valid = [l for l in lines[1:] if len(l.split(",")) >= 2 and l.split(",")[1].strip() not in ["", "."]]
    if valid:
        last = valid[-1].split(",")
        return f"Date={last[0]} Yield={last[1]}% ({len(valid)} valid rows)"
    return None
test("FRED DGS10", t_fred_10y)

def t_stooq_10y():
    r = requests.get("https://stooq.com/q/d/l/?s=10usy.b&i=d", headers={"User-Agent": UA}, timeout=15)
    lines = r.text.strip().split("\n")
    if len(lines) > 1 and "Date" in lines[0]:
        last = lines[-1].split(",")
        return f"Date={last[0]} Close={last[4]} ({len(lines)-1} rows)"
    return None
test("Stooq 10usy.b", t_stooq_10y)

def t_treasury_csv():
    year = datetime.now().strftime("%Y")
    url = f"https://home.treasury.gov/resource-center/data-chart-center/interest-rates/daily-treasury-rates.csv/all/{year}?type=daily_treasury_yield_curve&field_tdr_date_value={year}&page&_format=csv"
    r = requests.get(url, headers={"User-Agent": UA}, timeout=15)
    lines = r.text.strip().split("\n")
    header = lines[0].split(",")
    yr10_idx = None
    for i, h in enumerate(header):
        if "10" in h and ("yr" in h.lower() or "year" in h.lower()):
            yr10_idx = i
            break
    if yr10_idx and len(lines) > 1:
        last = lines[-1].split(",")
        return f"Date={last[0]} 10Yr={last[yr10_idx]}% ({len(lines)-1} rows, col_idx={yr10_idx})"
    return f"header={header[:5]}, 10Yr_idx={yr10_idx}"
test("US Treasury CSV", t_treasury_csv)

def t_google_finance_tnx():
    r = requests.get("https://www.google.com/finance/quote/TNX:INDEXCBOE",
                     headers={"User-Agent": UA, "Accept-Language": "en-US,en;q=0.9"}, timeout=12)
    if r.status_code == 200:
        m = re.search(r'data-last-price="([\d.]+)"', r.text)
        if m:
            val = float(m.group(1))
            return f"TNX={val} → Yield={val/10}%"
    return None
test("Google Finance TNX:INDEXCBOE", t_google_finance_tnx)


# ========== 美元指數 (DXY) ==========
print("\n💵 美元指數 (DXY)")
print("=" * 50)

def t_stooq_dxy():
    r = requests.get("https://stooq.com/q/d/l/?s=dxy.f&i=d", headers={"User-Agent": UA}, timeout=15)
    lines = r.text.strip().split("\n")
    if len(lines) > 1 and "Date" in lines[0]:
        last = lines[-1].split(",")
        return f"Date={last[0]} Close={last[4]} ({len(lines)-1} rows)"
    return None
test("Stooq dxy.f", t_stooq_dxy)

def t_fred_dtwexbgs():
    end = datetime.now().strftime("%Y-%m-%d")
    start = (datetime.now() - timedelta(days=45)).strftime("%Y-%m-%d")
    r = requests.get(f"https://fred.stlouisfed.org/graph/fredgraph.csv?id=DTWEXBGS&cosd={start}&coed={end}",
                     headers={"User-Agent": UA}, timeout=15)
    lines = r.text.strip().split("\n")
    valid = [l for l in lines[1:] if len(l.split(",")) >= 2 and l.split(",")[1].strip() not in ["", "."]]
    if valid:
        last = valid[-1].split(",")
        return f"Date={last[0]} Value={last[1]} ({len(valid)} valid rows)"
    return None
test("FRED DTWEXBGS (美元貿易加權)", t_fred_dtwexbgs)


# ========== 日圓匯率 (USD/JPY) ==========
print("\n💴 日圓匯率 (USD/JPY)")
print("=" * 50)

def t_fred_usdjpy():
    end = datetime.now().strftime("%Y-%m-%d")
    start = (datetime.now() - timedelta(days=45)).strftime("%Y-%m-%d")
    r = requests.get(f"https://fred.stlouisfed.org/graph/fredgraph.csv?id=DEXJPUS&cosd={start}&coed={end}",
                     headers={"User-Agent": UA}, timeout=15)
    lines = r.text.strip().split("\n")
    valid = [l for l in lines[1:] if len(l.split(",")) >= 2 and l.split(",")[1].strip() not in ["", "."]]
    if valid:
        last = valid[-1].split(",")
        return f"Date={last[0]} Rate={last[1]} ({len(valid)} valid rows)"
    return None
test("FRED DEXJPUS", t_fred_usdjpy)

def t_stooq_usdjpy():
    r = requests.get("https://stooq.com/q/d/l/?s=usdjpy&i=d", headers={"User-Agent": UA}, timeout=15)
    lines = r.text.strip().split("\n")
    if len(lines) > 1 and "Date" in lines[0]:
        last = lines[-1].split(",")
        return f"Date={last[0]} Close={last[4]} ({len(lines)-1} rows)"
    return None
test("Stooq usdjpy", t_stooq_usdjpy)

def t_google_finance_usdjpy():
    r = requests.get("https://www.google.com/finance/quote/USD-JPY",
                     headers={"User-Agent": UA, "Accept-Language": "en-US,en;q=0.9"}, timeout=12)
    if r.status_code == 200:
        m = re.search(r'data-last-price="([\d.]+)"', r.text)
        if m:
            return f"USD/JPY={m.group(1)}"
    return None
test("Google Finance USD-JPY", t_google_finance_usdjpy)


# ========== 漲跌家數 ==========
print("\n📊 漲跌家數")
print("=" * 50)

def t_mi_index():
    r = requests.get("https://www.twse.com.tw/exchangeReport/MI_INDEX",
                     params={"response": "json", "date": today, "type": "ALLBUT0999"},
                     headers={"User-Agent": UA, "Accept-Language": "zh-TW"}, timeout=15)
    data = r.json()
    stat = data.get("stat")
    data_keys = [k for k in data.keys() if k.startswith("data")]
    if stat == "OK" and data_keys:
        biggest = max(data_keys, key=lambda k: len(data[k]) if isinstance(data[k], list) else 0)
        rows = data[biggest]
        up = down = flat = 0
        for row in rows:
            if isinstance(row, list):
                for cell in row:
                    s = str(cell).strip()
                    if s == "+":
                        up += 1
                        break
                    elif s == "-":
                        down += 1
                        break
                    elif s == "X":
                        flat += 1
                        break
        return f"stat=OK, {biggest}={len(rows)} rows, 漲{up} 跌{down} 平{flat}"
    return f"stat={stat}, data_keys={data_keys}"
test("TWSE MI_INDEX", t_mi_index)

def t_stock_day_all():
    r = requests.get("https://openapi.twse.com.tw/v1/exchangeReport/STOCK_DAY_ALL",
                     headers={"User-Agent": UA}, timeout=15)
    data = r.json()
    if isinstance(data, list) and len(data) > 0:
        first = data[0]
        change_key = None
        for k in first.keys():
            if "change" in k.lower() or "漲跌" in k:
                change_key = k
                break
        up = down = flat = 0
        if change_key:
            for item in data:
                try:
                    val = float(str(item.get(change_key, "")).replace(",", ""))
                    if val > 0: up += 1
                    elif val < 0: down += 1
                    else: flat += 1
                except: pass
        return f"{len(data)} stocks, change_key={change_key}, 漲{up} 跌{down} 平{flat}"
    return None
test("TWSE STOCK_DAY_ALL", t_stock_day_all)


# ========== 外資買賣超 ==========
print("\n🌍 外資買賣超排行")
print("=" * 50)

def t_twt38u():
    r = requests.get("https://www.twse.com.tw/fund/TWT38U",
                     params={"response": "json", "date": today},
                     headers={"User-Agent": UA, "Accept-Language": "zh-TW"}, timeout=15)
    data = r.json()
    stat = data.get("stat")
    if stat == "OK" and "data" in data:
        rows = data["data"]
        if rows and len(rows) > 0:
            sample = rows[0]
            # 確認欄位位置
            return f"stat=OK, {len(rows)} rows, sample={sample[:4]}"
    return f"stat={stat}"
test("TWSE TWT38U (外資排行)", t_twt38u)

def t_t86():
    r = requests.get("https://www.twse.com.tw/fund/T86",
                     params={"response": "json", "date": today, "selectType": "ALLBUT0999"},
                     headers={"User-Agent": UA, "Accept-Language": "zh-TW"}, timeout=15)
    data = r.json()
    stat = data.get("stat")
    if stat == "OK" and "data" in data:
        return f"stat=OK, {len(data['data'])} rows"
    return f"stat={stat}"
test("TWSE T86 (外資備用)", t_t86)


# ========== KGI ==========
print("\n🏢 KGI 凱基證券")
print("=" * 50)

def t_kgi():
    kgi_headers = {
        "User-Agent": UA,
        "Accept": "text/html,application/xhtml+xml",
        "Accept-Language": "zh-TW,zh;q=0.9",
        "Referer": "https://www.kgi.com.tw/zh-tw/product-market/stock-market-overview/tw-stock-market",
    }
    r = requests.get("https://www.kgi.com.tw/zh-tw/product-market/stock-market-overview/tw-stock-market/tw-stock-market-detail",
                     params={"a": "B658010E71E243C4A1D6B5F7BE914BDC", "b": "5D48401A7CE148CD8ABAC965F9B5AFBF"},
                     headers=kgi_headers, timeout=15)
    if r.status_code == 200:
        text = r.text
        has_margin = "融資" in text
        has_ratio = "維持率" in text
        has_breadth = "上漲" in text or "漲" in text
        return f"HTTP 200, len={len(text)}, 融資={has_margin}, 維持率={has_ratio}, 漲跌={has_breadth}"
    return f"HTTP {r.status_code}"
test("KGI 大盤動態", t_kgi)


# ========== 結果 ==========
print(f"\n{'='*50}")
print(f"📊 結果: ✅ {passed} 通過, ❌ {failed} 失敗")
print(f"{'='*50}")
