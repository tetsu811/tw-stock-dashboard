#!/usr/bin/env python3
"""
初始化 history_cache.json — 從多個來源抓取30天歷史數據
只需要跑一次，之後每天的 GitHub Actions 會自動累積

用法: python seed_history.py
"""
import requests
import json
import math
import os
import time
from datetime import datetime, timedelta

UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
CACHE_PATH = os.path.join(SCRIPT_DIR, "history_cache.json")


def fetch_fred(series_id, days=50):
    """從 FRED 取得歷史資料"""
    end = datetime.now().strftime("%Y-%m-%d")
    start = (datetime.now() - timedelta(days=days)).strftime("%Y-%m-%d")
    url = f"https://fred.stlouisfed.org/graph/fredgraph.csv?id={series_id}&cosd={start}&coed={end}"
    try:
        r = requests.get(url, headers={"User-Agent": UA}, timeout=20)
        if r.status_code == 200:
            points = []
            for line in r.text.strip().split("\n")[1:]:
                fields = line.strip().split(",")
                if len(fields) >= 2:
                    d, v = fields[0].strip(), fields[1].strip()
                    if v and v != "." and v != "":
                        try:
                            points.append({"date": d, "close": round(float(v), 4)})
                        except ValueError:
                            pass
            return points
    except Exception as e:
        print(f"  FRED {series_id}: {e}")
    return []


def fetch_stooq(symbol, days=50):
    """從 Stooq 取得歷史資料"""
    try:
        r = requests.get(f"https://stooq.com/q/d/l/?s={symbol}&i=d",
                         headers={"User-Agent": UA}, timeout=20)
        if r.status_code == 200:
            lines = r.text.strip().split("\n")
            if len(lines) > 1 and "Date" in lines[0]:
                points = []
                for line in lines[1:]:
                    fields = line.strip().split(",")
                    if len(fields) >= 5:
                        try:
                            points.append({"date": fields[0].strip(), "close": round(float(fields[4].strip()), 4)})
                        except (ValueError, IndexError):
                            pass
                return points[-days:] if points else []
    except Exception as e:
        print(f"  Stooq {symbol}: {e}")
    return []


def compute_dxy_from_fred():
    """從 FRED 6 個匯率計算 DXY"""
    series = ["DEXUSEU", "DEXJPUS", "DEXUSUK", "DEXCAUS", "DEXSDUS", "DEXSZUS"]
    all_data = {}
    end = datetime.now().strftime("%Y-%m-%d")
    start = (datetime.now() - timedelta(days=55)).strftime("%Y-%m-%d")

    for sid in series:
        url = f"https://fred.stlouisfed.org/graph/fredgraph.csv?id={sid}&cosd={start}&coed={end}"
        try:
            r = requests.get(url, headers={"User-Agent": UA}, timeout=20)
            if r.status_code == 200:
                data_map = {}
                for line in r.text.strip().split("\n")[1:]:
                    fields = line.strip().split(",")
                    if len(fields) >= 2:
                        d, v = fields[0].strip(), fields[1].strip()
                        if v and v != "." and v != "":
                            try:
                                data_map[d] = float(v)
                            except ValueError:
                                pass
                all_data[sid] = data_map
                print(f"    {sid}: {len(data_map)} 筆")
            else:
                print(f"    {sid}: HTTP {r.status_code}")
                return []
        except Exception as e:
            print(f"    {sid}: {e}")
            return []
        time.sleep(0.5)

    if not all(sid in all_data and all_data[sid] for sid in series):
        return []

    common_dates = sorted(set(all_data["DEXUSEU"].keys()))
    results = []
    for date in common_dates:
        vals = {sid: all_data[sid].get(date) for sid in series}
        if all(v is not None and v > 0 for v in vals.values()):
            dxy = 50.14348112 * (
                math.pow(vals["DEXUSEU"], -0.576) *
                math.pow(vals["DEXJPUS"], 0.136) *
                math.pow(vals["DEXUSUK"], -0.119) *
                math.pow(vals["DEXCAUS"], 0.091) *
                math.pow(vals["DEXSDUS"], 0.042) *
                math.pow(vals["DEXSZUS"], 0.036)
            )
            results.append({"date": date, "close": round(dxy, 2)})

    return results


def main():
    cache = {}

    # VIX
    print("\n😱 VIX 歷史")
    vix = fetch_fred("VIXCLS", 50)
    if not vix:
        vix = fetch_stooq("^vix", 50)
    cache["vix"] = vix[-35:] if vix else []
    print(f"  → {len(cache['vix'])} 筆")

    # US10Y
    print("\n🏛️ US10Y 歷史")
    us10y = fetch_fred("DGS10", 50)
    if not us10y:
        us10y = fetch_stooq("10usy.b", 50)
    cache["us10y"] = us10y[-35:] if us10y else []
    print(f"  → {len(cache['us10y'])} 筆")

    # DXY (computed)
    print("\n💵 DXY 歷史 (從 FRED 匯率計算)")
    dxy = compute_dxy_from_fred()
    if not dxy:
        dxy = fetch_stooq("dx.f", 50)
    cache["usd_index"] = dxy[-35:] if dxy else []
    print(f"  → {len(cache['usd_index'])} 筆")

    # JPY
    print("\n💴 USD/JPY 歷史")
    jpy = fetch_stooq("usdjpy", 50)
    if not jpy:
        jpy = fetch_fred("DEXJPUS", 50)
    cache["jpy_rate"] = jpy[-35:] if jpy else []
    print(f"  → {len(cache['jpy_rate'])} 筆")

    # Save
    cache["last_updated"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    cache["seeded"] = True

    with open(CACHE_PATH, 'w', encoding='utf-8') as f:
        json.dump(cache, f, ensure_ascii=False, indent=2)

    print(f"\n✅ history_cache.json 已建立!")
    print(f"   VIX: {len(cache['vix'])} 筆")
    print(f"   US10Y: {len(cache['us10y'])} 筆")
    print(f"   DXY: {len(cache['usd_index'])} 筆")
    print(f"   JPY: {len(cache['jpy_rate'])} 筆")


if __name__ == "__main__":
    main()
