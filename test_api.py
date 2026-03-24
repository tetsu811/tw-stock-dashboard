#!/usr/bin/env python3
"""
API 連線測試腳本
用法：python test_api.py
（可選）設定 FinMind token: export FINMIND_TOKEN="your_token"
"""
import os
import sys
import json
from datetime import datetime

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from data_fetcher import (
    fetch_taiex, fetch_institutional, fetch_foreign_top10,
    fetch_margin_trading, fetch_market_breadth, fetch_taiex_futures,
    fetch_futures_oi, fetch_sentiment_indicators,
    fetch_margin_maintenance_ratio, fetch_put_call_ratio,
    fetch_vix, fetch_usd_index, fetch_jpy_rate,
    fetch_cnn_fear_greed, fetch_crypto_fear_greed, fetch_us10y,
)

date_str = datetime.now().strftime("%Y%m%d")
print(f"\n{'='*60}")
print(f"  台灣股市儀表板 - API 連線測試")
print(f"  測試日期: {date_str}")
print(f"  FinMind Token: {'已設定' if os.environ.get('FINMIND_TOKEN') else '未設定'}")
print(f"{'='*60}\n")

results = {}
errors = []

# 逐一測試各 API
tests = [
    ("加權指數", lambda: fetch_taiex(date_str)),
    ("三大法人買賣超", lambda: fetch_institutional(date_str)),
    ("外資買賣超排行", lambda: fetch_foreign_top10(date_str)),
    ("融資融券餘額", lambda: fetch_margin_trading(date_str)),
    ("漲跌家數", lambda: fetch_market_breadth(date_str)),
    ("台指期貨", lambda: fetch_taiex_futures(date_str)),
    ("三大法人期貨未平倉", lambda: fetch_futures_oi(date_str)),
    ("期權觀測指標", lambda: fetch_sentiment_indicators(date_str)),
    ("融資維持率", lambda: fetch_margin_maintenance_ratio(date_str)),
    ("Put/Call Ratio", lambda: fetch_put_call_ratio(date_str)),
    ("VIX 恐慌指數", lambda: fetch_vix()),
    ("美元指數", lambda: fetch_usd_index()),
    ("日圓匯率", lambda: fetch_jpy_rate()),
    ("CNN 恐慌貪婪", lambda: fetch_cnn_fear_greed()),
    ("加密貨幣恐慌貪婪", lambda: fetch_crypto_fear_greed()),
    ("美國10年公債", lambda: fetch_us10y()),
]

for name, func in tests:
    try:
        data = func()
        results[name] = data
    except Exception as e:
        errors.append((name, str(e)))
        results[name] = {"error": str(e)}

# 顯示結果摘要
print(f"\n{'='*60}")
print("  測試結果摘要")
print(f"{'='*60}\n")

def check_value(label, value, expected_type="number"):
    status = "✅" if value is not None else "❌"
    return f"  {status} {label}: {value}"

# 加權指數
d = results.get("加權指數", {})
print("【加權指數】")
print(check_value("指數", d.get("index")))
print(check_value("漲跌", d.get("change")))
print(check_value("成交金額", d.get("volume")))
print()

# 三大法人
d = results.get("三大法人買賣超", {})
print("【三大法人現貨買賣超】")
print(check_value("外資買賣超", d.get("foreign_net")))
print(check_value("投信買賣超", d.get("trust_net")))
print(check_value("自營商買賣超", d.get("dealer_net")))
print()

# 融資
d = results.get("融資融券餘額", {})
print("【融資融券】")
print(check_value("融資餘額(張)", d.get("margin_balance")))
print(check_value("融資餘額(金額)", d.get("margin_balance_amount")))
print()

# 台指期貨
d = results.get("台指期貨", {})
print("【台指期貨】")
print(check_value("收盤價", d.get("close")))
print(check_value("漲跌", d.get("change")))
print(check_value("成交量", d.get("volume")))
print()

# 三大法人期貨未平倉 (重點修正項目)
d = results.get("三大法人期貨未平倉", {})
print("【三大法人期貨未平倉】⭐ 重點修正")
for key in ["foreign", "trust", "dealer", "total"]:
    item = d.get(key, {})
    label_map = {"foreign": "外資", "trust": "投信", "dealer": "自營", "total": "合計"}
    print(f"  {label_map[key]}: 增減={item.get('change')}, 未平倉={item.get('oi')}")
print()

# 期權觀測指標 (重點修正項目)
d = results.get("期權觀測指標", {})
print("【期權觀測指標】⭐ 重點修正")
print(check_value("微台多空指標", d.get("micro_sentiment")))
print(check_value("微台多空(前日)", d.get("micro_sentiment_prev")))
print(check_value("小台多空指標", d.get("mini_sentiment")))
print(check_value("小台多空(前日)", d.get("mini_sentiment_prev")))
print(check_value("PCR(前日)", d.get("pcr_prev")))
print()

# 融資維持率
d = results.get("融資維持率", {})
print("【融資維持率】⭐ 重點修正")
print(check_value("維持率", d.get("ratio")))
print()

# PCR
d = results.get("Put/Call Ratio", {})
print("【Put/Call Ratio】")
print(check_value("PCR", d.get("ratio")))
print(check_value("Put OI", d.get("put_oi")))
print(check_value("Call OI", d.get("call_oi")))
print()

# 漲跌家數
d = results.get("漲跌家數", {})
print("【漲跌家數】")
print(check_value("上市-漲", d.get("tse_up")))
print(check_value("上市-跌", d.get("tse_down")))
print(check_value("上櫃-漲", d.get("otc_up")))
print(check_value("上櫃-跌", d.get("otc_down")))
print()

# 外資排行
d = results.get("外資買賣超排行", {})
top_buy = d.get("top_buy", [])
top_sell = d.get("top_sell", [])
print("【外資排行】")
if top_buy:
    buy_strs = [f"{s.get('stock_name','')}{s.get('stock_id','')}" for s in top_buy[:3]]
    print(f"  買超前3: {', '.join(buy_strs)}")
else:
    print("  買超前3: ❌ 無資料")
if top_sell:
    sell_strs = [f"{s.get('stock_name','')}{s.get('stock_id','')}" for s in top_sell[:3]]
    print(f"  賣超前3: {', '.join(sell_strs)}")
else:
    print("  賣超前3: ❌ 無資料")
print()

# 國際指標
print("【國際指標】")
d = results.get("VIX 恐慌指數", {})
print(check_value("VIX", d.get("value")))
d = results.get("美元指數", [])
print(check_value("美元指數", d[-1]["close"] if d else None))
d = results.get("日圓匯率", [])
print(check_value("USD/JPY", d[-1]["close"] if d else None))
d = results.get("美國10年公債", {})
print(check_value("US10Y", d.get("value")))
d = results.get("CNN 恐慌貪婪", {})
print(check_value("CNN FG", d.get("value")))
d = results.get("加密貨幣恐慌貪婪", {})
print(check_value("Crypto FG", d.get("value")))
print()

# 錯誤報告
if errors:
    print(f"\n⚠️ 發生 {len(errors)} 個錯誤:")
    for name, err in errors:
        print(f"  ❌ {name}: {err}")

# 統計
total = len(tests)
ok_count = sum(1 for name in results if not isinstance(results[name], dict) or "error" not in results[name])
print(f"\n{'='*60}")
print(f"  測試完成: {ok_count}/{total} 項目成功")
print(f"{'='*60}\n")
