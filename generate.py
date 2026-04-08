#!/usr/bin/env python3
"""
å°ç£è¡å¸æ¯æ¥æ°ç¥åè¡¨æ¿ - ä¸»ç¨å¼
ç¨æ³:
    python generate.py              # æåä»å¤©çè³æä¸¦çæåè¡¨æ¿
    python generate.py 20260320     # æå®æ¥æ
    python generate.py --output /path/to/output  # æå®è¼¸åºç®é
"""

import os
import sys
import json
import argparse
from datetime import datetime
from jinja2 import Template

# å å¥ç¶åç®éå°è·¯å¾
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from data_fetcher import fetch_all_data


def format_amount(value, unit="å"):
    """å°éé¡æ ¼å¼åçºååæè¬åé¡¯ç¤º"""
    if value is None:
        return "N/A", ""
    abs_val = abs(value)
    sign = "+" if value > 0 else "" if value < 0 else ""

    if abs_val >= 1e8:  # å
        return f"{sign}{value / 1e8:,.2f}", "åå"
    elif abs_val >= 1e4:  # è¬
        return f"{sign}{value / 1e4:,.1f}", "è¬å"
    else:
        return f"{sign}{value:,.0f}", unit


def format_number(value, decimals=0):
    """æ ¼å¼åæ¸å­å éè"""
    if value is None:
        return "N/A"
    if decimals > 0:
        return f"{value:,.{decimals}f}"
    return f"{value:,.0f}"


def format_shares(value):
    """æ ¼å¼åè¡æ¸/å¼µæ¸"""
    if value is None:
        return "N/A"
    abs_val = abs(value)
    sign = "+" if value > 0 else ""
    if abs_val >= 1000:
        return f"{sign}{value / 1000:,.0f}åå¼µ"
    return f"{sign}{value:,.0f}å¼µ"


def get_vix_level(value):
    """VIX ææç­ç´"""
    if value is None:
        return "N/A", "text-secondary", ""
    if value < 15:
        return "æ¥µåº¦æ¨è§", "text-green", "bg-green"
    elif value < 20:
        return "æ­£å¸¸", "text-blue", "background: rgba(59,130,246,0.15); color: var(--accent-blue);"
    elif value < 25:
        return "åé«è­¦æ", "text-yellow", "background: rgba(234,179,8,0.15); color: var(--accent-yellow);"
    elif value < 30:
        return "ææ", "text-red", "bg-red"
    else:
        return "æ¥µåº¦ææ", "text-red", "bg-red"


def prepare_template_data(data):
    """å°åå§è³æè½æçºæ¨¡æ¿éè¦çæ ¼å¼"""
    ctx = {}

    # æ¥æé¡¯ç¤º
    date_str = data["date"]
    try:
        dt = datetime.strptime(date_str, "%Y%m%d")
        weekdays = ["ä¸", "äº", "ä¸", "å", "äº", "å­", "æ¥"]
        ctx["date_display"] = f"{dt.year}/{dt.month:02d}/{dt.day:02d} (ææ{weekdays[dt.weekday()]})"
    except:
        ctx["date_display"] = date_str
    ctx["generated_at"] = data["generated_at"]

    # å æ¬ææ¸
    taiex = data.get("taiex", {})
    ctx["taiex_index"] = format_number(taiex.get("index"), 2) if taiex.get("index") else "N/A"
    ctx["taiex_change"] = taiex.get("change")
    ctx["taiex_change_display"] = format_number(abs(taiex["change"]), 2) if taiex.get("change") else "N/A"
    ctx["taiex_change_pct"] = taiex.get("change_pct")

    # æäº¤é
    vol = taiex.get("volume")
    if vol:
        vol_display, vol_unit = format_amount(vol)
        ctx["volume_display"] = f"{vol_display} {vol_unit}"
    else:
        ctx["volume_display"] = "N/A"
    ctx["volume_change_pct"] = taiex.get("volume_change")

    # å°ææè²¨
    futures = data.get("futures", {})
    ctx["futures_close"] = format_number(futures.get("close"), 0) if futures.get("close") else "N/A"
    ctx["futures_change"] = futures.get("change")
    ctx["futures_change_display"] = format_number(abs(futures["change"]), 0) if futures.get("change") else "N/A"
    ctx["futures_change_pct"] = futures.get("change_pct")
    ctx["futures_open"] = format_number(futures.get("open"), 0) if futures.get("open") else "N/A"
    ctx["futures_high"] = format_number(futures.get("high"), 0) if futures.get("high") else "N/A"
    ctx["futures_low"] = format_number(futures.get("low"), 0) if futures.get("low") else "N/A"
    ctx["futures_volume"] = format_number(futures.get("volume"), 0) if futures.get("volume") else "N/A"

    # æç¾è²¨å¹å·®
    taiex_val = data.get("taiex", {}).get("index")
    futures_val = futures.get("close")
    if taiex_val and futures_val:
        spread = futures_val - taiex_val
        ctx["futures_spread"] = spread
        ctx["futures_spread_display"] = format_number(abs(spread), 0)
    else:
        ctx["futures_spread"] = None
        ctx["futures_spread_display"] = "N/A"

    # ä¸å¤§æ³äºº
    inst = data.get("institutional", {})

    # å¤è³ (å«è²·é²/è³£åºæç´°)
    ctx["foreign_net"] = inst.get("foreign_net")
    fn_display, fn_unit = format_amount(inst.get("foreign_net"))
    ctx["foreign_net_display"] = fn_display
    ctx["foreign_net_yi"] = fn_unit

    # å¤è³è²·é²/è³£åºéé¡
    fb_display, fb_unit = format_amount(inst.get("foreign_buy"))
    ctx["foreign_buy_display"] = f"{fb_display} {fb_unit}" if inst.get("foreign_buy") else "N/A"
    fs_display, fs_unit = format_amount(inst.get("foreign_sell"))
    ctx["foreign_sell_display"] = f"{fs_display} {fs_unit}" if inst.get("foreign_sell") else "N/A"

    prev_fn = inst.get("foreign_prev_net")
    if prev_fn is not None:
        pfn_display, pfn_unit = format_amount(prev_fn)
        ctx["foreign_prev_net_display"] = f"{pfn_display} {pfn_unit}"
        ctx["foreign_net_diff"] = (inst.get("foreign_net") or 0) - prev_fn
    else:
        ctx["foreign_prev_net_display"] = None
        ctx["foreign_net_diff"] = 0

    # æä¿¡ (å«è²·é²/è³£åºæç´°)
    ctx["trust_net"] = inst.get("trust_net")
    tn_display, tn_unit = format_amount(inst.get("trust_net"))
    ctx["trust_net_display"] = tn_display
    ctx["trust_net_yi"] = tn_unit

    # æä¿¡è²·é²/è³£åºéé¡
    tb_display, tb_unit = format_amount(inst.get("trust_buy"))
    ctx["trust_buy_display"] = f"{tb_display} {tb_unit}" if inst.get("trust_buy") else "N/A"
    ts_display, ts_unit = format_amount(inst.get("trust_sell"))
    ctx["trust_sell_display"] = f"{ts_display} {ts_unit}" if inst.get("trust_sell") else "N/A"

    prev_tn = inst.get("trust_prev_net")
    if prev_tn is not None:
        ptn_display, ptn_unit = format_amount(prev_tn)
        ctx["trust_prev_net_display"] = f"{ptn_display} {ptn_unit}"
        ctx["trust_net_diff"] = (inst.get("trust_net") or 0) - prev_tn
    else:
        ctx["trust_prev_net_display"] = None
        ctx["trust_net_diff"] = 0

    # èè³
    margin = data.get("margin", {})
    margin_amt = margin.get("margin_balance_amount")
    if margin_amt:
        m_display, m_unit = format_amount(margin_amt)
        ctx["margin_display"] = m_display
        ctx["margin_sub"] = m_unit
    elif margin.get("margin_balance"):
        ctx["margin_display"] = format_number(margin["margin_balance"])
        ctx["margin_sub"] = "å¼µ"
    else:
        ctx["margin_display"] = "N/A"
        ctx["margin_sub"] = ""

    # VIX (å«7å¤©åè¡¨)
    vix = data.get("vix", {})
    vix_val = vix.get("value")
    ctx["vix_value"] = format_number(vix_val, 2) if vix_val else "N/A"
    vix_label, vix_color, vix_bg = get_vix_level(vix_val)
    ctx["vix_label"] = vix_label
    ctx["vix_color_class"] = vix_color
    # ç¨ inline style é¿åæ¨¡æ¿ class åé¡
    ctx["vix_label_style"] = vix_bg if "background" in str(vix_bg) else (
        "background: rgba(34,197,94,0.15); color: #22c55e;" if vix_val and vix_val < 15 else
        "background: rgba(59,130,246,0.15); color: #3b82f6;" if vix_val and vix_val < 20 else
        "background: rgba(234,179,8,0.15); color: #eab308;" if vix_val and vix_val < 25 else
        "background: rgba(239,68,68,0.15); color: #ef4444;" if vix_val else ""
    )

    # VIX è®å
    vix_prev = vix.get("prev_value")
    if vix_val and vix_prev:
        vix_change = round(vix_val - vix_prev, 2)
        ctx["vix_change"] = vix_change
        ctx["vix_change_display"] = format_number(abs(vix_change), 2)
    else:
        ctx["vix_change"] = None
        ctx["vix_change_display"] = None

    # VIX 30å¤©åè¡¨è³æ
    vix_chart = vix.get("chart", [])
    ctx["vix_chart_data"] = json.dumps({
        "labels": [p["date"][-5:] for p in vix_chart],
        "values": [p["close"] for p in vix_chart],
    })

    # æ¼²è·å®¶æ¸
    breadth = data.get("breadth", {})
    ctx["tse_up"] = breadth.get("tse_up")
    ctx["tse_down"] = breadth.get("tse_down")
    ctx["tse_flat"] = breadth.get("tse_flat")
    ctx["otc_up"] = breadth.get("otc_up")
    ctx["otc_down"] = breadth.get("otc_down")
    ctx["otc_flat"] = breadth.get("otc_flat")

    # å¤è³æè¡
    foreign = data.get("foreign_top10", {})
    ctx["top_buy"] = []
    for stock in foreign.get("top_buy", []):
        net = stock.get("net", 0)
        ctx["top_buy"].append({
            "stock_id": stock["stock_id"],
            "stock_name": stock["stock_name"],
            "net_display": format_number(abs(net)),
        })

    ctx["top_sell"] = []
    for stock in foreign.get("top_sell", []):
        net = stock.get("net", 0)
        ctx["top_sell"].append({
            "stock_id": stock["stock_id"],
            "stock_name": stock["stock_name"],
            "net_display": format_number(abs(net)),
        })

    # åè¡¨è³æ
    usd_data = data.get("usd_index", [])
    ctx["usd_chart_data"] = json.dumps({
        "labels": [p["date"][-5:] for p in usd_data],  # MM-DD
        "values": [p["close"] for p in usd_data],
    })
    ctx["usd_latest"] = format_number(usd_data[-1]["close"], 2) if usd_data else None

    jpy_data = data.get("jpy_rate", [])
    ctx["jpy_chart_data"] = json.dumps({
        "labels": [p["date"][-5:] for p in jpy_data],
        "values": [p["close"] for p in jpy_data],
    })
    ctx["jpy_latest"] = format_number(jpy_data[-1]["close"], 2) if jpy_data else None

    # ====== æ°å¢ææ¨ ======

    # èè³ç¶­æç
    mr = data.get("margin_ratio", {})
    mr_val = mr.get("ratio")
    ctx["margin_ratio_value"] = format_number(mr_val, 1) if mr_val else "N/A"
    ctx["margin_ratio_raw"] = mr_val
    if mr_val:
        if mr_val >= 170:
            ctx["margin_ratio_color"] = "text-green"
            ctx["margin_ratio_label"] = "å®å¨"
            ctx["margin_ratio_style"] = "background: rgba(34,197,94,0.15); color: #22c55e;"
        elif mr_val >= 150:
            ctx["margin_ratio_color"] = "text-blue"
            ctx["margin_ratio_label"] = "æ­£å¸¸"
            ctx["margin_ratio_style"] = "background: rgba(59,130,246,0.15); color: #3b82f6;"
        elif mr_val >= 130:
            ctx["margin_ratio_color"] = "text-yellow"
            ctx["margin_ratio_label"] = "è­¦æ"
            ctx["margin_ratio_style"] = "background: rgba(234,179,8,0.15); color: #eab308;"
        else:
            ctx["margin_ratio_color"] = "text-red"
            ctx["margin_ratio_label"] = "å±éª"
            ctx["margin_ratio_style"] = "background: rgba(239,68,68,0.15); color: #ef4444;"
    else:
        ctx["margin_ratio_color"] = ""
        ctx["margin_ratio_label"] = ""
        ctx["margin_ratio_style"] = ""

    # CNN Fear & Greed Index
    cnn = data.get("cnn_fg", {})
    cnn_val = cnn.get("value")
    ctx["cnn_fg_value"] = format_number(cnn_val, 0) if cnn_val else "N/A"
    ctx["cnn_fg_raw"] = cnn_val
    ctx["cnn_fg_label"] = cnn.get("label", "")
    if cnn_val:
        if cnn_val >= 75:
            ctx["cnn_fg_color"] = "text-green"
            ctx["cnn_fg_style"] = "background: rgba(34,197,94,0.15); color: #22c55e;"
        elif cnn_val >= 55:
            ctx["cnn_fg_color"] = "text-cyan"
            ctx["cnn_fg_style"] = "background: rgba(6,182,212,0.15); color: #06b6d4;"
        elif cnn_val >= 45:
            ctx["cnn_fg_color"] = "text-blue"
            ctx["cnn_fg_style"] = "background: rgba(59,130,246,0.15); color: #3b82f6;"
        elif cnn_val >= 25:
            ctx["cnn_fg_color"] = "text-orange"
            ctx["cnn_fg_style"] = "background: rgba(249,115,22,0.15); color: #f97316;"
        else:
            ctx["cnn_fg_color"] = "text-red"
            ctx["cnn_fg_style"] = "background: rgba(239,68,68,0.15); color: #ef4444;"
    else:
        ctx["cnn_fg_color"] = ""
        ctx["cnn_fg_style"] = ""

    cnn_prev = cnn.get("prev_value")
    if cnn_val and cnn_prev:
        ctx["cnn_fg_change"] = round(cnn_val - cnn_prev, 1)
    else:
        ctx["cnn_fg_change"] = None

    # æ¯ç¹å¹£ææèè²ªå©ªææ¸
    crypto = data.get("crypto_fg", {})
    crypto_val = crypto.get("value")
    ctx["crypto_fg_value"] = str(crypto_val) if crypto_val else "N/A"
    ctx["crypto_fg_raw"] = crypto_val
    ctx["crypto_fg_label"] = crypto.get("label", "")
    if crypto_val:
        if crypto_val >= 75:
            ctx["crypto_fg_color"] = "text-green"
            ctx["crypto_fg_style"] = "background: rgba(34,197,94,0.15); color: #22c55e;"
        elif crypto_val >= 55:
            ctx["crypto_fg_color"] = "text-cyan"
            ctx["crypto_fg_style"] = "background: rgba(6,182,212,0.15); color: #06b6d4;"
        elif crypto_val >= 45:
            ctx["crypto_fg_color"] = "text-blue"
            ctx["crypto_fg_style"] = "background: rgba(59,130,246,0.15); color: #3b82f6;"
        elif crypto_val >= 25:
            ctx["crypto_fg_color"] = "text-orange"
            ctx["crypto_fg_style"] = "background: rgba(249,115,22,0.15); color: #f97316;"
        else:
            ctx["crypto_fg_color"] = "text-red"
            ctx["crypto_fg_style"] = "background: rgba(239,68,68,0.15); color: #ef4444;"
    else:
        ctx["crypto_fg_color"] = ""
        ctx["crypto_fg_style"] = ""

    crypto_prev = crypto.get("prev_value")
    if crypto_val and crypto_prev:
        ctx["crypto_fg_change"] = crypto_val - crypto_prev
    else:
        ctx["crypto_fg_change"] = None

    # Put/Call Ratio
    pcr = data.get("pcr", {})
    pcr_val = pcr.get("ratio")
    ctx["pcr_value"] = format_number(pcr_val, 1) if pcr_val else "N/A"
    ctx["pcr_raw"] = pcr_val
    if pcr_val:
        if pcr_val > 100:
            ctx["pcr_color"] = "text-green"
            ctx["pcr_label"] = "åå¤"
            ctx["pcr_style"] = "background: rgba(34,197,94,0.15); color: #22c55e;"
        elif pcr_val == 100:
            ctx["pcr_color"] = "text-blue"
            ctx["pcr_label"] = "ä¸­æ§"
            ctx["pcr_style"] = "background: rgba(59,130,246,0.15); color: #3b82f6;"
        else:
            ctx["pcr_color"] = "text-red"
            ctx["pcr_label"] = "åç©º"
            ctx["pcr_style"] = "background: rgba(239,68,68,0.15); color: #ef4444;"
    else:
        ctx["pcr_color"] = ""
        ctx["pcr_label"] = ""
        ctx["pcr_style"] = ""

    put_oi = pcr.get("put_oi")
    call_oi = pcr.get("call_oi")
    ctx["pcr_put_oi"] = format_number(put_oi) if put_oi else "N/A"
    ctx["pcr_call_oi"] = format_number(call_oi) if call_oi else "N/A"

    # ç¾å 10 å¹´æå¬åµæ®å©ç
    us10y = data.get("us10y", {})
    us10y_val = us10y.get("value")
    ctx["us10y_value"] = format_number(us10y_val, 3) if us10y_val else "N/A"
    us10y_prev = us10y.get("prev_value")
    if us10y_val and us10y_prev:
        ctx["us10y_change"] = round(us10y_val - us10y_prev, 3)
        ctx["us10y_change_display"] = format_number(abs(us10y_val - us10y_prev), 3)
    else:
        ctx["us10y_change"] = None
        ctx["us10y_change_display"] = None

    us10y_chart = us10y.get("chart", [])
    ctx["us10y_chart_data"] = json.dumps({
        "labels": [p["date"][-5:] for p in us10y_chart],
        "values": [p["close"] for p in us10y_chart],
    })

    # ====== èªçåè²·è³£è¶ ======
    dealer_net = inst.get("dealer_net")
    ctx["dealer_net"] = dealer_net
    dn_display, dn_unit = format_amount(dealer_net)
    ctx["dealer_net_display"] = dn_display
    ctx["dealer_net_yi"] = dn_unit

    # ====== ä¸å¤§æ³äººå°æææªå¹³å ======
    foi = data.get("futures_oi", {})
    for key in ["foreign", "trust", "dealer", "total"]:
        item = foi.get(key, {})
        chg = item.get("change")
        oi = item.get("oi")
        ctx[f"foi_{key}_change"] = format_number(chg) if chg is not None else "N/A"
        ctx[f"foi_{key}_change_raw"] = chg
        ctx[f"foi_{key}_oi"] = format_number(oi) if oi is not None else "N/A"
        ctx[f"foi_{key}_oi_raw"] = oi

    # ====== ææ¬è§æ¸¬ææ¨ ======
    senti = data.get("sentiment", {})

    # å¾®å°å¤ç©º
    ctx["micro_sentiment"] = senti.get("micro_sentiment")
    ctx["micro_sentiment_display"] = f"{senti['micro_sentiment']:.2f}%" if senti.get("micro_sentiment") is not None else "N/A"
    ctx["micro_sentiment_prev"] = senti.get("micro_sentiment_prev")
    ctx["micro_sentiment_prev_display"] = f"{senti['micro_sentiment_prev']:.2f}%" if senti.get("micro_sentiment_prev") is not None else "N/A"

    # å°å°å¤ç©º
    ctx["mini_sentiment"] = senti.get("mini_sentiment")
    ctx["mini_sentiment_display"] = f"{senti['mini_sentiment']:.2f}%" if senti.get("mini_sentiment") is not None else "N/A"
    ctx["mini_sentiment_prev"] = senti.get("mini_sentiment_prev")
    ctx["mini_sentiment_prev_display"] = f"{senti['mini_sentiment_prev']:.2f}%" if senti.get("mini_sentiment_prev") is not None else "N/A"

    # PCR å«åæ¥
    ctx["pcr_prev_value"] = format_number(senti.get("pcr_prev"), 1) if senti.get("pcr_prev") else "N/A"

    return ctx


def generate_dashboard(data, output_dir=None):
    """çæ HTML åè¡¨æ¿"""
    # è®åæ¨¡æ¿
    template_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "dashboard_template.html")
    with open(template_path, "r", encoding="utf-8") as f:
        template_str = f.read()

    template = Template(template_str)

    # æºåæ¨¡æ¿è³æ
    ctx = prepare_template_data(data)

    # æ¸²æ
    html = template.render(**ctx)

    # è¼¸åº
    if output_dir is None:
        output_dir = os.path.dirname(os.path.abspath(__file__))

    os.makedirs(output_dir, exist_ok=True)

    # è¼¸åº index.html (åºå®æªåï¼é©åéæç¶²ç«)
    index_path = os.path.join(output_dir, "index.html")
    with open(index_path, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"ð åè¡¨æ¿å·²çæ: {index_path}")

    # åæè¼¸åºå¸¶æ¥æççæ¬ (æ­·å²ç´é)
    date_str = data["date"]
    archive_dir = os.path.join(output_dir, "archive")
    os.makedirs(archive_dir, exist_ok=True)
    archive_path = os.path.join(archive_dir, f"dashboard_{date_str}.html")
    with open(archive_path, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"ð æ­·å²æªæ¡: {archive_path}")

    # è¼¸åº JSON è³æ (æ¹ä¾¿å¶ä»ç¨å¼ä½¿ç¨)
    json_path = os.path.join(output_dir, "latest_data.json")
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2, default=str)
    print(f"ð JSON è³æ: {json_path}")

    return index_path



HISTORY_CACHE_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "history_cache.json")
HISTORY_KEEP_DAYS = 20


def _append_history_point(series, today_iso, value):
    """Append a single data point to a history series, dedupe by date, and trim."""
    if value is None:
        return series
    if not isinstance(series, list):
        series = []
    series = [p for p in series if isinstance(p, dict) and p.get("date") != today_iso]
    try:
        series.append({"date": today_iso, "close": float(value)})
    except (TypeError, ValueError):
        return series
    series.sort(key=lambda p: p.get("date", ""))
    return series[-HISTORY_KEEP_DAYS:]


def merge_and_persist_history(data):
    """Load history_cache.json, append today's metrics, save, and inject back into data."""
    cache_keys = ["vix", "us10y", "usd_index", "jpy_rate",
                  "micro_sentiment", "mini_sentiment", "pcr"]
    cache = {k: [] for k in cache_keys}

    if os.path.exists(HISTORY_CACHE_PATH):
        try:
            with open(HISTORY_CACHE_PATH, "r", encoding="utf-8") as f:
                loaded = json.load(f)
                for k in cache_keys:
                    if isinstance(loaded.get(k), list):
                        cache[k] = loaded[k]
        except Exception as e:
            print(f"  [history] load failed: {e}")

    today_iso = datetime.now().strftime("%Y-%m-%d")

    vix_obj = data.get("vix") or {}
    if isinstance(vix_obj, dict):
        cache["vix"] = _append_history_point(cache["vix"], today_iso, vix_obj.get("value"))

    us10y_obj = data.get("us10y") or {}
    if isinstance(us10y_obj, dict):
        cache["us10y"] = _append_history_point(cache["us10y"], today_iso, us10y_obj.get("value"))

    usd_arr = data.get("usd_index")
    if isinstance(usd_arr, list) and usd_arr:
        latest = usd_arr[-1]
        if isinstance(latest, dict) and latest.get("close") is not None:
            cache["usd_index"] = _append_history_point(
                cache["usd_index"], latest.get("date") or today_iso, latest.get("close"))

    jpy_arr = data.get("jpy_rate")
    if isinstance(jpy_arr, list) and jpy_arr:
        latest = jpy_arr[-1]
        if isinstance(latest, dict) and latest.get("close") is not None:
            cache["jpy_rate"] = _append_history_point(
                cache["jpy_rate"], latest.get("date") or today_iso, latest.get("close"))

    senti = data.get("sentiment") or {}
    if isinstance(senti, dict):
        cache["micro_sentiment"] = _append_history_point(
            cache["micro_sentiment"], today_iso, senti.get("micro_sentiment"))
        cache["mini_sentiment"] = _append_history_point(
            cache["mini_sentiment"], today_iso, senti.get("mini_sentiment"))

    pcr_obj = data.get("pcr") or {}
    if isinstance(pcr_obj, dict):
        cache["pcr"] = _append_history_point(cache["pcr"], today_iso, pcr_obj.get("ratio"))

    cache["last_updated"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    try:
        with open(HISTORY_CACHE_PATH, "w", encoding="utf-8") as f:
            json.dump(cache, f, ensure_ascii=False, indent=2)
        print(f"[history] saved vix={len(cache['vix'])} us10y={len(cache['us10y'])} "
              f"dxy={len(cache['usd_index'])} jpy={len(cache['jpy_rate'])} "
              f"micro={len(cache['micro_sentiment'])} mini={len(cache['mini_sentiment'])} "
              f"pcr={len(cache['pcr'])}")
    except Exception as e:
        print(f"[history] save failed: {e}")

    # Inject historical chart and prev values back into data so templates render properly
    if cache["vix"]:
        if not isinstance(data.get("vix"), dict):
            data["vix"] = {}
        data["vix"]["chart"] = list(cache["vix"])
        if len(cache["vix"]) >= 2 and data["vix"].get("prev_value") in (None, ""):
            data["vix"]["prev_value"] = cache["vix"][-2].get("close")

    if cache["us10y"]:
        if not isinstance(data.get("us10y"), dict):
            data["us10y"] = {}
        data["us10y"]["chart"] = list(cache["us10y"])
        if len(cache["us10y"]) >= 2 and data["us10y"].get("prev_value") in (None, ""):
            data["us10y"]["prev_value"] = cache["us10y"][-2].get("close")

    if cache["usd_index"]:
        data["usd_index"] = list(cache["usd_index"])
    if cache["jpy_rate"]:
        data["jpy_rate"] = list(cache["jpy_rate"])

    if not isinstance(data.get("sentiment"), dict):
        data["sentiment"] = {}
    senti_out = data["sentiment"]
    if len(cache["micro_sentiment"]) >= 2 and senti_out.get("micro_sentiment_prev") in (None, ""):
        senti_out["micro_sentiment_prev"] = cache["micro_sentiment"][-2].get("close")
    if len(cache["mini_sentiment"]) >= 2 and senti_out.get("mini_sentiment_prev") in (None, ""):
        senti_out["mini_sentiment_prev"] = cache["mini_sentiment"][-2].get("close")
    if len(cache["pcr"]) >= 2 and senti_out.get("pcr_prev") in (None, ""):
        senti_out["pcr_prev"] = cache["pcr"][-2].get("close")
    senti_out["micro_sentiment_chart"] = list(cache["micro_sentiment"])
    senti_out["mini_sentiment_chart"] = list(cache["mini_sentiment"])
    senti_out["pcr_chart"] = list(cache["pcr"])

    return data



def main():
    parser = argparse.ArgumentParser(description="å°ç£è¡å¸æ¯æ¥æ°ç¥åè¡¨æ¿çæå¨")
    parser.add_argument("date", nargs="?", default=None, help="æ¥æ YYYYMMDD (é è¨­ä»å¤©)")
    parser.add_argument("--output", "-o", default=None, help="è¼¸åºç®é")
    args = parser.parse_args()

    date_str = args.date or datetime.now().strftime("%Y%m%d")

    # æåè³æ
    data = fetch_all_data(date_str)

    # Persist history (VIX/DXY/JPY/US10Y/sentiment/PCR) and inject prev values + charts
    data = merge_and_persist_history(data)

    # çæåè¡¨æ¿
    output_path = generate_dashboard(data, args.output)

    print(f"\nð å®æï¼è«ç¨çè¦½å¨æé {output_path}")


if __name__ == "__main__":
    main()
