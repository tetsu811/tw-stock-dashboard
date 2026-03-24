#!/usr/bin/env python3
"""
台灣股市每日戰略儀表板 - 主程式
用法:
    python generate.py              # 抓取今天的資料並生成儀表板
    python generate.py 20260320     # 指定日期
    python generate.py --output /path/to/output  # 指定輸出目錄
"""

import os
import sys
import json
import argparse
from datetime import datetime
from jinja2 import Template

# 加入當前目錄到路徑
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from data_fetcher import fetch_all_data


def format_amount(value, unit="元"):
    """將金額格式化為億元或萬元顯示"""
    if value is None:
        return "N/A", ""
    abs_val = abs(value)
    sign = "+" if value > 0 else "" if value < 0 else ""

    if abs_val >= 1e8:  # 億
        return f"{sign}{value / 1e8:,.2f}", "億元"
    elif abs_val >= 1e4:  # 萬
        return f"{sign}{value / 1e4:,.1f}", "萬元"
    else:
        return f"{sign}{value:,.0f}", unit


def format_number(value, decimals=0):
    """格式化數字加逗號"""
    if value is None:
        return "N/A"
    if decimals > 0:
        return f"{value:,.{decimals}f}"
    return f"{value:,.0f}"


def format_shares(value):
    """格式化股數/張數"""
    if value is None:
        return "N/A"
    abs_val = abs(value)
    sign = "+" if value > 0 else ""
    if abs_val >= 1000:
        return f"{sign}{value / 1000:,.0f}千張"
    return f"{sign}{value:,.0f}張"


def get_vix_level(value):
    """VIX 恐慌等級"""
    if value is None:
        return "N/A", "text-secondary", ""
    if value < 15:
        return "極度樂觀", "text-green", "bg-green"
    elif value < 20:
        return "正常", "text-blue", "background: rgba(59,130,246,0.15); color: var(--accent-blue);"
    elif value < 25:
        return "偏高警戒", "text-yellow", "background: rgba(234,179,8,0.15); color: var(--accent-yellow);"
    elif value < 30:
        return "恐慌", "text-red", "bg-red"
    else:
        return "極度恐慌", "text-red", "bg-red"


def prepare_template_data(data):
    """將原始資料轉換為模板需要的格式"""
    ctx = {}

    # 日期顯示
    date_str = data["date"]
    try:
        dt = datetime.strptime(date_str, "%Y%m%d")
        weekdays = ["一", "二", "三", "四", "五", "六", "日"]
        ctx["date_display"] = f"{dt.year}/{dt.month:02d}/{dt.day:02d} (星期{weekdays[dt.weekday()]})"
    except:
        ctx["date_display"] = date_str
    ctx["generated_at"] = data["generated_at"]

    # 加權指數
    taiex = data.get("taiex", {})
    ctx["taiex_index"] = format_number(taiex.get("index"), 2) if taiex.get("index") else "N/A"
    ctx["taiex_change"] = taiex.get("change")
    ctx["taiex_change_display"] = format_number(abs(taiex["change"]), 2) if taiex.get("change") else "N/A"
    ctx["taiex_change_pct"] = taiex.get("change_pct")

    # 成交量
    vol = taiex.get("volume")
    if vol:
        vol_display, vol_unit = format_amount(vol)
        ctx["volume_display"] = f"{vol_display} {vol_unit}"
    else:
        ctx["volume_display"] = "N/A"
    ctx["volume_change_pct"] = taiex.get("volume_change")

    # 台指期貨
    futures = data.get("futures", {})
    ctx["futures_close"] = format_number(futures.get("close"), 0) if futures.get("close") else "N/A"
    ctx["futures_change"] = futures.get("change")
    ctx["futures_change_display"] = format_number(abs(futures["change"]), 0) if futures.get("change") else "N/A"
    ctx["futures_change_pct"] = futures.get("change_pct")
    ctx["futures_open"] = format_number(futures.get("open"), 0) if futures.get("open") else "N/A"
    ctx["futures_high"] = format_number(futures.get("high"), 0) if futures.get("high") else "N/A"
    ctx["futures_low"] = format_number(futures.get("low"), 0) if futures.get("low") else "N/A"
    ctx["futures_volume"] = format_number(futures.get("volume"), 0) if futures.get("volume") else "N/A"

    # 期現貨價差
    taiex_val = data.get("taiex", {}).get("index")
    futures_val = futures.get("close")
    if taiex_val and futures_val:
        spread = futures_val - taiex_val
        ctx["futures_spread"] = spread
        ctx["futures_spread_display"] = format_number(abs(spread), 0)
    else:
        ctx["futures_spread"] = None
        ctx["futures_spread_display"] = "N/A"

    # 三大法人
    inst = data.get("institutional", {})

    # 外資 (含買進/賣出明細)
    ctx["foreign_net"] = inst.get("foreign_net")
    fn_display, fn_unit = format_amount(inst.get("foreign_net"))
    ctx["foreign_net_display"] = fn_display
    ctx["foreign_net_yi"] = fn_unit

    # 外資買進/賣出金額
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

    # 投信 (含買進/賣出明細)
    ctx["trust_net"] = inst.get("trust_net")
    tn_display, tn_unit = format_amount(inst.get("trust_net"))
    ctx["trust_net_display"] = tn_display
    ctx["trust_net_yi"] = tn_unit

    # 投信買進/賣出金額
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

    # 融資
    margin = data.get("margin", {})
    margin_amt = margin.get("margin_balance_amount")
    if margin_amt:
        m_display, m_unit = format_amount(margin_amt)
        ctx["margin_display"] = m_display
        ctx["margin_sub"] = m_unit
    elif margin.get("margin_balance"):
        ctx["margin_display"] = format_number(margin["margin_balance"])
        ctx["margin_sub"] = "張"
    else:
        ctx["margin_display"] = "N/A"
        ctx["margin_sub"] = ""

    # VIX (含7天圖表)
    vix = data.get("vix", {})
    vix_val = vix.get("value")
    ctx["vix_value"] = format_number(vix_val, 2) if vix_val else "N/A"
    vix_label, vix_color, vix_bg = get_vix_level(vix_val)
    ctx["vix_label"] = vix_label
    ctx["vix_color_class"] = vix_color
    # 用 inline style 避免模板 class 問題
    ctx["vix_label_style"] = vix_bg if "background" in str(vix_bg) else (
        "background: rgba(34,197,94,0.15); color: #22c55e;" if vix_val and vix_val < 15 else
        "background: rgba(59,130,246,0.15); color: #3b82f6;" if vix_val and vix_val < 20 else
        "background: rgba(234,179,8,0.15); color: #eab308;" if vix_val and vix_val < 25 else
        "background: rgba(239,68,68,0.15); color: #ef4444;" if vix_val else ""
    )

    # VIX 變動
    vix_prev = vix.get("prev_value")
    if vix_val and vix_prev:
        vix_change = round(vix_val - vix_prev, 2)
        ctx["vix_change"] = vix_change
        ctx["vix_change_display"] = format_number(abs(vix_change), 2)
    else:
        ctx["vix_change"] = None
        ctx["vix_change_display"] = None

    # VIX 7天圖表資料
    vix_chart = vix.get("chart", [])
    ctx["vix_chart_data"] = json.dumps({
        "labels": [p["date"][-5:] for p in vix_chart],
        "values": [p["close"] for p in vix_chart],
    })

    # 漲跌家數
    breadth = data.get("breadth", {})
    ctx["tse_up"] = breadth.get("tse_up")
    ctx["tse_down"] = breadth.get("tse_down")
    ctx["tse_flat"] = breadth.get("tse_flat")
    ctx["otc_up"] = breadth.get("otc_up")
    ctx["otc_down"] = breadth.get("otc_down")
    ctx["otc_flat"] = breadth.get("otc_flat")

    # 外資排行
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

    # 圖表資料
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

    # ====== 新增指標 ======

    # 融資維持率
    mr = data.get("margin_ratio", {})
    mr_val = mr.get("ratio")
    ctx["margin_ratio_value"] = format_number(mr_val, 1) if mr_val else "N/A"
    ctx["margin_ratio_raw"] = mr_val
    if mr_val:
        if mr_val >= 170:
            ctx["margin_ratio_color"] = "text-green"
            ctx["margin_ratio_label"] = "安全"
            ctx["margin_ratio_style"] = "background: rgba(34,197,94,0.15); color: #22c55e;"
        elif mr_val >= 150:
            ctx["margin_ratio_color"] = "text-blue"
            ctx["margin_ratio_label"] = "正常"
            ctx["margin_ratio_style"] = "background: rgba(59,130,246,0.15); color: #3b82f6;"
        elif mr_val >= 130:
            ctx["margin_ratio_color"] = "text-yellow"
            ctx["margin_ratio_label"] = "警戒"
            ctx["margin_ratio_style"] = "background: rgba(234,179,8,0.15); color: #eab308;"
        else:
            ctx["margin_ratio_color"] = "text-red"
            ctx["margin_ratio_label"] = "危險"
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

    # 比特幣恐慌與貪婪指數
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
            ctx["pcr_label"] = "偏多"
            ctx["pcr_style"] = "background: rgba(34,197,94,0.15); color: #22c55e;"
        elif pcr_val == 100:
            ctx["pcr_color"] = "text-blue"
            ctx["pcr_label"] = "中性"
            ctx["pcr_style"] = "background: rgba(59,130,246,0.15); color: #3b82f6;"
        else:
            ctx["pcr_color"] = "text-red"
            ctx["pcr_label"] = "偏空"
            ctx["pcr_style"] = "background: rgba(239,68,68,0.15); color: #ef4444;"
    else:
        ctx["pcr_color"] = ""
        ctx["pcr_label"] = ""
        ctx["pcr_style"] = ""

    put_oi = pcr.get("put_oi")
    call_oi = pcr.get("call_oi")
    ctx["pcr_put_oi"] = format_number(put_oi) if put_oi else "N/A"
    ctx["pcr_call_oi"] = format_number(call_oi) if call_oi else "N/A"

    # 美國 10 年期公債殖利率
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

    # ====== 自營商買賣超 ======
    dealer_net = inst.get("dealer_net")
    ctx["dealer_net"] = dealer_net
    dn_display, dn_unit = format_amount(dealer_net)
    ctx["dealer_net_display"] = dn_display
    ctx["dealer_net_yi"] = dn_unit

    # ====== 三大法人台指期未平倉 ======
    foi = data.get("futures_oi", {})
    for key in ["foreign", "trust", "dealer", "total"]:
        item = foi.get(key, {})
        chg = item.get("change")
        oi = item.get("oi")
        ctx[f"foi_{key}_change"] = format_number(chg) if chg is not None else "N/A"
        ctx[f"foi_{key}_change_raw"] = chg
        ctx[f"foi_{key}_oi"] = format_number(oi) if oi is not None else "N/A"
        ctx[f"foi_{key}_oi_raw"] = oi

    # ====== 期權觀測指標 ======
    senti = data.get("sentiment", {})

    # 微台多空
    ctx["micro_sentiment"] = senti.get("micro_sentiment")
    ctx["micro_sentiment_display"] = f"{senti['micro_sentiment']:.2f}%" if senti.get("micro_sentiment") is not None else "N/A"
    ctx["micro_sentiment_prev"] = senti.get("micro_sentiment_prev")
    ctx["micro_sentiment_prev_display"] = f"{senti['micro_sentiment_prev']:.2f}%" if senti.get("micro_sentiment_prev") is not None else "N/A"

    # 小台多空
    ctx["mini_sentiment"] = senti.get("mini_sentiment")
    ctx["mini_sentiment_display"] = f"{senti['mini_sentiment']:.2f}%" if senti.get("mini_sentiment") is not None else "N/A"
    ctx["mini_sentiment_prev"] = senti.get("mini_sentiment_prev")
    ctx["mini_sentiment_prev_display"] = f"{senti['mini_sentiment_prev']:.2f}%" if senti.get("mini_sentiment_prev") is not None else "N/A"

    # PCR 含前日
    ctx["pcr_prev_value"] = format_number(senti.get("pcr_prev"), 1) if senti.get("pcr_prev") else "N/A"

    return ctx


def generate_dashboard(data, output_dir=None):
    """生成 HTML 儀表板"""
    # 讀取模板
    template_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "dashboard_template.html")
    with open(template_path, "r", encoding="utf-8") as f:
        template_str = f.read()

    template = Template(template_str)

    # 準備模板資料
    ctx = prepare_template_data(data)

    # 渲染
    html = template.render(**ctx)

    # 輸出
    if output_dir is None:
        output_dir = os.path.dirname(os.path.abspath(__file__))

    os.makedirs(output_dir, exist_ok=True)

    # 輸出 index.html (固定檔名，適合靜態網站)
    index_path = os.path.join(output_dir, "index.html")
    with open(index_path, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"📄 儀表板已生成: {index_path}")

    # 同時輸出帶日期的版本 (歷史紀錄)
    date_str = data["date"]
    archive_dir = os.path.join(output_dir, "archive")
    os.makedirs(archive_dir, exist_ok=True)
    archive_path = os.path.join(archive_dir, f"dashboard_{date_str}.html")
    with open(archive_path, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"📁 歷史檔案: {archive_path}")

    # 輸出 JSON 資料 (方便其他程式使用)
    json_path = os.path.join(output_dir, "latest_data.json")
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2, default=str)
    print(f"📊 JSON 資料: {json_path}")

    return index_path


def main():
    parser = argparse.ArgumentParser(description="台灣股市每日戰略儀表板生成器")
    parser.add_argument("date", nargs="?", default=None, help="日期 YYYYMMDD (預設今天)")
    parser.add_argument("--output", "-o", default=None, help="輸出目錄")
    args = parser.parse_args()

    date_str = args.date or datetime.now().strftime("%Y%m%d")

    # 抓取資料
    data = fetch_all_data(date_str)

    # 生成儀表板
    output_path = generate_dashboard(data, args.output)

    print(f"\n🎉 完成！請用瀏覽器打開 {output_path}")


if __name__ == "__main__":
    main()
