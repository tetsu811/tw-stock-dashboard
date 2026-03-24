# 台灣股市每日戰略儀表板

自動化抓取台灣股市、期貨、法人籌碼、國際指標等數據，生成深色主題 HTML 儀表板。

## 儀表板內容

- **加權指數 & 台指期貨** — 收盤價、漲跌幅、成交量、期現貨價差
- **三大法人現貨買賣超** — 外資（含買進/賣出明細）、投信、自營商、融資餘額
- **三大法人台指期未平倉** — 外資/投信/自營商 未平倉口數與日變動
- **期權觀測指標** — 微台多空比、小台多空比、Put/Call Ratio（含前日比較）
- **VIX 恐慌指數** — 近 7 天走勢圖 + 恐慌等級標示
- **市場情緒指標** — CNN Fear & Greed、比特幣恐慌貪婪指數、融資維持率
- **漲跌家數** — 上市 / 上櫃漲跌平統計比例
- **外資排行** — 買超 / 賣超前 10 名個股
- **國際指標趨勢** — 美元指數（30 天）、日圓匯率（30 天）、美國 10 年期公債殖利率（7 天）

## 資料來源

| 資料 | 來源 |
|------|------|
| 加權指數、成交量 | 台灣證券交易所 (TWSE) 開放資料 API |
| 三大法人買賣超 | TWSE 三大法人買賣超日報 |
| 外資買賣超排行 | TWSE 外資及陸資買賣超彙總 |
| 台指期貨 | 台灣期貨交易所 (TAIFEX) |
| 期貨未平倉 | TAIFEX 三大法人未平倉 CSV |
| 微台/小台多空比 | TAIFEX 散戶多空比計算 |
| 融資融券 | TWSE 融資融券統計 |
| 融資維持率 | TWSE 融資維持率報表 |
| 漲跌家數 | TWSE / TPEx 每日收盤統計 |
| Put/Call Ratio | TAIFEX 選擇權未平倉 |
| 美元指數、日圓、VIX | Yahoo Finance API |
| US 10Y 殖利率 | Yahoo Finance API |
| CNN Fear & Greed | CNN DataViz API |
| 比特幣恐慌貪婪指數 | Alternative.me API |

## 自動更新排程

**每個交易日（週一至週五）下午 5:00（台灣時間）** 自動抓取最新資料並更新儀表板。

選擇 5:00 PM 是因為台灣股市下午 1:30 收盤，收盤後各項資料（法人買賣超、融資融券等）通常在 3:00–4:30 PM 之間陸續公布，5:00 PM 執行可確保所有資料皆已到位。

### 方式 A：GitHub Actions（推薦）

專案已內建 `.github/workflows/daily_update.yml`，推送到 GitHub 後會自動在 UTC 09:00（= 台灣 17:00）執行。

設定步驟：
1. 在 GitHub 建立一個 repository 並把專案推上去
2. GitHub Actions 會自動偵測 workflow 檔案，無需額外設定
3. 也可以到 Actions 頁面手動觸發 `workflow_dispatch`

### 方式 B：Linux Crontab

在你的主機上設定 crontab：

```bash
crontab -e
```

加入這行：

```
0 17 * * 1-5 /path/to/tw-stock-dashboard/run_daily.sh >> /path/to/tw-stock-dashboard/logs/cron.log 2>&1
```

### 方式 C：手動執行

```bash
python generate.py              # 抓取今天的資料
python generate.py 20260324     # 指定日期
python generate.py --output /other/path  # 指定輸出目錄
```

## 部署方式

### 靜態網站（GitHub Pages / Nginx）

`generate.py` 執行後會產生 `index.html`，直接用任何靜態伺服器託管即可。

### WordPress

1. 到 WordPress 後台 → 使用者 → 安全 → 建立「應用程式密碼」
2. 編輯 `config.py`，填入：
   ```python
   WORDPRESS_URL = "https://your-site.com"
   WORDPRESS_USER = "your-username"
   WORDPRESS_APP_PASSWORD = "xxxx xxxx xxxx xxxx"
   FIXED_PAGE_ID = 123  # 你要更新的頁面 ID
   ```
3. 測試連線：
   ```bash
   python wp_publisher.py --test
   ```
4. 發布：
   ```bash
   python wp_publisher.py
   ```
5. 在 `run_daily.sh` 中取消 `python3 wp_publisher.py` 的註解，即可每日自動發布

## 安裝

```bash
pip install -r requirements.txt
```

依賴套件：requests、beautifulsoup4、jinja2

## 專案結構

```
tw-stock-dashboard/
├── generate.py              # 主程式：抓取資料 + 生成儀表板
├── data_fetcher.py          # 資料抓取模組（所有 API 呼叫）
├── dashboard_template.html  # HTML 模板（Jinja2）
├── wp_publisher.py          # WordPress 發布模組
├── config.py                # 設定檔（WordPress 連線、排程）
├── run_daily.sh             # Crontab 排程腳本
├── requirements.txt         # Python 依賴
├── index.html               # 生成的儀表板（每日更新）
├── latest_data.json         # 最新資料 JSON
├── archive/                 # 歷史儀表板
│   └── dashboard_YYYYMMDD.html
└── .github/
    └── workflows/
        └── daily_update.yml # GitHub Actions 自動排程
```
