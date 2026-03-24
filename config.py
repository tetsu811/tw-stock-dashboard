"""
台灣股市戰略儀表板 - 設定檔
請在此填入你的 WordPress 連線資訊
"""

# ============================================================
# WordPress 設定 (請填入你的資訊)
# ============================================================
WORDPRESS_URL = "https://your-site.com"          # 你的 WordPress 網址
WORDPRESS_USER = "your-username"                  # WordPress 使用者名稱
WORDPRESS_APP_PASSWORD = "xxxx xxxx xxxx xxxx"    # Application Password (在 WordPress 後台 > 使用者 > 安全 > 應用程式密碼 生成)

# 發布設定
POST_STATUS = "publish"        # "publish" 公開發布, "draft" 存為草稿
POST_CATEGORY_IDS = []         # 文章分類 ID，例如 [5, 12]
POST_TAG_IDS = []              # 標籤 ID，例如 [3, 7]
UPDATE_FIXED_PAGE = True       # 是否同時更新固定頁面
FIXED_PAGE_ID = None           # 固定頁面的 ID (到 WordPress 後台查看)

# ============================================================
# 排程設定
# ============================================================
CRON_HOUR = 17                 # 每天幾點執行 (24小時制，台灣時間下午5點)
CRON_MINUTE = 0                # 幾分

# ============================================================
# FinMind API 設定 (補充資料來源)
# ============================================================
import os
FINMIND_TOKEN = os.environ.get("FINMIND_TOKEN", "")  # FinMind API token (from env var)
