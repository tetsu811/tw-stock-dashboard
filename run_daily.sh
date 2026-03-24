#!/bin/bash
# ============================================================
# 台灣股市戰略儀表板 - 每日排程腳本
#
# 使用方式:
#   1. 手動執行: bash run_daily.sh
#   2. 加入 crontab: crontab -e
#      然後加入這行 (每天下午5:00執行):
#      0 17 * * 1-5 /path/to/tw-stock-dashboard/run_daily.sh >> /path/to/tw-stock-dashboard/logs/cron.log 2>&1
# ============================================================

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$SCRIPT_DIR"

# 建立 logs 目錄
mkdir -p logs

echo "========================================"
echo "$(date '+%Y-%m-%d %H:%M:%S') - 開始執行"
echo "========================================"

# 執行儀表板生成
python3 generate.py

# 可選：如果你要部署到伺服器，取消下面的註解
# 部署到 GitHub Pages
# git add -A && git commit -m "Update dashboard $(date '+%Y-%m-%d')" && git push

# 部署到你的伺服器 (用 rsync 或 scp)
# rsync -avz index.html your-server:/var/www/html/dashboard/
# scp index.html your-server:/var/www/html/dashboard/

# 可選：發布到 WordPress
# python3 wp_publisher.py

echo "$(date '+%Y-%m-%d %H:%M:%S') - 執行完成"
echo ""
