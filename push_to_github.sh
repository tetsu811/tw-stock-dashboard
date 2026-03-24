#!/bin/bash
# ============================================================
# 一鍵推送到 GitHub
#
# 使用方式：
#   1. 先在 GitHub 建好 repository
#   2. 修改下面的 REPO_URL 為你的 repo 網址
#   3. 執行: bash push_to_github.sh
# ============================================================

# ⬇️ 請改成你的 GitHub repo URL
REPO_URL="https://github.com/YOUR_USERNAME/tw-stock-dashboard.git"

cd "$(dirname "${BASH_SOURCE[0]}")"

git init
git add -A
git commit -m "Initial commit: 台灣股市每日戰略儀表板

- 加權指數 & 台指期貨即時數據
- 三大法人現貨買賣超 & 期貨未平倉
- 期權觀測指標 (微台/小台多空比、PCR)
- VIX 7天走勢、CNN/Crypto 恐慌貪婪指數
- 融資維持率、漲跌家數、外資排行
- 國際指標 (美元指數、日圓、US10Y)
- GitHub Actions 每日下午5點自動更新
- WordPress 發布模組 (選用)"

git branch -M main
git remote add origin "$REPO_URL"
git push -u origin main

echo ""
echo "✅ 推送完成！GitHub Actions 會在每個交易日 17:00 自動執行"
