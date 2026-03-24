"""
台灣股市戰略儀表板 - WordPress 發布模組 (備用)
當你準備好搬家到 WordPress 時，使用這個模組

使用方式:
    1. 編輯 config.py 填入 WordPress 連線資訊
    2. python wp_publisher.py              # 發布今天的儀表板
    3. python wp_publisher.py 20260320     # 發布指定日期
"""

import os
import sys
import json
import base64
import argparse
import requests
from datetime import datetime

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

try:
    from config import (
        WORDPRESS_URL, WORDPRESS_USER, WORDPRESS_APP_PASSWORD,
        POST_STATUS, POST_CATEGORY_IDS, POST_TAG_IDS,
        UPDATE_FIXED_PAGE, FIXED_PAGE_ID
    )
except ImportError:
    print("❌ 請先建立 config.py 並填入 WordPress 連線資訊")
    sys.exit(1)


class WordPressPublisher:
    def __init__(self):
        self.base_url = WORDPRESS_URL.rstrip("/")
        self.api_url = f"{self.base_url}/wp-json/wp/v2"
        self.auth = base64.b64encode(
            f"{WORDPRESS_USER}:{WORDPRESS_APP_PASSWORD}".encode()
        ).decode()
        self.headers = {
            "Authorization": f"Basic {self.auth}",
            "Content-Type": "application/json",
        }

    def test_connection(self):
        """測試 WordPress 連線"""
        try:
            resp = requests.get(
                f"{self.api_url}/users/me",
                headers=self.headers,
                timeout=10
            )
            if resp.status_code == 200:
                user = resp.json()
                print(f"✅ WordPress 連線成功！使用者: {user.get('name', 'N/A')}")
                return True
            else:
                print(f"❌ WordPress 連線失敗: {resp.status_code} {resp.text[:200]}")
                return False
        except Exception as e:
            print(f"❌ WordPress 連線錯誤: {e}")
            return False

    def publish_post(self, title, content, date_str=None):
        """發布新文章"""
        payload = {
            "title": title,
            "content": content,
            "status": POST_STATUS,
        }
        if POST_CATEGORY_IDS:
            payload["categories"] = POST_CATEGORY_IDS
        if POST_TAG_IDS:
            payload["tags"] = POST_TAG_IDS

        try:
            resp = requests.post(
                f"{self.api_url}/posts",
                headers=self.headers,
                json=payload,
                timeout=30
            )
            if resp.status_code in (200, 201):
                post = resp.json()
                print(f"✅ 文章已發布: {post.get('link', 'N/A')}")
                return post
            else:
                print(f"❌ 發布失敗: {resp.status_code} {resp.text[:300]}")
                return None
        except Exception as e:
            print(f"❌ 發布錯誤: {e}")
            return None

    def update_page(self, page_id, title, content):
        """更新固定頁面"""
        if not page_id:
            print("⚠️ 未設定固定頁面 ID，跳過更新")
            return None

        payload = {
            "title": title,
            "content": content,
        }

        try:
            resp = requests.post(
                f"{self.api_url}/pages/{page_id}",
                headers=self.headers,
                json=payload,
                timeout=30
            )
            if resp.status_code == 200:
                page = resp.json()
                print(f"✅ 頁面已更新: {page.get('link', 'N/A')}")
                return page
            else:
                print(f"❌ 更新頁面失敗: {resp.status_code} {resp.text[:300]}")
                return None
        except Exception as e:
            print(f"❌ 更新頁面錯誤: {e}")
            return None


def publish_dashboard(date_str=None):
    """主流程：生成儀表板並發布到 WordPress"""
    from data_fetcher import fetch_all_data
    from generate import generate_dashboard

    if date_str is None:
        date_str = datetime.now().strftime("%Y%m%d")

    # 1. 抓取資料 & 生成 HTML
    data = fetch_all_data(date_str)
    output_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "output")
    html_path = generate_dashboard(data, output_dir)

    # 2. 讀取生成的 HTML
    with open(html_path, "r", encoding="utf-8") as f:
        html_content = f.read()

    # 3. 發布到 WordPress
    wp = WordPressPublisher()
    if not wp.test_connection():
        return

    dt = datetime.strptime(date_str, "%Y%m%d")
    title = f"台灣股市戰略儀表板 {dt.year}/{dt.month:02d}/{dt.day:02d}"

    # 用 HTML block 包裝
    wp_content = f'<!-- wp:html -->\n{html_content}\n<!-- /wp:html -->'

    # 發布新文章
    wp.publish_post(title, wp_content, date_str)

    # 更新固定頁面
    if UPDATE_FIXED_PAGE:
        wp.update_page(FIXED_PAGE_ID, title, wp_content)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="發布儀表板到 WordPress")
    parser.add_argument("date", nargs="?", default=None, help="日期 YYYYMMDD")
    parser.add_argument("--test", action="store_true", help="僅測試連線")
    args = parser.parse_args()

    if args.test:
        wp = WordPressPublisher()
        wp.test_connection()
    else:
        publish_dashboard(args.date)
