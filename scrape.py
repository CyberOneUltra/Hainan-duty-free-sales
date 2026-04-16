#!/usr/bin/env python3
"""
海南离岛免税销售数据爬虫
从海口海关官网抓取月度免税销售数据，解析xlsx/xls文件，保存到data.json

用法:
    python3 scrape.py              # 抓取所有缺失月份
    python3 scrape.py --force      # 强制重新抓取所有月份
    python3 scrape.py --month 2026-02  # 只抓取指定月份
"""

import os
import re
import json
import time
import argparse
import subprocess
from datetime import datetime

try:
    import openpyxl
except ImportError:
    print("需要 openpyxl: pip install openpyxl")
    exit(1)

try:
    import xlrd
except ImportError:
    print("需要 xlrd: pip install xlrd")
    exit(1)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, "data")
DATA_FILE = os.path.join(BASE_DIR, "data.json")
LISTING_URL = "http://haikou.customs.gov.cn/haikou_customs/605737/fdzdgknr82/605745/58527f05-{page}.html"
BASE_URL = "http://haikou.customs.gov.cn"
UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"

# Known article paths (month -> article page path)
# These are discovered from the listing pages
KNOWN_ARTICLES = {
    "2024-01": "/haikou_customs/605737/fdzdgknr82/605745/5684450/index.html",
    "2024-02": "/haikou_customs/605737/fdzdgknr82/605745/5757645/index.html",
    "2024-03": "/haikou_customs/605737/fdzdgknr82/605745/5825176/index.html",
    "2024-04": "/haikou_customs/605737/fdzdgknr82/605745/5884831/index.html",
    "2024-05": "/haikou_customs/605737/fdzdgknr82/605745/5941485/index.html",
    "2024-06": "/haikou_customs/605737/fdzdgknr82/605745/5992300/index.html",
    "2024-07": "/haikou_customs/605737/fdzdgknr82/605745/6050921/index.html",
    "2024-08": "/haikou_customs/605737/fdzdgknr82/605745/6108463/index.html",
    "2024-09": "/haikou_customs/605737/fdzdgknr82/605745/6159070/index.html",
    "2024-10": "/haikou_customs/605737/fdzdgknr82/605745/6211777/index.html",
    "2024-11": "/haikou_customs/605737/fdzdgknr82/605745/6272443/index.html",
    "2024-12": "/haikou_customs/605737/fdzdgknr82/605745/6326412/index.html",
    "2025-01": "/haikou_customs/605737/fdzdgknr82/605745/6369960/index.html",
    "2025-02": "/haikou_customs/605737/fdzdgknr82/605745/6422507/index.html",
    "2025-03": "/haikou_customs/605737/fdzdgknr82/605745/6471759/index.html",
    "2025-04": "/haikou_customs/605737/fdzdgknr82/605745/6524076/index.html",
    "2025-05": "/haikou_customs/605737/fdzdgknr82/605745/6584794/index.html",
    "2025-06": "/haikou_customs/605737/fdzdgknr82/605745/6630041/index.html",
    "2025-07": "/haikou_customs/605737/fdzdgknr82/605745/6681393/index.html",
    "2025-08": "/haikou_customs/605737/fdzdgknr82/605745/6741898/index.html",
    "2025-09": "/haikou_customs/605737/fdzdgknr82/605745/6782549/index.html",
    "2025-10": "/haikou_customs/605737/fdzdgknr82/605745/6830442/index.html",
    "2025-11": "/haikou_customs/605737/fdzdgknr82/605745/6896365/index.html",
    "2025-12": "/haikou_customs/605737/fdzdgknr82/605745/6952249/index.html",
    "2026-01": "/haikou_customs/605737/fdzdgknr82/605745/7037278/index.html",
    "2026-02": "/haikou_customs/605737/fdzdgknr82/605745/7074681/index.html",
}


def curl_fetch(url, output_path=None):
    """Fetch URL using curl with anti-bot headers"""
    cmd = ["curl", "-s", "-L", "-A", UA, "--max-time", "30"]
    if output_path:
        cmd.extend(["-o", output_path, url])
        result = subprocess.run(cmd, capture_output=True, timeout=35)
        return result.returncode == 0
    else:
        cmd.append(url)
        result = subprocess.run(cmd, capture_output=True, timeout=35)
        return result.stdout.decode("utf-8", errors="replace")


def discover_articles_from_listing():
    """Scrape the listing pages to discover article URLs"""
    articles = {}
    for page in range(1, 10):  # Check first 10 pages
        url = LISTING_URL.format(page=page)
        html = curl_fetch(url)
        if not html:
            break

        # Find all duty-free report links
        pattern = r'<a[^>]*href="([^"]+)"[^>]*>([^<]*离岛免税[^<]*)</a>'
        matches = re.findall(pattern, html)

        if not matches:
            break

        for href, title in matches:
            # Extract month from title like "2025年6月海南离岛免税销售情况表"
            m = re.search(r"(\d{4})年(\d{1,2})月", title)
            if m:
                year = m.group(1)
                month = m.group(2).zfill(2)
                key = f"{year}-{month}"
                if key not in articles:
                    articles[key] = href

        time.sleep(0.5)  # Be polite

    return articles


def get_xlsx_url(article_path):
    """Get the xlsx/xls download URL from an article page"""
    if article_path.startswith("/"):
        article_path = BASE_URL + article_path

    html = curl_fetch(article_path)
    if not html:
        return None

    # Look for xlsx or xls download links
    m = re.search(r'href="([^"]+\.xls[x]?)"', html)
    if m:
        url = m.group(1)
        if not url.startswith("http"):
            url = BASE_URL + url
        return url

    return None


def parse_xlsx(filepath):
    """Parse xlsx/xls file and extract duty-free data"""
    rows = []

    if filepath.endswith(".xlsx"):
        wb = openpyxl.load_workbook(filepath)
        ws = wb.worksheets[0]
        for row in ws.iter_rows(min_row=1, max_row=ws.max_row, values_only=True):
            rows.append([str(c) if c is not None else "" for c in row])
    else:
        wb = xlrd.open_workbook(filepath)
        ws = wb.sheet_by_index(0)
        for r in range(ws.nrows):
            rows.append(
                [
                    str(ws.cell_value(r, c)) if ws.cell_value(r, c) != "" else ""
                    for c in range(ws.ncols)
                ]
            )

    month_data = {}

    for row in rows:
        row_str = " ".join(row)

        if (
            "免税购物金额" in row_str
            and "实际人次" not in row_str
            and "件数" not in row_str
        ):
            unit = row[2].strip() if len(row) > 2 else ""
            multiplier = 10000 if "亿" in unit else 1
            nums = _extract_nums(row)
            if len(nums) >= 1:
                month_data["amount"] = round(nums[0] * multiplier, 2)
            if len(nums) >= 2:
                month_data["amount_yoy"] = round(nums[1], 2)
            if len(nums) >= 3:
                month_data["amount_cumulative"] = round(nums[2] * multiplier, 2)
            if len(nums) >= 4:
                month_data["amount_cumulative_yoy"] = round(nums[3], 2)

        elif "免税购物实际人次" in row_str:
            nums = _extract_nums(row)
            if len(nums) >= 1:
                month_data["visitors"] = round(nums[0], 4)
            if len(nums) >= 2:
                month_data["visitors_yoy"] = round(nums[1], 2)
            if len(nums) >= 3:
                month_data["visitors_cumulative"] = round(nums[2], 4)
            if len(nums) >= 4:
                month_data["visitors_cumulative_yoy"] = round(nums[3], 2)

        elif "免税购物件数" in row_str:
            nums = _extract_nums(row)
            if len(nums) >= 1:
                month_data["items"] = round(nums[0], 4)
            if len(nums) >= 2:
                month_data["items_yoy"] = round(nums[1], 2)
            if len(nums) >= 3:
                month_data["items_cumulative"] = round(nums[2], 4)
            if len(nums) >= 4:
                month_data["items_cumulative_yoy"] = round(nums[3], 2)

    # Calculate per-capita: 万元 / 万人次 = 元/人
    if (
        month_data.get("amount")
        and month_data.get("visitors")
        and month_data["visitors"] > 0
    ):
        month_data["per_capita"] = round(
            month_data["amount"] / month_data["visitors"], 2
        )
    if (
        month_data.get("amount_cumulative")
        and month_data.get("visitors_cumulative")
        and month_data.get("visitors_cumulative", 0) > 0
    ):
        month_data["per_capita_cumulative"] = round(
            month_data["amount_cumulative"] / month_data["visitors_cumulative"], 2
        )

    return month_data


def _extract_nums(row):
    """Extract numeric values from a row"""
    nums = []
    for c in row:
        try:
            v = float(c)
            nums.append(v)
        except (ValueError, TypeError):
            pass
    return nums


def load_existing_data():
    """Load existing data from data.json"""
    if os.path.exists(DATA_FILE):
        with open(DATA_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return []


def save_data(data):
    """Save data to data.json"""
    data.sort(key=lambda x: x["month"])
    with open(DATA_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    print(f"已保存 {len(data)} 条记录到 data.json")


def scrape_month(month_key, article_path, force=False):
    """Scrape a single month's data"""
    existing = load_existing_data()
    existing_months = {d["month"] for d in existing}

    if not force and month_key in existing_months:
        print(f"  {month_key}: 已存在，跳过 (使用 --force 强制更新)")
        return True

    print(f"  {month_key}: 正在获取下载链接...")
    xlsx_url = get_xlsx_url(article_path)
    if not xlsx_url:
        print(f"  {month_key}: 未找到xlsx下载链接")
        return False

    ext = ".xlsx" if xlsx_url.endswith(".xlsx") else ".xls"
    tmp_file = os.path.join(DATA_DIR, f"{month_key}{ext}")

    print(f"  {month_key}: 正在下载...")
    if not curl_fetch(xlsx_url, tmp_file):
        print(f"  {month_key}: 下载失败")
        return False

    print(f"  {month_key}: 正在解析...")
    data = parse_xlsx(tmp_file)
    if not data:
        print(f"  {month_key}: 解析失败")
        return False

    data["month"] = month_key

    # Update existing data
    existing = [d for d in existing if d["month"] != month_key]
    existing.append(data)
    save_data(existing)

    print(
        f"  {month_key}: 完成 (金额={data.get('amount')}万元, "
        f"人次={data.get('visitors')}万, 件数={data.get('items')}万)"
    )
    return True


def scrape_all(force=False):
    """Scrape all known months"""
    os.makedirs(DATA_DIR, exist_ok=True)

    articles = KNOWN_ARTICLES.copy()

    # Also try to discover new articles from listing pages
    print("正在从海关官网发现新数据...")
    discovered = discover_articles_from_listing()
    articles.update(discovered)

    print(f"发现 {len(articles)} 个月份的数据")
    print()

    success = 0
    fail = 0
    for month_key in sorted(articles.keys()):
        if scrape_month(month_key, articles[month_key], force=force):
            success += 1
        else:
            fail += 1
        time.sleep(0.3)  # Be polite

    print(f"\n完成: 成功 {success}, 失败 {fail}")


def main():
    parser = argparse.ArgumentParser(description="海南离岛免税销售数据爬虫")
    parser.add_argument("--force", action="store_true", help="强制重新抓取所有月份")
    parser.add_argument("--month", type=str, help="只抓取指定月份 (格式: 2026-02)")
    parser.add_argument("--discover", action="store_true", help="只发现新文章链接，不下载")
    args = parser.parse_args()

    if args.discover:
        articles = discover_articles_from_listing()
        for k, v in sorted(articles.items()):
            print(f"  {k}: {v}")
        return

    if args.month:
        os.makedirs(DATA_DIR, exist_ok=True)
        if args.month in KNOWN_ARTICLES:
            scrape_month(args.month, KNOWN_ARTICLES[args.month], force=True)
        else:
            print(f"未知月份: {args.month}")
            print("已知月份:", ", ".join(sorted(KNOWN_ARTICLES.keys())))
    else:
        scrape_all(force=args.force)


if __name__ == "__main__":
    main()
