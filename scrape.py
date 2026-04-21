#!/usr/bin/env python3
"""
海南离岛免税销售数据爬虫
从海口海关官网抓取月度免税销售数据，解析xlsx/xls文件，保存到data.json

用 Playwright 渲染 JS 反爬页面，绕过 WAF 挑战。

用法:
    python3 scrape.py              # 抓取所有缺失月份
    python3 scrape.py --force      # 强制重新抓取所有月份
    python3 scrape.py --month 2026-02  # 抓取指定月份
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

try:
    from playwright.sync_api import sync_playwright
except ImportError:
    print("需要 playwright: pip install playwright && playwright install chromium")
    exit(1)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, "data")
DATA_FILE = os.path.join(BASE_DIR, "data.json")
LISTING_URL = "http://haikou.customs.gov.cn/haikou_customs/605737/fdzdgknr82/605745/58527f05-{page}.html"
BASE_URL = "http://haikou.customs.gov.cn"

# Known article paths (month -> article page path)
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

# ─── Playwright 浏览器管理 ─────────────────────────────────────

_pw = None
_browser = None


def get_browser():
    """获取全局 Playwright 浏览器实例（懒初始化）"""
    global _pw, _browser
    if _browser is None:
        _pw = sync_playwright().start()
        _browser = _pw.chromium.launch(headless=True)
    return _browser


def close_browser():
    """关闭浏览器"""
    global _pw, _browser
    if _browser:
        _browser.close()
        _browser = None
    if _pw:
        _pw.stop()
        _pw = None


def pw_fetch_html(url, wait_ms=8000):
    """用 Playwright 打开页面，等待 JS 渲染完成，返回 HTML"""
    browser = get_browser()
    context = browser.new_context(
        user_agent=(
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/124.0.0.0 Safari/537.36"
        )
    )
    page = context.new_page()
    try:
        page.goto(url, wait_until="networkidle", timeout=30000)
        # 额外等待，确保反爬 JS 执行完毕
        page.wait_for_timeout(wait_ms)
        html = page.content()
        return html
    except Exception as e:
        print(f"  Playwright 抓取失败 {url}: {e}")
        return ""
    finally:
        context.close()


def pw_download(url, output_path):
    """用 curl 下载文件（Playwright 页面中已拿到真实 URL 后）"""
    cmd = ["curl", "-s", "-L", "-o", output_path, "--max-time", "60", url]
    result = subprocess.run(cmd, capture_output=True, timeout=65)
    return result.returncode == 0


# ─── 发现 & 获取 ────────────────────────────────────────────────


def discover_articles_from_listing():
    """用 Playwright 渲染列表页，发现新的文章链接"""
    articles = {}
    for page_num in range(1, 10):
        url = LISTING_URL.format(page=page_num)
        print(f"  渲染列表页 {page_num}...")
        html = pw_fetch_html(url, wait_ms=5000)
        if not html:
            break

        pattern = r'<a[^>]*href="([^"]+)"[^>]*>([^<]*离岛免税[^<]*)</a>'
        matches = re.findall(pattern, html)

        if not matches:
            break

        found_new = False
        for href, title in matches:
            m = re.search(r"(\d{4})年(\d{1,2})月", title)
            if m:
                year = m.group(1)
                month = m.group(2).zfill(2)
                key = f"{year}-{month}"
                if key not in articles:
                    articles[key] = href
                    found_new = True

        if not found_new:
            break

        time.sleep(1)

    return articles


def get_xlsx_url(article_path):
    """用 Playwright 渲染文章页，获取 xlsx 下载链接"""
    if article_path.startswith("/"):
        article_path = BASE_URL + article_path

    html = pw_fetch_html(article_path, wait_ms=5000)
    if not html:
        return None

    # 查找 xlsx/xls 下载链接
    m = re.search(r'href="([^"]+\.xls[x]?)"', html)
    if m:
        url = m.group(1)
        if not url.startswith("http"):
            url = BASE_URL + url
        return url

    return None


# ─── 解析 ──────────────────────────────────────────────────────


def parse_xlsx(filepath):
    """解析 xlsx/xls 文件，提取免税数据"""
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

    # 计算人均消费: 万元 / 万人次 = 元/人
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
    """从一行中提取数值"""
    nums = []
    for c in row:
        try:
            v = float(c)
            nums.append(v)
        except (ValueError, TypeError):
            pass
    return nums


# ─── 数据存储 ──────────────────────────────────────────────────


def load_existing_data():
    """加载已有数据（兼容旧数组和新对象格式）"""
    if os.path.exists(DATA_FILE):
        with open(DATA_FILE, "r", encoding="utf-8") as f:
            raw = json.load(f)
        if isinstance(raw, list):
            return raw
        if isinstance(raw, dict) and "data" in raw:
            return raw["data"]
    return []


def save_data(data):
    """保存数据到 data.json"""
    data.sort(key=lambda x: x["month"])
    output = {
        "last_updated": datetime.now().isoformat(),
        "data": data,
    }
    with open(DATA_FILE, "w", encoding="utf-8") as f:
        json.dump(output, f, ensure_ascii=False, indent=2)
    print(f"已保存 {len(data)} 条记录到 data.json (更新时间: {output['last_updated']})")


# ─── 抓取逻辑 ──────────────────────────────────────────────────


def scrape_month(month_key, article_path, force=False):
    """抓取单个月份的数据"""
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

    print(f"  {month_key}: 正在下载 {xlsx_url}...")
    if not pw_download(xlsx_url, tmp_file):
        print(f"  {month_key}: 下载失败")
        return False

    print(f"  {month_key}: 正在解析...")
    data = parse_xlsx(tmp_file)
    if not data:
        print(f"  {month_key}: 解析失败")
        return False

    data["month"] = month_key

    # 更新已有数据
    existing = [d for d in existing if d["month"] != month_key]
    existing.append(data)
    save_data(existing)

    print(
        f"  {month_key}: 完成 (金额={data.get('amount')}万元, "
        f"人次={data.get('visitors')}万, 件数={data.get('items')}万)"
    )
    return True


def scrape_all(force=False):
    """抓取所有已知月份"""
    os.makedirs(DATA_DIR, exist_ok=True)

    articles = KNOWN_ARTICLES.copy()

    # 尝试从列表页发现新文章
    print("正在从海关官网发现新数据...")
    discovered = discover_articles_from_listing()
    new_count = sum(1 for k in discovered if k not in KNOWN_ARTICLES)
    articles.update(discovered)

    print(f"共发现 {len(articles)} 个月份 (其中 {new_count} 个新增)")
    print()

    success = 0
    fail = 0
    for month_key in sorted(articles.keys()):
        if scrape_month(month_key, articles[month_key], force=force):
            success += 1
        else:
            fail += 1
        time.sleep(0.5)

    print(f"\n完成: 成功 {success}, 失败 {fail}")


def main():
    parser = argparse.ArgumentParser(description="海南离岛免税销售数据爬虫")
    parser.add_argument("--force", action="store_true", help="强制重新抓取所有月份")
    parser.add_argument("--month", type=str, help="只抓取指定月份 (格式: 2026-02)")
    parser.add_argument(
        "--discover", action="store_true", help="只发现新文章链接，不下载"
    )
    args = parser.parse_args()

    try:
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
                # 尝试动态发现该月份
                print(f"未知月份 {args.month}，尝试从海关官网查找...")
                discovered = discover_articles_from_listing()
                if args.month in discovered:
                    scrape_month(args.month, discovered[args.month], force=True)
                else:
                    print(f"未找到 {args.month} 的数据")
                    print("已知月份:", ", ".join(sorted(KNOWN_ARTICLES.keys())))
        else:
            scrape_all(force=args.force)
    finally:
        close_browser()


if __name__ == "__main__":
    main()
