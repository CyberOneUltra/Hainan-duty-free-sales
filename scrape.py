#!/usr/bin/env python3
"""
海南离岛免税销售数据爬虫
从海口海关官网抓取月度免税销售数据，解析xlsx/xls文件，保存到data.json

使用 nodriver (undetected Chrome) 绕过 WAF 反爬和 TLS 指纹检测。

用法:
    python3 scrape.py              # 抓取所有缺失月份
    python3 scrape.py --force      # 强制重新抓取所有月份
    python3 scrape.py --month 2026-02  # 抓取指定月份
"""

import os
import re
import json
import time
import asyncio
import argparse
import subprocess
from datetime import datetime
from pathlib import Path

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
    import nodriver as uc
except ImportError:
    print("需要 nodriver: pip install nodriver")
    exit(1)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, "data")
DATA_FILE = os.path.join(BASE_DIR, "data.json")
LISTING_URL = "https://haikou.customs.gov.cn/haikou_customs/605737/fdzdgknr82/605745/58527f05-{page}.html"
BASE_URL = "https://haikou.customs.gov.cn"
# HTTP 备用（某些 WAF 对 https 限制更严）
BASE_URL_HTTP = "http://haikou.customs.gov.cn"

# 已知月份数据 (month -> {"article": path, "xlsx": url or None})
# xlsx: 直接下载链接（GitHub Actions IP 被 WAF 拦截时的 fallback，xlsx 文件不受限）
# article: 文章页路径（本机浏览器可用时获取 xlsx 链接）
KNOWN_ARTICLES = {
    "2024-01": {"article": "/haikou_customs/605737/fdzdgknr82/605745/5684450/index.html"},
    "2024-02": {"article": "/haikou_customs/605737/fdzdgknr82/605745/5757645/index.html"},
    "2024-03": {"article": "/haikou_customs/605737/fdzdgknr82/605745/5825176/index.html"},
    "2024-04": {"article": "/haikou_customs/605737/fdzdgknr82/605745/5884831/index.html"},
    "2024-05": {"article": "/haikou_customs/605737/fdzdgknr82/605745/5941485/index.html"},
    "2024-06": {"article": "/haikou_customs/605737/fdzdgknr82/605745/5992300/index.html"},
    "2024-07": {"article": "/haikou_customs/605737/fdzdgknr82/605745/6050921/index.html"},
    "2024-08": {"article": "/haikou_customs/605737/fdzdgknr82/605745/6108463/index.html"},
    "2024-09": {"article": "/haikou_customs/605737/fdzdgknr82/605745/6159070/index.html"},
    "2024-10": {"article": "/haikou_customs/605737/fdzdgknr82/605745/6211777/index.html"},
    "2024-11": {"article": "/haikou_customs/605737/fdzdgknr82/605745/6272443/index.html"},
    "2024-12": {"article": "/haikou_customs/605737/fdzdgknr82/605745/6326412/index.html"},
    "2025-01": {"article": "/haikou_customs/605737/fdzdgknr82/605745/6369960/index.html"},
    "2025-02": {"article": "/haikou_customs/605737/fdzdgknr82/605745/6422507/index.html"},
    "2025-03": {"article": "/haikou_customs/605737/fdzdgknr82/605745/6471759/index.html"},
    "2025-04": {"article": "/haikou_customs/605737/fdzdgknr82/605745/6524076/index.html"},
    "2025-05": {"article": "/haikou_customs/605737/fdzdgknr82/605745/6584794/index.html"},
    "2025-06": {"article": "/haikou_customs/605737/fdzdgknr82/605745/6630041/index.html"},
    "2025-07": {"article": "/haikou_customs/605737/fdzdgknr82/605745/6681393/index.html"},
    "2025-08": {"article": "/haikou_customs/605737/fdzdgknr82/605745/6741898/index.html"},
    "2025-09": {"article": "/haikou_customs/605737/fdzdgknr82/605745/6782549/index.html"},
    "2025-10": {"article": "/haikou_customs/605737/fdzdgknr82/605745/6830442/index.html"},
    "2025-11": {"article": "/haikou_customs/605737/fdzdgknr82/605745/6896365/index.html"},
    "2025-12": {"article": "/haikou_customs/605737/fdzdgknr82/605745/6952249/index.html"},
    "2026-01": {"article": "/haikou_customs/605737/fdzdgknr82/605745/7037278/index.html"},
    "2026-02": {"article": "/haikou_customs/605737/fdzdgknr82/605745/7074681/index.html"},
    "2026-03": {
        "article": "/haikou_customs/605737/fdzdgknr82/605745/7117655/index.html",
        "xlsx": "http://haikou.customs.gov.cn/haikou_customs/605737/fdzdgknr82/605745/7117655/2026042016044569957.xlsx",
    },
}

# 兼容旧格式（纯字符串值 -> dict）
def _normalize_articles():
    """兼容旧版 KNOWN_ARTICLES（字符串值）和新版（dict 值）"""
    for k, v in list(KNOWN_ARTICLES.items()):
        if isinstance(v, str):
            KNOWN_ARTICLES[k] = {"article": v}

_normalize_articles()

# ─── nodriver 浏览器管理 ───────────────────────────────────────

_browser = None


async def get_browser():
    """获取全局 nodriver 浏览器实例"""
    global _browser
    if _browser is None:
        # GitHub Actions: browser-actions/setup-chrome installs here
        chrome_path = os.environ.get("CHROME_PATH") or _find_chrome()
        kwargs = dict(
            headless=True,
            sandbox=False,
            browser_args=[
                "--no-sandbox",
                "--disable-dev-shm-usage",
                "--ignore-certificate-errors",
                "--disable-gpu",
            ],
        )
        if chrome_path:
            kwargs["browser_executable_path"] = chrome_path
        print(f"  启动浏览器: {chrome_path or '自动检测'}")
        _browser = await uc.start(**kwargs)
    return _browser


def _find_chrome():
    """尝试查找 Chrome 可执行文件"""
    import shutil
    # 优先在 PATH 中找
    for name in ["google-chrome-stable", "google-chrome", "chromium", "chrome"]:
        found = shutil.which(name)
        if found:
            return found
    # GitHub Actions setup-chrome 默认路径
    gh_paths = [
        "/opt/hostedtoolcache/setup-chrome/chromium/stable/x64/chrome",
        "/usr/bin/google-chrome-stable",
        "/usr/bin/google-chrome",
    ]
    for p in gh_paths:
        if os.path.exists(p):
            return p
    return None


async def close_browser():
    """关闭浏览器"""
    global _browser
    if _browser:
        try:
            _browser.stop()
        except (RuntimeError, Exception) as e:
            print(f"  关闭浏览器时忽略: {e}")
        finally:
            _browser = None


async def nw_fetch_html(url, wait_sec=10):
    """用 nodriver 打开页面，等待渲染完成，返回 HTML"""
    browser = await get_browser()
    try:
        page = await browser.get(url)
        await asyncio.sleep(wait_sec)

        # 检查是否遇到错误页面
        html = await page.get_content()
        if "504" in html and "连接超时" in html:
            print(f"  ⚠️  服务器返回 504，源站不可达")
            return ""
        if "502" in html and "Bad Gateway" in html:
            print(f"  ⚠️  服务器返回 502")
            return ""

        return html
    except Exception as e:
        print(f"  nodriver 抓取失败 {url}: {e}")
        return ""


def curl_download(url, output_path):
    """用 curl 下载文件"""
    cmd = ["curl", "-s", "-k", "-L", "-o", output_path, "--max-time", "60", url]
    result = subprocess.run(cmd, capture_output=True, timeout=65)
    return result.returncode == 0


def curl_fetch_html(url, referer=None):
    """用 curl + 浏览器 UA 抓取 HTML（绕过基础 WAF，nodriver 的 fallback）"""
    cmd = [
        "curl", "-s", "-k", "-L", "--max-time", "60",
        "-w", "\n__CURL_HTTP_CODE__%{http_code}",
        "-H", "User-Agent: Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
              "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/135.0.0.0 Safari/537.36",
        "-H", "Accept: text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
        "-H", "Accept-Language: zh-CN,zh;q=0.9,en;q=0.8",
    ]
    if referer:
        cmd += ["-H", f"Referer: {referer}"]
    cmd.append(url)
    try:
        result = subprocess.run(cmd, capture_output=True, timeout=65)
        raw = result.stdout.decode("utf-8", errors="ignore") if result.stdout else ""

        # 提取 HTTP 状态码
        http_code = 0
        m = re.search(r"__CURL_HTTP_CODE__(\d+)", raw)
        if m:
            http_code = int(m.group(1))
            raw = raw[:m.start()]

        if result.returncode != 0:
            print(f"    curl 退出码 {result.returncode}")
            return ""
        if http_code and http_code >= 400:
            print(f"    curl HTTP {http_code}")
            return ""
        if "504" in raw and "连接超时" in raw:
            print(f"    curl 拿到 504 页面")
            return ""
        if len(raw) < 500:
            print(f"    curl 响应过短 ({len(raw)} bytes), 可能被拦截")
            return ""

        return raw
    except subprocess.TimeoutExpired:
        print(f"    curl 超时 (60s)")
        return ""
    except Exception as e:
        print(f"    curl 异常: {e}")
        return ""


def curl_download(url, output_path, referer=None):
    """用 curl 下载文件（带浏览器 UA 绕 WAF）"""
    cmd = [
        "curl", "-s", "-k", "-L",
        "-o", output_path, "--max-time", "60",
        "-H", "User-Agent: Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
              "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/135.0.0.0 Safari/537.36",
    ]
    if referer:
        cmd += ["-H", f"Referer: {referer}"]
    cmd.append(url)
    try:
        result = subprocess.run(cmd, capture_output=True, timeout=65)
        return result.returncode == 0
    except Exception:
        return False


def _extract_xlsx_from_html(html):
    """从 HTML 中提取 xlsx/xls 下载链接"""
    m = re.search(r'href="([^"]+\.xls[x]?)"', html)
    if m:
        url = m.group(1)
        if not url.startswith("http"):
            url = BASE_URL + url
        return url
    return None


# ─── 发现 & 获取 ────────────────────────────────────────────────


async def discover_articles_from_listing():
    """从列表页发现新的文章链接（curl 优先，nodriver 兜底）"""
    articles = {}
    no_match_streak = 0
    for page_num in range(1, 10):
        # HTTP 优先，HTTPS 兜底
        urls = [
            f"http://haikou.customs.gov.cn/haikou_customs/605737/fdzdgknr82/605745/58527f05-{page_num}.html",
            LISTING_URL.format(page=page_num),
        ]
        html = ""
        for url in urls:
            proto = "http" if "http://" in url and "https://" not in url else "https"
            print(f"  curl 获取列表页 {page_num} ({proto})...")
            html = curl_fetch_html(url)
            if html:
                break

        if not html:
            print(f"  curl 失败，尝试 nodriver 渲染列表页 {page_num}...")
            html = await nw_fetch_html(LISTING_URL.format(page=page_num), wait_sec=12)

        if not html:
            print(f"  列表页 {page_num}: 获取失败，停止")
            break

        title_m = re.search(r'<title>([^<]*)</title>', html)
        page_title = title_m.group(1) if title_m else "(无标题)"
        print(f"  列表页 {page_num}: HTML长度={len(html)}, 标题={page_title[:60]}")

        pattern = r'<a[^>]*href="([^"]+)"[^>]*>([^<]*离岛免税[^<]*)</a>'
        matches = re.findall(pattern, html)

        if not matches:
            pattern2 = r'href="([^"]+)"[^>]*>([^<]*(?:离岛免税|免税销售|免税购物)[^<]*)</a>'
            matches = re.findall(pattern2, html)

        if not matches:
            print(f"  列表页 {page_num}: 未找到免税相关链接")
            no_match_streak += 1
            if no_match_streak >= 2:
                break
            continue

        no_match_streak = 0
        print(f"  列表页 {page_num}: 找到 {len(matches)} 个相关链接")
        page_has_articles = False
        for href, title in matches:
            m = re.search(r"(\d{4})年(\d{1,2})月", title)
            if m:
                year = m.group(1)
                month = m.group(2).zfill(2)
                key = f"{year}-{month}"
                if key not in articles:
                    articles[key] = href
                    print(f"    新增: {key} -> {href}")
                page_has_articles = True

        if not page_has_articles:
            print(f"  列表页 {page_num}: 无日期匹配的链接，停止翻页")
            break

        await asyncio.sleep(1)

    return articles


async def get_xlsx_url(article_path):
    """获取文章页中的 xlsx 下载链接（curl 优先，nodriver 兜底）"""
    if article_path.startswith("/"):
        urls_to_try = [
            BASE_URL_HTTP + article_path,   # HTTP 优先（WAF 对 https 限制更严）
            BASE_URL + article_path,         # HTTPS 兜底
        ]
    else:
        urls_to_try = [article_path]

    referer = BASE_URL_HTTP + "/haikou_customs/605737/fdzdgknr82/605745/58527f05-1.html"

    for full_url in urls_to_try:
        protocol = "https" if "https" in full_url else "http"
        print(f"    尝试 {protocol}...")

        # curl 方式
        html = curl_fetch_html(full_url, referer=referer)
        if html:
            url = _extract_xlsx_from_html(html)
            if url:
                # 统一用 http 协议下载（实测 http 更容易通过 WAF）
                url = url.replace("https://haikou", "http://haikou")
                print(f"    ✅ {protocol} 成功找到 xlsx 链接")
                return url

    # curl 全部失败，用 nodriver 兜底
    for full_url in urls_to_try:
        html = await nw_fetch_html(full_url, wait_sec=10)
        if html:
            url = _extract_xlsx_from_html(html)
            if url:
                url = url.replace("https://haikou", "http://haikou")
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
    """保存数据到 data.json（纯数组格式，保持与 dashboard 兼容）"""
    data.sort(key=lambda x: x["month"])
    with open(DATA_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    print(f"已保存 {len(data)} 条记录到 data.json")


# ─── 抓取逻辑 ──────────────────────────────────────────────────


async def scrape_month(month_key, article_info, force=False):
    """抓取单个月份的数据

    article_info 可以是:
      - dict: {"article": "/path/...", "xlsx": "http://...xlsx"}
      - str:  "/path/..."  (兼容旧格式)
    """
    existing = load_existing_data()
    existing_months = {d["month"] for d in existing}

    if not force and month_key in existing_months:
        print(f"  {month_key}: 已存在，跳过 (使用 --force 强制更新)")
        return True

    # 兼容旧格式
    if isinstance(article_info, str):
        article_info = {"article": article_info}

    article_path = article_info.get("article", "")
    direct_xlsx = article_info.get("xlsx")

    # ── 获取 xlsx 下载链接 ──
    xlsx_url = None

    # 优先级 1: 直接 xlsx URL（绕过 WAF，GitHub Actions 可用）
    if direct_xlsx:
        print(f"  {month_key}: 使用已知 xlsx 直链")
        xlsx_url = direct_xlsx

    # 优先级 2: 从文章页获取 xlsx 链接（需要浏览器访问）
    if not xlsx_url and article_path:
        print(f"  {month_key}: 正在从文章页获取下载链接...")
        xlsx_url = await get_xlsx_url(article_path)

    if not xlsx_url:
        print(f"  {month_key}: 未找到xlsx下载链接")
        return False

    # ── 下载 ──
    ext = ".xlsx" if xlsx_url.endswith(".xlsx") else ".xls"
    tmp_file = os.path.join(DATA_DIR, f"{month_key}{ext}")

    referer = (BASE_URL_HTTP + article_path if article_path.startswith("/") else article_path) if article_path else None
    print(f"  {month_key}: 正在下载...")
    if not curl_download(xlsx_url, tmp_file, referer=referer):
        print(f"  {month_key}: 下载失败")
        return False

    # ── 解析 ──
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


def _normalize_discovered(discovered):
    """将 discover_articles_from_listing 返回的 {month: path} 转为 {month: {article: path}}"""
    return {k: {"article": v} if isinstance(v, str) else v for k, v in discovered.items()}


async def scrape_all(force=False):
    """抓取所有已知月份"""
    os.makedirs(DATA_DIR, exist_ok=True)

    articles = {k: v.copy() for k, v in KNOWN_ARTICLES.items()}

    # 尝试从列表页发现新文章
    print("正在从海关官网发现新数据...")
    discovered = await discover_articles_from_listing()
    discovered = _normalize_discovered(discovered)
    new_count = sum(1 for k in discovered if k not in KNOWN_ARTICLES)
    articles.update(discovered)

    print(f"共发现 {len(articles)} 个月份 (其中 {new_count} 个新增)")
    print()

    success = 0
    fail = 0
    for month_key in sorted(articles.keys()):
        if await scrape_month(month_key, articles[month_key], force=force):
            success += 1
        else:
            fail += 1
        await asyncio.sleep(0.5)

    print(f"\n完成: 成功 {success}, 失败 {fail}")


async def async_main(args):
    """异步主函数"""
    try:
        if args.discover:
            articles = await discover_articles_from_listing()
            for k, v in sorted(articles.items()):
                print(f"  {k}: {v}")
            return

        if args.month:
            os.makedirs(DATA_DIR, exist_ok=True)
            if args.month in KNOWN_ARTICLES:
                await scrape_month(
                    args.month, KNOWN_ARTICLES[args.month], force=True
                )
            else:
                # 尝试动态发现该月份
                print(f"未知月份 {args.month}，尝试从海关官网查找...")
                discovered = await discover_articles_from_listing()
                discovered = _normalize_discovered(discovered)
                if args.month in discovered:
                    await scrape_month(args.month, discovered[args.month], force=True)
                else:
                    print(f"未找到 {args.month} 的数据")
                    print("已知月份:", ", ".join(sorted(KNOWN_ARTICLES.keys())))
        else:
            await scrape_all(force=args.force)
    finally:
        await close_browser()


def main():
    parser = argparse.ArgumentParser(description="海南离岛免税销售数据爬虫")
    parser.add_argument("--force", action="store_true", help="强制重新抓取所有月份")
    parser.add_argument("--month", type=str, help="只抓取指定月份 (格式: 2026-02)")
    parser.add_argument(
        "--discover", action="store_true", help="只发现新文章链接，不下载"
    )
    args = parser.parse_args()

    # 手动管理事件循环，避免 asyncio.run() 提前关闭循环导致
    # nodriver 子进程 transport 的 __del__ 回调报 "Event loop is closed"
    import gc
    loop = asyncio.new_event_loop()
    asyncio.set_event_loop(loop)
    try:
        loop.run_until_complete(async_main(args))
    finally:
        # 强制 GC 回收 nodriver 的子进程 transport，此时循环仍可用
        gc.collect()
        try:
            loop.run_until_complete(asyncio.sleep(0.1))
        except RuntimeError:
            pass
        loop.close()


if __name__ == "__main__":
    main()

