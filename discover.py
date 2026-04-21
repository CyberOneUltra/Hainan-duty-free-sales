#!/usr/bin/env python3
"""
自动发现海口海关新增离岛免税数据
从海关列表页发现新文章，提取 xlsx 直链，更新 KNOWN_ARTICLES。

由 OpenClaw cron 每月 18-22 日运行。
输出：新发现的月份信息（供 cron job 回报给用户）
"""

import re
import sys
import json
import asyncio
import subprocess
from pathlib import Path

try:
    import nodriver as uc
except ImportError:
    print("需要 nodriver: pip install nodriver")
    sys.exit(1)

BASE_URL = "https://haikou.customs.gov.cn"
LISTING_URL = BASE_URL + "/haikou_customs/605737/fdzdgknr82/605745/58527f05-{page}.html"
REPO_DIR = Path(__file__).parent
SCRAPE_PY = REPO_DIR / "scrape.py"


async def discover():
    """发现新文章并提取 xlsx 直链"""
    browser = await uc.start(
        headless=True,
        sandbox=False,
        browser_args=["--no-sandbox", "--disable-dev-shm-usage", "--disable-gpu"],
    )

    try:
        # 加载已知月份
        known = set()
        if SCRAPE_PY.exists():
            content = SCRAPE_PY.read_text()
            known = set(re.findall(r'"(\d{4}-\d{2})":', content))

        new_articles = {}

        for page_num in range(1, 5):
            url = LISTING_URL.format(page=page_num)
            page = await browser.get(url)
            await asyncio.sleep(8)
            html = await page.get_content()

            if not html or ("504" in html and "连接超时" in html):
                print(f"列表页 {page_num}: 获取失败")
                break

            pattern = r'href="([^"]+)"[^>]*>([^<]*离岛免税销售情况表[^<]*)</a>'
            matches = re.findall(pattern, html)

            for href, title in matches:
                m = re.search(r"(\d{4})年(\d{1,2})月", title)
                if not m:
                    continue
                month_key = f"{m.group(1)}-{m.group(2).zfill(2)}"
                if month_key not in known and month_key not in new_articles:
                    new_articles[month_key] = href
                    print(f"发现新月份: {month_key} -> {href}")

            await asyncio.sleep(1)

        if not new_articles:
            print("没有发现新月份")
            return {}

        # 提取每个新文章的 xlsx 直链
        results = {}
        for month_key, article_path in new_articles.items():
            full_url = BASE_URL + article_path if article_path.startswith("/") else article_path
            page = await browser.get(full_url)
            await asyncio.sleep(5)
            html = await page.get_content()

            m = re.search(r'href="([^"]+\.xls[x]?)"', html)
            if m:
                xlsx_url = m.group(1)
                if not xlsx_url.startswith("http"):
                    xlsx_url = BASE_URL + xlsx_url
                # 统一用 http
                xlsx_url = xlsx_url.replace("https://haikou", "http://haikou")
                results[month_key] = {
                    "article": article_path,
                    "xlsx": xlsx_url,
                }
                print(f"  {month_key}: xlsx = {xlsx_url}")
            else:
                results[month_key] = {"article": article_path}
                print(f"  {month_key}: 未找到 xlsx 链接")

            await asyncio.sleep(1)

        return results

    finally:
        browser.stop()


def update_scrape_py(new_entries):
    """将新发现的条目写入 scrape.py 的 KNOWN_ARTICLES"""
    if not new_entries:
        return False

    content = SCRAPE_PY.read_text()

    # 找到 KNOWN_ARTICLES 字典的结束位置
    # 在最后一个条目之后、"}" 之前插入
    last_brace = content.rfind("}", 0, content.find("_normalize_articles"))

    insert_lines = []
    for month_key in sorted(new_entries.keys()):
        info = new_entries[month_key]
        if "xlsx" in info:
            insert_lines.append(
                f'    "{month_key}": {{\n'
                f'        "article": "{info["article"]}",\n'
                f'        "xlsx": "{info["xlsx"]}",\n'
                f"    }},"
            )
        else:
            insert_lines.append(
                f'    "{month_key}": {{"article": "{info["article"]}"}},'
            )

    insert_text = "\n" + "\n".join(insert_lines) + "\n"
    content = content[:last_brace] + insert_text + content[last_brace:]
    SCRAPE_PY.write_text(content)
    return True


def git_push(new_entries):
    """提交并推送更新"""
    months = ", ".join(sorted(new_entries.keys()))
    cmds = [
        f"cd {REPO_DIR} && git add scrape.py",
        f'cd {REPO_DIR} && git commit -m "data: auto-discover {months}"',
        f"cd {REPO_DIR} && git push origin main",
    ]
    for cmd in cmds:
        result = subprocess.run(cmd, shell=True, capture_output=True, text=True)
        if result.returncode != 0 and "nothing to commit" not in result.stderr:
            print(f"git error: {result.stderr}")
            return False
    return True


async def main():
    print("=== 海口海关数据自动发现 ===\n")
    new_entries = await discover()

    if not new_entries:
        print("\n无新数据需要更新")
        return

    print(f"\n发现 {len(new_entries)} 个新月份，正在更新 scrape.py...")
    if update_scrape_py(new_entries):
        if git_push(new_entries):
            print("✅ 已更新并推送")
            # 触发 GitHub Actions 抓取新数据
            for month_key in new_entries:
                print(f"建议手动触发: gh workflow run scrape.yml -f month={month_key}")
        else:
            print("⚠️ 更新了文件但 git push 失败")
    else:
        print("⚠️ 更新 scrape.py 失败")


if __name__ == "__main__":
    asyncio.run(main())
