# 海南离岛免税销售数据

自动从[海口海关官网](http://haikou.customs.gov.cn/haikou_customs/605737/fdzdgknr82/605745/58527f05-1.html)抓取月度免税销售数据，提供可视化看板。

## 添加新月份数据

海关网站有 WAF 防护，GitHub Actions 无法访问文章页，但 **xlsx 文件下载不受限制**。

每月海关发布新数据后，按以下步骤操作：

### 1. 获取 xlsx 直链

1. 用浏览器打开海关文章页（如 `http://haikou.customs.gov.cn/.../7117655/index.html`）
2. 点击页面上的 xlsx 下载链接
3. 复制下载请求的完整 URL（可通过 F12 → Network 面板拦截）

### 2. 添加到 scrape.py

在 `scrape.py` 的 `KNOWN_ARTICLES` 字典末尾添加一行，如：

```python
"2026-04": "http://haikou.customs.gov.cn/haikou_customs/605737/fdzdgknr82/605745/XXXXXXX/YYYYYYYYYYYYYYYYYYY.xlsx",
```

### 3. 推送

```bash
git add scrape.py
git commit -m "data: add 2026-04"
git push
```

GitHub Actions 会自动运行，下载 xlsx 并更新 `data.json`。

---

## 文件说明

| 文件 | 说明 |
|------|------|
| `scrape.py` | 数据爬虫脚本，从海口海关下载 xlsx 并解析 |
| `data.json` | 解析后的结构化数据（所有月份汇总） |
| `dashboard.html` | 可视化看板（Chart.js），本地打开即可查看 |
| `data/` | 原始 xlsx/xls 文件存档 |

## 本地运行

```bash
pip install -r requirements.txt
python3 scrape.py                  # 抓取所有缺失月份
python3 scrape.py --force          # 强制重新抓取
python3 scrape.py --month 2026-02  # 抓取指定月份
```

> 本地运行时，对于没有 xlsx 直链的旧月份，会尝试从文章页获取下载链接。

## 自动更新

GitHub Actions 每月 17-22 日每 12 小时自动运行，抓取已有 xlsx 直链的月份并提交到仓库。

## 数据指标

- **购物金额**（万元）
- **购物实际人次**（万人次）
- **购物件数**（万件）
- **人均消费**（元/人次）
- 各指标的**同比变化率**和**年度累计值**
