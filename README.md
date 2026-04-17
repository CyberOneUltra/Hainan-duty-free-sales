# 海南离岛免税销售数据

自动从[海口海关官网](http://haikou.customs.gov.cn/haikou_customs/605737/fdzdgknr82/605745/58527f05-1.html)抓取月度免税销售数据，提供可视化看板。

## 文件说明

| 文件 | 说明 |
|------|------|
| `scrape.py` | 数据爬虫脚本，从海口海关下载 xlsx 并解析 |
| `data.json` | 解析后的结构化数据（所有月份汇总） |
| `dashboard.html` | 可视化看板（Chart.js），本地打开即可查看 |
| `data/` | 原始 xlsx/xls 文件存档 |

## 快速开始

```bash
pip install -r requirements.txt
python3 scrape.py            # 抓取所有缺失月份
python3 scrape.py --force    # 强制重新抓取
python3 scrape.py --month 2026-02  # 抓取指定月份
```

看板直接用浏览器打开 `dashboard.html` 即可（需同目录下的 `data.json`）。

## 自动更新

GitHub Actions 每月 17~21 日自动运行爬虫，抓取上月数据并提交到仓库。

## 数据指标

- **购物金额**（万元）
- **购物实际人次**（万人次）
- **购物件数**（万件）
- **人均消费**（元/人次）
- 各指标的**同比变化率**和**年度累计值**
