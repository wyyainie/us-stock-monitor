# 美股事件监控（US Stock Monitor）

基于 **SEC EDGAR** 的美股公告监控与查询工具，用 Python + Tkinter 写的桌面小工具。抓的是 SEC 官方披露（8-K / 6-K / 424B / S-1 / S-3），不依赖行情源，专注「合股/拆股、公司更名、增发」这几类事件。

## 功能

- **市场扫描**（`market_scan_gui.py` / `sec_events.py`）
  - 15 条关键词 × 表单类型（8-K、6-K、424B2/3/5、S-1、S-3）走 SEC 全文检索 API
  - 按 Items + 关键词分类：合股/拆股、公司更名、增发
  - 0~100 打分排序，导出 Excel（全部 + 分类 Sheet）
  - 支持「近 N 天」时间窗（默认 30 天）
- **单股查询**（`query_gui.py` / `query_tool.py`）
  - 输入股票代码，查该股在指定窗口内的相关公告
  - CLI 版 `query_tool.py` 方便脚本化
- **自选股**（`watchlist.py`）——自定义关注列表
- **CIK 映射缓存**：首次运行下载 SEC 全量股票表并缓存一天

## 安装

```bash
pip install -r requirements.txt
```

Python 3.8+（Tkinter 为标准库，Linux 上可能需要单独装 `python3-tk`）。

> **重要**：运行前把 `sec_events.py` 里的 `SEC_USER_AGENT` 改成你自己的邮箱。
> SEC 要求请求方提供可联系的 User-Agent，占位符会被限流甚至封禁。

## 使用

```bash
# 图形界面：市场扫描
python market_scan_gui.py

# 图形界面：单股查询
python query_gui.py

# 命令行查询
python query_tool.py AAPL --days 30
```

## 目录结构

```
sec_events.py      # 核心引擎：SEC 全文检索 + 清洗 + 打分 + Excel 导出
market_scan_gui.py # 市场扫描 GUI（批量扫描多只股票的事件）
query_gui.py       # 单股查询 GUI
query_tool.py      # 单股查询 CLI
watchlist.py       # 自选股读写
main.py            # 入口占位
requirements.txt   # 依赖
```

## 说明与限制

- 数据来自 SEC EDGAR 公开检索接口，**不是实时行情**；全文检索索引本身有几分钟到几十分钟延迟。
- 每类事件用关键词 + Items 打分，不保证召回全部公告；只是筛选候选，不构成投资建议。
- 交易所白名单（NYSE / Nasdaq 等）目前**未启用**，可能混入 OTC 标的。
- SEC 接口有访问频率限制，脚本内置了节流与指数退避，请勿把并发拉高。

## License

MIT