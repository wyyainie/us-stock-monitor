# -*- coding: utf-8 -*-
"""
美股事件查询工具
输入股票代码 → 查它近期的更名 / 合股 / 增发 事件

用法:
    python query_tool.py
"""

import os
import re
import sys
import json
import time
import logging
from datetime import datetime, timedelta

import requests
import pandas as pd

logging.basicConfig(
    level=logging.WARNING,   # 只显示警告和错误，让界面干净
    format="%(asctime)s [%(levelname)s] %(message)s",
)
logger = logging.getLogger(__name__)


# ==================== 路径处理（兼容打包） ====================
def get_base_dir():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))

BASE_DIR = get_base_dir()


# ==================== 配置 ====================
# ⚠️ 打包前改成你自己的邮箱（SEC 强制要求）
SEC_USER_AGENT = "StockEventBot/1.0 (your_email@example.com)"

TICKER_CACHE = os.path.join(BASE_DIR, "sec_tickers_cache.json")
SEC_TICKERS_URL = "https://www.sec.gov/files/company_tickers_exchange.json"

# 默认查询窗口（天）
DEFAULT_DAYS_BACK = 365

# 关注的表单
FORMS_OF_INTEREST = {"8-K", "8-K/A", "6-K", "6-K/A",
                     "424B2", "424B3", "424B5",
                     "S-1", "S-1/A", "S-3", "S-3/A"}

# 关键词
SPLIT_PAT = re.compile(
    r"reverse stock split|forward stock split|stock split|"
    r"share consolidation|reverse split|reverse share split|"
    r"consolidation of shares",
    re.IGNORECASE,
)
NAME_PAT = re.compile(
    r"name change|change of name|change (its|the|our) name|"
    r"changed (its|the|our) name|corporate name|"
    r"ticker symbol|symbol change|new ticker|"
    r"articles of amendment|certificate of amendment|"
    r"amended and restated articles|renam",
    re.IGNORECASE,
)
OFFERING_PAT = re.compile(
    r"private placement|registered direct offering|public offering|"
    r"securities purchase agreement|subscription agreement|"
    r"underwritten offering|\bPIPE\b",
    re.IGNORECASE,
)


# ==================== 股票列表 ====================
def load_ticker_map(force_refresh=False):
    """加载 cik -> {ticker, name, exchange} 映射"""
    if not force_refresh and os.path.exists(TICKER_CACHE):
        mtime = os.path.getmtime(TICKER_CACHE)
        if time.time() - mtime < 86400:  # 1 天缓存
            with open(TICKER_CACHE, "r", encoding="utf-8") as f:
                return json.load(f)

    print("首次运行，正在下载股票列表（约 5MB，10~30 秒）...")
    r = requests.get(SEC_TICKERS_URL,
                     headers={"User-Agent": SEC_USER_AGENT}, timeout=60)
    r.raise_for_status()
    data = r.json()

    fields = data.get("fields", [])
    rows = data.get("data", [])
    idx = {f: i for i, f in enumerate(fields)}

    mapping = {}
    for row in rows:
        try:
            cik = str(row[idx["cik"]]).zfill(10)
            mapping[cik] = {
                "ticker": row[idx["ticker"]] or "",
                "name": row[idx["name"]] or "",
                "exchange": row[idx["exchange"]] or "",
            }
        except Exception:
            continue

    with open(TICKER_CACHE, "w", encoding="utf-8") as f:
        json.dump(mapping, f)
    print("已缓存 {} 只股票\n".format(len(mapping)))
    return mapping


def find_cik(ticker, ticker_map):
    ticker = ticker.upper().strip()
    for cik, info in ticker_map.items():
        if info.get("ticker", "").upper() == ticker:
            return cik, info
    return None, None


# ==================== 查询逻辑 ====================
def get_submissions(cik):
    url = "https://data.sec.gov/submissions/CIK{}.json".format(cik)
    r = requests.get(url, headers={"User-Agent": SEC_USER_AGENT}, timeout=30)
    r.raise_for_status()
    return r.json()


def check_document(url):
    """下载文件正文，检查是否含关键词"""
    try:
        r = requests.get(url, headers={"User-Agent": SEC_USER_AGENT}, timeout=30)
        if r.status_code != 200:
            return [], ""
        text = re.sub(r"<[^>]+>", " ", r.text)
        text = re.sub(r"\s+", " ", text)

        events = []
        if SPLIT_PAT.search(text):
            events.append("合股/拆股")
        if NAME_PAT.search(text):
            events.append("公司更名")
        if OFFERING_PAT.search(text):
            events.append("增发")

        # 提取细节
        detail = ""
        if "合股/拆股" in events:
            m = re.search(r"(\d+)\s*[- ]for[- ](\d+)", text)
            if m:
                detail = "{}-for-{}".format(m.group(1), m.group(2))
        if "公司更名" in events and not detail:
            m = re.search(
                r"from\s+([A-Z][A-Za-z0-9&\.\-, ]{2,60}?)\s+to\s+([A-Z][A-Za-z0-9&\.\-, ]{2,60}?)[\.\,]",
                text,
            )
            if m:
                detail = "{} → {}".format(m.group(1).strip(), m.group(2).strip())
        if "增发" in events and not detail:
            m = re.search(r"\$\s*([\d,]+(?:\.\d+)?)\s*(million|billion)?", text)
            if m:
                detail = "${}{}".format(m.group(1),
                                        " " + m.group(2) if m.group(2) else "")

        return events, detail
    except Exception as e:
        logger.warning("下载失败 %s: %s", url, e)
        return [], ""


def query_ticker(ticker, ticker_map, days_back=DEFAULT_DAYS_BACK):
    """查询单个 ticker，返回记录列表"""
    cik, info = find_cik(ticker, ticker_map)
    if not cik:
        print("  ❌ 未找到股票代码: {}".format(ticker))
        print("     可能原因：代码拼错、已退市、或缓存过旧")
        return []

    print("  📌 {} (CIK {})".format(ticker, cik))
    print("     公司名: {}".format(info.get("name")))
    print("     交易所: {}".format(info.get("exchange") or "（未登记）"))
    print("     正在检查近 {} 天的文件...".format(days_back))

    try:
        data = get_submissions(cik)
    except Exception as e:
        print("  ❌ 拉取文件列表失败: {}".format(e))
        return []

    filings = data.get("filings", {}).get("recent", {})
    forms = filings.get("form", [])
    dates = filings.get("filingDate", [])
    accessions = filings.get("accessionNumber", [])
    primary_docs = filings.get("primaryDocument", [])
    items_arr = filings.get("items", [])

    cik_int = str(int(cik))
    end_dt = datetime.now()
    start_dt = end_dt - timedelta(days=days_back)

    results = []

    for i, form in enumerate(forms):
        if form not in FORMS_OF_INTEREST:
            continue

        date_str = dates[i]
        try:
            date_obj = datetime.strptime(date_str, "%Y-%m-%d")
        except Exception:
            continue
        if date_obj < start_dt:
            continue

        acc = accessions[i].replace("-", "")
        doc = primary_docs[i]
        items = items_arr[i] if i < len(items_arr) else ""
        url = "https://www.sec.gov/Archives/edgar/data/{}/{}/{}".format(
            cik_int, acc, doc)

        # 先看 items（8-K 有）
        events = []
        if items:
            if "5.03" in items:
                events.append("公司更名")
            if "3.03" in items:
                events.append("合股/拆股")
            if "3.02" in items:
                events.append("增发")

        detail = ""

        # items 没信息 → 下载正文检查
        if not events:
            time.sleep(0.3)
            events, detail = check_document(url)

        if events:
            print("  ✅ {} | {:<6} | {:<12} | {}".format(
                date_str, form, ", ".join(events), detail))
            results.append({
                "股票代码": ticker,
                "公司名称": info.get("name", ""),
                "交易所": info.get("exchange", ""),
                "日期": date_str,
                "事件类型": ", ".join(events),
                "细节": detail,
                "表单": form,
                "Items": items,
                "SEC链接": url,
            })

    if not results:
        print("  ⚠️ 近 {} 天没有找到相关事件".format(days_back))
        print("     建议：加大窗口（改 query_tool.py 里 DEFAULT_DAYS_BACK）")
    else:
        print("  → 共命中 {} 条".format(len(results)))

    return results


# ==================== 输出 ====================
def save_and_open(df, ticker):
    if df.empty:
        return

    filename = "查询结果_{}.xlsx".format(ticker)
    out = os.path.join(BASE_DIR, filename)

    try:
        df.to_excel(out, index=False, engine="openpyxl")
    except Exception as e:
        logger.warning("Excel 保存失败: %s", e)
        out = os.path.join(BASE_DIR, "查询结果_{}.csv".format(ticker))
        df.to_csv(out, index=False, encoding="utf-8-sig")

    print("  💾 已保存: {}".format(out))

    # 自动打开
    try:
        if sys.platform == "win32":
            os.startfile(out)
        elif sys.platform == "darwin":
            os.system('open "{}"'.format(out))
        else:
            os.system('xdg-open "{}"'.format(out))
    except Exception as e:
        logger.warning("自动打开失败: %s", e)


# ==================== 主入口 ====================
def main():
    print("=" * 60)
    print("  美股事件查询工具")
    print("  （更名 / 合股 / 增发）")
    print("=" * 60)
    print()

    # 加载股票列表
    ticker_map = load_ticker_map()

    while True:
        try:
            raw = input("请输入股票代码（多个用逗号分隔，直接回车退出）: ").strip()
        except (EOFError, KeyboardInterrupt):
            print("\n再见！")
            break

        if not raw:
            print("再见！")
            break

        # 支持逗号 / 空格 / 分号分隔
        tickers = [t.strip().upper() for t in re.split(r"[,;\s]+", raw) if t.strip()]
        if not tickers:
            continue

        print()
        all_results = []
        for tk in tickers:
            print("-" * 60)
            results = query_ticker(tk, ticker_map)
            all_results.extend(results)
            print()

        if not all_results:
            print("⚠️ 本次查询无结果\n")
            continue

        df = pd.DataFrame(all_results)
        df = df.sort_values("日期", ascending=False).reset_index(drop=True)

        # 多个 ticker 合并保存，单个单独保存
        if len(tickers) == 1:
            save_and_open(df, tickers[0])
        else:
            save_and_open(df, "_".join(tickers[:3]))

        print()


if __name__ == "__main__":
    main()