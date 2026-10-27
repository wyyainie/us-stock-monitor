# -*- coding: utf-8 -*-
"""
SEC EDGAR 爬虫 v3
- 支持 8-K（美国公司）+ 6-K（外国私人发行人 / 中概股）
- 支持增发相关表单：424B2/3/5、S-1、S-3
- 事件类型：合股拆股 / 公司更名 / 增发
- 只输出 1 个 Excel（多 Sheet）
- 用打分制代替硬性 Item 过滤，避免数据被误删
"""

import os
import re
import json
import time
import logging
from datetime import datetime, timedelta
from typing import List, Dict, Optional, Tuple

import requests
import pandas as pd
from tenacity import (
    retry, stop_after_attempt, wait_exponential,
    retry_if_exception_type,
)

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
)
logger = logging.getLogger(__name__)

import os
import sys

def get_base_dir():
    """获取程序所在目录（兼容 PyInstaller 打包）"""
    if getattr(sys, 'frozen', False):
        # PyInstaller 打包后的 exe：返回 exe 所在目录
        return os.path.dirname(sys.executable)
    else:
        # 开发环境：返回 .py 文件所在目录
        return os.path.dirname(os.path.abspath(__file__))

BASE_DIR = get_base_dir()

# ==================== 配置 ====================
# ⚠️ 改成你自己的邮箱，SEC 强制要求
SEC_USER_AGENT = "StockEventBot/1.0 (your_email@example.com)"

SEC_SEARCH_API = "https://efts.sec.gov/LATEST/search-index"
SEC_TICKERS_URL = "https://www.sec.gov/files/company_tickers_exchange.json"

# ⭐ 关键修复：加上 6-K（中概股）+ 增发表单
FORMS = (
    "8-K,8-K/A,"          # 美国公司重大事件
    "6-K,6-K/A,"          # 外国私人发行人（中概股）
    "424B2,424B3,424B5,"  # 招股书（增发）
    "S-1,S-1/A,S-3,S-3/A" # 注册声明（增发）
)

QUERIES = [
    # 合股/拆股
    '"reverse stock split"',
    '"forward stock split"',
    '"stock split"',
    '"share consolidation"',
    '"reverse split"',
    # 改名
    '"name change"',
    '"change of name"',
    '"change its name"',
    '"change the name"',
    '"ticker symbol"',
    '"symbol change"',
    '"new ticker"',
    # 增发
    '"private placement"',
    '"registered direct offering"',
    '"public offering"',
    '"securities purchase agreement"',
    '"subscription agreement"',
]

# ⭐ 检索窗口（改这里）
DAYS_BACK = 30
MAX_PER_QUERY = 500

# 美股交易所白名单
US_EXCHANGES = {"NYSE", "Nasdaq", "NYSE American", "NYSE Arca", "Cboe", "AMEX"}

# Item 中文对照
ITEM_CN = {
    "1.01": "重大协议",
    "1.02": "重大协议终止",
    "1.03": "破产",
    "2.01": "资产收购/处置完成",
    "2.02": "经营成果",
    "2.03": "直接财务义务",
    "2.04": "加速义务事件",
    "2.05": "退出/处置成本",
    "2.06": "重大减值",
    "3.01": "退市/未达上市标准",
    "3.02": "非公开销售股权（增发）",
    "3.03": "重大修改证券持有人权利",
    "4.01": "变更会计师",
    "4.02": "不依赖先前财报",
    "5.01": "控制权变更",
    "5.02": "高管/董事变更",
    "5.03": "章程修改（改名/合股）",
    "5.04": "员工福利计划暂停",
    "5.05": "道德准则修订",
    "5.06": "公司状态变更",
    "5.07": "股东表决",
    "6.01": "ABS信息",
    "7.01": "公平披露",
    "8.01": "其他事项",
    "9.01": "财务报表和附件",
}

EVENT_CN = {
    "split": "合股/拆股",
    "name_change": "公司更名",
    "offering": "增发",
    "other": "其他",
}

# 缓存
TICKER_CACHE = os.path.join(BASE_DIR, "sec_tickers_cache.json")
TICKER_CACHE_DAYS = 1


# ==================== 美股映射 ====================
def load_us_stock_map(cache_file: str = TICKER_CACHE,
                      cache_days: int = TICKER_CACHE_DAYS) -> Dict:
    if os.path.exists(cache_file):
        mtime = os.path.getmtime(cache_file)
        if time.time() - mtime < cache_days * 86400:
            logger.info("使用缓存的股票列表: %s", cache_file)
            with open(cache_file, "r", encoding="utf-8") as f:
                return json.load(f)

    logger.info("首次下载 SEC 全量股票列表（约 5MB）...")
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

    with open(cache_file, "w", encoding="utf-8") as f:
        json.dump(mapping, f)
    logger.info("已缓存 %d 只股票", len(mapping))
    return mapping


# ==================== display_name 解析 ====================
def parse_display_name(name: str) -> Tuple[str, str]:
    """
    "reAlpha Tech Corp.  (AIRE)  (CIK 0001859199)"
      → ("reAlpha Tech Corp.", "AIRE")
    """
    if not name:
        return "", ""
    name = re.sub(r"\s*\(CIK\s+\d+\)\s*$", "", name).strip()
    m = re.search(r"\(([^)]+)\)\s*$", name)
    if m:
        return name[:m.start()].strip(), m.group(1).strip()
    return name.strip(), ""


# ==================== 爬虫 ====================
class SECEdgarScraper:

    def __init__(self, user_agent: str = SEC_USER_AGENT):
        self.session = requests.Session()
        self.session.headers.update({
            "User-Agent": user_agent,
            "Accept": "application/json, text/plain, */*",
            "Accept-Language": "en-US,en;q=0.9",
            "Origin": "https://www.sec.gov",
            "Referer": "https://www.sec.gov/edgar/search/",
        })
        self._last_request_ts = 0.0

    def _throttle(self):
        elapsed = time.time() - self._last_request_ts
        if elapsed < 0.15:
            time.sleep(0.15 - elapsed)
        self._last_request_ts = time.time()

    @retry(
        stop=stop_after_attempt(4),
        wait=wait_exponential(multiplier=1.5, min=2, max=20),
        retry=retry_if_exception_type(requests.RequestException),
        reraise=True,
    )
    def _get_json(self, url: str, params: dict) -> Optional[dict]:
        self._throttle()
        resp = self.session.get(url, params=params, timeout=20)
        if resp.status_code == 200:
            try:
                return resp.json()
            except ValueError:
                logger.error("非 JSON: %s", resp.text[:200])
                return None
        if resp.status_code in (403, 429):
            raise requests.RequestException(f"rate limited {resp.status_code}")
        logger.warning("SEC status=%s", resp.status_code)
        return None

    def _parse_hit(self, h: dict, query: str) -> Optional[Dict]:
        src = h.get("_source") or {}
        ciks = src.get("ciks") or []
        names = src.get("display_names") or []
        if not ciks:
            return None

        cik_raw = str(ciks[0])
        cik = cik_raw.zfill(10) if cik_raw.isdigit() else cik_raw

        adsh = src.get("adsh") or ""
        accession = adsh.replace("-", "")
        url = ("https://www.sec.gov/Archives/edgar/data/"
               f"{int(cik)}/{accession}/") if adsh else ""

        company, ticker = parse_display_name(names[0] if names else "")

        return {
            "company": company,
            "ticker": ticker,
            "cik": cik,
            "form_type": src.get("form_type") or "",
            "items": src.get("items") or [],
            "file_date": src.get("file_date") or "",
            "url": url,
            "query": query,
        }

    def search(self, query: str, start_date: str, end_date: str,
               forms: str = FORMS, max_results: int = MAX_PER_QUERY) -> List[Dict]:
        results, offset, page_size = [], 0, 10
        while offset < max_results:
            params = {
                "q": query,
                "forms": forms,
                "dateRange": "custom",
                "startdt": start_date,
                "enddt": end_date,
                "from": offset,
            }
            data = self._get_json(SEC_SEARCH_API, params)
            if not data:
                break
            hits = (data.get("hits") or {}).get("hits") or []
            if not hits:
                break
            for h in hits:
                item = self._parse_hit(h, query)
                if item:
                    results.append(item)
            offset += len(hits)
            if len(hits) < page_size:
                break
            time.sleep(0.2)
        return results

    def batch_search(self, queries: List[str], days_back: int = DAYS_BACK) -> List[Dict]:
        end_dt = datetime.now()
        start_dt = end_dt - timedelta(days=days_back)
        start_date = start_dt.strftime("%Y-%m-%d")
        end_date = end_dt.strftime("%Y-%m-%d")

        all_items, seen = [], set()
        for q in queries:
            logger.info("检索: %s [%s ~ %s]", q, start_date, end_date)
            try:
                items = self.search(q, start_date, end_date)
            except Exception as e:
                logger.error("失败 %s: %s", q, e)
                continue
            logger.info("  命中 %d 条", len(items))
            for it in items:
                key = (it["cik"], it["url"])
                if key in seen:
                    continue
                seen.add(key)
                all_items.append(it)
            time.sleep(0.2)
        return all_items


# ==================== 分类 & 打分 ====================
SPLIT_PAT = re.compile(
    r"reverse stock split|forward stock split|stock split|"
    r"share consolidation|reverse split|forward split",
    re.IGNORECASE,
)
NAME_PAT = re.compile(
    r"name change|change of name|change its name|change the name|"
    r"ticker symbol|symbol change|new ticker|rebrand|renam",
    re.IGNORECASE,
)
OFFERING_PAT = re.compile(
    r"private placement|registered direct offering|public offering|"
    r"securities purchase agreement|subscription agreement|"
    r"underwritten offering|at-the-market|atm offering",
    re.IGNORECASE,
)


def classify(record: Dict) -> str:
    """按 items 优先，其次按 query 判断事件类型"""
    query = record.get("query", "") or ""
    items = set(record.get("items") or [])

    # 先看 items（最准）
    if "3.02" in items:
        return "offering"
    if "5.03" in items or "3.03" in items:
        if SPLIT_PAT.search(query):
            return "split"
        if NAME_PAT.search(query):
            return "name_change"

    # 再看 query
    if OFFERING_PAT.search(query):
        return "offering"
    if SPLIT_PAT.search(query):
        return "split"
    if NAME_PAT.search(query):
        return "name_change"
    return "other"


def score_record(record: Dict) -> int:
    """0~100 打分，越高越可能是真事件"""
    score = 0
    query = record.get("query", "") or ""
    items = set(record.get("items") or [])
    form = record.get("form_type", "") or ""

    # Item 命中加分（最重要）
    if "5.03" in items:
        score += 40
    if "3.03" in items:
        score += 30
    if "3.02" in items:
        score += 30
    if "5.07" in items:
        score += 15
    if "1.01" in items:
        score += 10

    # 表单类型加分
    if form in ("8-K", "6-K"):
        score += 20
    elif form in ("424B5", "S-1", "S-3"):
        score += 15

    # query 关键词加分
    if SPLIT_PAT.search(query) or NAME_PAT.search(query) or OFFERING_PAT.search(query):
        score += 20

    return min(100, score)


def items_to_cn(items_list) -> str:
    if not items_list:
        return ""
    if isinstance(items_list, str):
        items_list = re.findall(r"\d+\.\d+", items_list)
    return "；".join(ITEM_CN.get(i, i) for i in items_list)


# ==================== 清洗 ====================
def clean(records: List[Dict], us_map: Dict) -> pd.DataFrame:
    cols = ["日期", "事件类型", "股票代码", "公司名称", "交易所",
            "重要度", "表格类型", "涉及事项", "SEC链接"]

    if not records:
        return pd.DataFrame(columns=cols)

    df = pd.DataFrame(records)

    # 1. 分类
    df["event_type_raw"] = df.apply(classify, axis=1)
    df = df[df["event_type_raw"].isin(["split", "name_change", "offering"])].copy()
    if df.empty:
        return pd.DataFrame(columns=cols)

    # 2. 打分
    df["score"] = df.apply(score_record, axis=1)

    # 3. 补全 ticker / exchange / 公司名
    def get_us_info(row):
        cik = str(row.get("cik", "")).zfill(10)
        info = us_map.get(cik)
        if info:
            return pd.Series({
                "ticker_new": info["ticker"],
                "exchange_new": info["exchange"],
                "company_new": info["name"],
            })
        return pd.Series({
            "ticker_new": row.get("ticker", ""),
            "exchange_new": "",
            "company_new": row.get("company", ""),
        })

    us_info = df.apply(get_us_info, axis=1)
    df["ticker"] = us_info["ticker_new"]
    df["exchange"] = us_info["exchange_new"]
    df["company"] = us_info["company_new"]

    # 4. 去重
    df = df.drop_duplicates(subset=["cik", "url"], keep="first")

    # 5. 翻译列
    df["日期"] = pd.to_datetime(df["file_date"], errors="coerce").dt.strftime("%Y-%m-%d")
    df["事件类型"] = df["event_type_raw"].map(EVENT_CN)
    df["股票代码"] = df["ticker"].fillna("").astype(str)
    df["公司名称"] = df["company"].fillna("").astype(str)
    df["交易所"] = df["exchange"].fillna("").astype(str)
    df["重要度"] = df["score"].apply(lambda s: "⭐" * min(5, s // 20) if s >= 40 else "")
    df["表格类型"] = df["form_type"].fillna("").astype(str)
    df["涉及事项"] = df["items"].apply(items_to_cn)
    df["SEC链接"] = df["url"]

    # ⭐ 关键修复：先用 score 排序，再选 cols
    df = df.sort_values(by=["score", "日期"], ascending=[False, False])
    df = df.drop(columns=["score"], errors="ignore").reset_index(drop=True)

    # 兜底
    for c in cols:
        if c not in df.columns:
            df[c] = ""

    return df[cols]


# ==================== 输出（只出 1 个 Excel） ====================
def save_output(df: pd.DataFrame, days_back: int) -> str:
    out = os.path.join(BASE_DIR, "美股事件_近{}天.xlsx".format(days_back))
    try:
        with pd.ExcelWriter(out, engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name="全部", index=False)
            for event_cn in ["合股/拆股", "公司更名", "增发"]:
                sub = df[df["事件类型"] == event_cn]
                if not sub.empty:
                    sheet = event_cn.replace("/", "_")
                    sub.to_excel(writer, sheet_name=sheet, index=False)
        logger.info("已保存: %s", out)
    except Exception as e:
        logger.warning("Excel 导出失败: %s", e)
        out = "美股事件_近{}天.csv".format(days_back)
        df.to_csv(out, index=False, encoding="utf-8-sig")
        logger.info("已改为保存 CSV: %s", out)
    return out


# ==================== 主入口 ====================
def main():
    us_map = load_us_stock_map()

    scraper = SECEdgarScraper()
    raw = scraper.batch_search(QUERIES, days_back=DAYS_BACK)
    logger.info("原始记录: %d 条", len(raw))

    df = clean(raw, us_map)
    logger.info("清洗后: %d 条", len(df))

    if df.empty:
        print("\n⚠️ 近 {} 天没有找到记录。".format(DAYS_BACK))
        print("建议把 DAYS_BACK 改成 90 或 180 再试")
        return

    out = save_output(df, DAYS_BACK)

    # 打印
    print("\n" + "=" * 140)
    print("近 {} 天 美股更名 / 合股 / 增发 事件".format(DAYS_BACK))
    print("=" * 140)
    print("{:<12} {:<10} {:<8} {:<10} {:<12} {:<30} {}".format(
        "日期", "事件类型", "代码", "交易所", "重要度", "公司名称", "涉及事项"))
    print("-" * 140)
    for _, r in df.head(50).iterrows():
        print("{:<12} {:<10} {:<8} {:<10} {:<12} {:<30} {}".format(
            r["日期"], r["事件类型"], r["股票代码"][:8],
            r["交易所"][:10], r["重要度"], r["公司名称"][:28],
            r["涉及事项"][:40]))
    print("=" * 140)

    # 统计
    print("\n统计:")
    for et in ["合股/拆股", "公司更名", "增发"]:
        n = len(df[df["事件类型"] == et])
        print("  {}: {} 条".format(et, n))
    print("  总计: {} 条 | 输出: {}".format(len(df), out))


if __name__ == "__main__":
    main()