# -*- coding: utf-8 -*-
"""
美股市场事件扫描工具（图形界面版）
- 扫描全市场近 N 天的更名 / 合股 / 增发事件
- 时间窗口可选：1天 / 1周 / 1月 / 3月 / 1年
- 结果自动保存 Excel 并打开
"""

import os
import re
import sys
import json
import time
import threading
import logging
from datetime import datetime, timedelta

import requests
import pandas as pd

import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext

logging.basicConfig(level=logging.WARNING)
logger = logging.getLogger(__name__)


# ==================== 路径处理 ====================
def get_base_dir():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))

BASE_DIR = get_base_dir()


# ==================== 配置 ====================
SEC_USER_AGENT = "StockEventBot/1.0 (your_email@example.com)"  # ⚠️ 改成你的邮箱
SEC_SEARCH_API = "https://efts.sec.gov/LATEST/search-index"
SEC_TICKERS_URL = "https://www.sec.gov/files/company_tickers_exchange.json"
TICKER_CACHE = os.path.join(BASE_DIR, "sec_tickers_cache.json")

# 表单
FORMS = (
    "8-K,8-K/A,"
    "6-K,6-K/A,"
    "424B2,424B3,424B5,"
    "S-1,S-1/A,S-3,S-3/A"
)

# 检索关键词
QUERIES = [
    # 合股/拆股
    '"reverse stock split"',
    '"forward stock split"',
    '"stock split"',
    '"share consolidation"',
    '"reverse split"',
    '"reverse share split"',
    '"consolidation of shares"',
    # 改名
    '"name change"',
    '"change of name"',
    '"change its name"',
    '"change the name"',
    '"change our name"',
    '"changed its name"',
    '"changed the name"',
    '"corporate name"',
    '"ticker symbol"',
    '"symbol change"',
    '"new ticker"',
    '"articles of amendment"',
    '"certificate of amendment"',
    # 增发
    '"private placement"',
    '"registered direct offering"',
    '"public offering"',
    '"securities purchase agreement"',
    '"subscription agreement"',
    '"definitive agreement"',
    '"PIPE financing"',
]

MAX_PER_QUERY = 500

# 美股交易所白名单
US_EXCHANGES = {"NYSE", "Nasdaq", "NYSE American", "NYSE Arca", "Cboe", "AMEX"}

# Item 中文
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


# ==================== 美股映射 ====================
def load_us_stock_map(force_refresh=False):
    if not force_refresh and os.path.exists(TICKER_CACHE):
        mtime = os.path.getmtime(TICKER_CACHE)
        if time.time() - mtime < 86400:
            with open(TICKER_CACHE, "r", encoding="utf-8") as f:
                return json.load(f)

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
    return mapping


# ==================== 工具函数 ====================
def parse_display_name(name):
    if not name:
        return "", ""
    name = re.sub(r"\s*\(CIK\s+\d+\)\s*$", "", name).strip()
    m = re.search(r"\(([^)]+)\)\s*$", name)
    if m:
        return name[:m.start()].strip(), m.group(1).strip()
    return name.strip(), ""


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


def classify(record):
    query = record.get("query", "") or ""
    items = set(record.get("items") or [])

    if "3.02" in items:
        return "offering"
    if "5.03" in items or "3.03" in items:
        if SPLIT_PAT.search(query):
            return "split"
        if NAME_PAT.search(query):
            return "name_change"

    if OFFERING_PAT.search(query):
        return "offering"
    if SPLIT_PAT.search(query):
        return "split"
    if NAME_PAT.search(query):
        return "name_change"
    return "other"


def score_record(record):
    score = 0
    query = record.get("query", "") or ""
    items = set(record.get("items") or [])
    form = record.get("form_type", "") or ""

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

    if form in ("8-K", "6-K"):
        score += 20
    elif form in ("424B5", "S-1", "S-3"):
        score += 15

    if SPLIT_PAT.search(query) or NAME_PAT.search(query) or OFFERING_PAT.search(query):
        score += 20

    return min(100, score)


def items_to_cn(items_list):
    if not items_list:
        return ""
    if isinstance(items_list, str):
        items_list = re.findall(r"\d+\.\d+", items_list)
    return "；".join(ITEM_CN.get(i, i) for i in items_list)


# ==================== 爬虫 ====================
class SECEdgarScraper:

    def __init__(self, user_agent=SEC_USER_AGENT):
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

    def _get_json(self, url, params):
        self._throttle()
        for attempt in range(4):
            try:
                resp = self.session.get(url, params=params, timeout=20)
                if resp.status_code == 200:
                    try:
                        return resp.json()
                    except ValueError:
                        return None
                if resp.status_code in (403, 429):
                    time.sleep(2 ** attempt)
                    continue
                return None
            except requests.RequestException:
                time.sleep(2 ** attempt)
        return None

    def _parse_hit(self, h, query):
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

    def search(self, query, start_date, end_date, forms=FORMS,
               max_results=MAX_PER_QUERY):
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

    def batch_search(self, queries, days_back, log_fn=None):
        end_dt = datetime.now()
        start_dt = end_dt - timedelta(days=days_back)
        start_date = start_dt.strftime("%Y-%m-%d")
        end_date = end_dt.strftime("%Y-%m-%d")

        all_items, seen = [], set()
        total_q = len(queries)
        for i, q in enumerate(queries, 1):
            if log_fn:
                log_fn("[{}/{}] 检索: {}".format(i, total_q, q))
            try:
                items = self.search(q, start_date, end_date)
            except Exception as e:
                if log_fn:
                    log_fn("  失败: {}".format(e))
                continue
            if log_fn:
                log_fn("  → 命中 {} 条".format(len(items)))
            for it in items:
                key = (it["cik"], it["url"])
                if key in seen:
                    continue
                seen.add(key)
                all_items.append(it)
            time.sleep(0.2)
        return all_items


# ==================== 清洗 ====================
def clean(records, us_map):
    cols = ["日期", "事件类型", "股票代码", "公司名称", "交易所",
            "重要度", "表格类型", "涉及事项", "SEC链接"]

    if not records:
        return pd.DataFrame(columns=cols)

    df = pd.DataFrame(records)

    df["event_type_raw"] = df.apply(classify, axis=1)
    df = df[df["event_type_raw"].isin(["split", "name_change", "offering"])].copy()
    if df.empty:
        return pd.DataFrame(columns=cols)

    df["score"] = df.apply(score_record, axis=1)

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

    df = df.drop_duplicates(subset=["cik", "url"], keep="first")

    df["日期"] = pd.to_datetime(df["file_date"], errors="coerce").dt.strftime("%Y-%m-%d")
    df["事件类型"] = df["event_type_raw"].map(EVENT_CN)
    df["股票代码"] = df["ticker"].fillna("").astype(str)
    df["公司名称"] = df["company"].fillna("").astype(str)
    df["交易所"] = df["exchange"].fillna("").astype(str)
    df["重要度"] = df["score"].apply(lambda s: "⭐" * min(5, s // 20) if s >= 40 else "")
    df["表格类型"] = df["form_type"].fillna("").astype(str)
    df["涉及事项"] = df["items"].apply(items_to_cn)
    df["SEC链接"] = df["url"]

    df = df.sort_values(by=["score", "日期"], ascending=[False, False])
    df = df.drop(columns=["score"], errors="ignore").reset_index(drop=True)

    for c in cols:
        if c not in df.columns:
            df[c] = ""

    return df[cols]


# ==================== GUI ====================
class App:

    def __init__(self, root):
        self.root = root
        self.root.title("美股市场事件扫描工具")
        self.root.geometry("820x620")
        self.root.configure(bg="#f5f6fa")

        self.running = False
        self.us_map = None

        self._build_ui()

    def _build_ui(self):
        # ============ 顶部标题 ============
        header = tk.Frame(self.root, bg="#2c3e50", height=75)
        header.pack(fill="x")
        header.pack_propagate(False)

        tk.Label(
            header, text="美股市场事件扫描",
            font=("微软雅黑", 19, "bold"),
            bg="#2c3e50", fg="white",
        ).pack(side="left", padx=25)
        tk.Label(
            header, text="全市场 · 更名 / 合股 / 增发",
            font=("微软雅黑", 11),
            bg="#2c3e50", fg="#95a5a6",
        ).pack(side="left")

        # ============ 控制区 ============
        ctrl = tk.Frame(self.root, bg="#f5f6fa")
        ctrl.pack(fill="x", padx=25, pady=(20, 10))

        tk.Label(
            ctrl, text="扫描时间范围",
            font=("微软雅黑", 11),
            bg="#f5f6fa", fg="#2c3e50",
        ).grid(row=0, column=0, sticky="w", pady=5)

        self.range_var = tk.StringVar(value="1个月")
        range_combo = ttk.Combobox(
            ctrl, textvariable=self.range_var,
            values=["1天", "1周", "1个月", "3个月", "半年", "1年"],
            width=12, font=("微软雅黑", 11), state="readonly",
        )
        range_combo.grid(row=1, column=0, sticky="w")

        tk.Label(
            ctrl, text="扫描全市场所有公司",
            font=("微软雅黑", 9),
            bg="#f5f6fa", fg="#7f8c8d",
        ).grid(row=2, column=0, sticky="w", pady=(2, 0))

        # 按钮
        btn_frame = tk.Frame(ctrl, bg="#f5f6fa")
        btn_frame.grid(row=1, column=1, sticky="w", padx=(30, 0))

        self.scan_btn = tk.Button(
            btn_frame, text="开始扫描",
            font=("微软雅黑", 11, "bold"),
            bg="#3498db", fg="white",
            activebackground="#2980b9", activeforeground="white",
            relief="flat", padx=30, pady=8, cursor="hand2",
            command=self.on_scan,
        )
        self.scan_btn.pack(side="left")

        tk.Button(
            btn_frame, text="清空日志",
            font=("微软雅黑", 11),
            bg="#ecf0f1", fg="#2c3e50",
            activebackground="#dcdde1",
            relief="flat", padx=20, pady=8, cursor="hand2",
            command=self.on_clear,
        ).pack(side="left", padx=8)

        # ============ 日志区 ============
        log_frame = tk.Frame(self.root, bg="#f5f6fa")
        log_frame.pack(fill="both", expand=True, padx=25, pady=(5, 10))

        tk.Label(
            log_frame, text="扫描日志",
            font=("微软雅黑", 11),
            bg="#f5f6fa", fg="#2c3e50",
        ).pack(anchor="w")

        self.log_text = scrolledtext.ScrolledText(
            log_frame, font=("Consolas", 10),
            bg="white", fg="#2c3e50",
            relief="flat", bd=0,
            wrap="word",
        )
        self.log_text.pack(fill="both", expand=True, pady=(5, 0))
        self.log_text.configure(state="disabled")

        # ============ 状态栏 ============
        self.status_var = tk.StringVar(value="就绪")
        status = tk.Label(
            self.root, textvariable=self.status_var,
            font=("微软雅黑", 9),
            bg="#ecf0f1", fg="#7f8c8d",
            anchor="w", padx=25, pady=5,
        )
        status.pack(fill="x", side="bottom")

    # ============ 日志 ============
    def log(self, msg):
        def _append():
            self.log_text.configure(state="normal")
            self.log_text.insert("end", msg + "\n")
            self.log_text.see("end")
            self.log_text.configure(state="disabled")
        self.root.after(0, _append)

    def set_status(self, msg):
        self.root.after(0, lambda: self.status_var.set(msg))

    # ============ 事件 ============
    def on_clear(self):
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.configure(state="disabled")
        self.status_var.set("就绪")

    def on_scan(self):
        if self.running:
            messagebox.showinfo("提示", "正在扫描中，请稍候...")
            return

        range_map = {
            "1天": 1, "1周": 7, "1个月": 30,
            "3个月": 90, "半年": 180, "1年": 365,
        }
        days = range_map.get(self.range_var.get(), 30)

        self.running = True
        self.scan_btn.configure(state="disabled", bg="#95a5a6")

        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.configure(state="disabled")

        t = threading.Thread(target=self._do_scan, args=(days,), daemon=True)
        t.start()

    def _do_scan(self, days):
        try:
            # 1. 加载美股列表
            if self.us_map is None:
                self.set_status("加载美股列表...")
                self.log("正在加载美股列表（首次约 10~30 秒）...")
                self.us_map = load_us_stock_map()
                self.log("已加载 {} 只股票\n".format(len(self.us_map)))

            # 2. 扫描
            self.set_status("正在扫描...")
            self.log("=" * 60)
            self.log("开始扫描近 {} 天的市场事件".format(days))
            self.log("=" * 60)

            scraper = SECEdgarScraper()
            raw = scraper.batch_search(QUERIES, days, log_fn=self.log)
            self.log("\n原始记录: {} 条".format(len(raw)))

            # 3. 清洗
            self.set_status("清洗数据...")
            df = clean(raw, self.us_map)
            self.log("清洗后: {} 条".format(len(df)))

            if df.empty:
                self.log("\n⚠️ 该时间段内没有找到事件")
                self.set_status("扫描完成：无结果")
                return

            # 4. 保存
            fname = "市场事件_近{}.xlsx".format(self.range_var.get())
            out = os.path.join(BASE_DIR, fname)
            try:
                with pd.ExcelWriter(out, engine="openpyxl") as writer:
                    df.to_excel(writer, sheet_name="全部", index=False)
                    for event_cn in ["合股/拆股", "公司更名", "增发"]:
                        sub = df[df["事件类型"] == event_cn]
                        if not sub.empty:
                            sheet = event_cn.replace("/", "_")
                            sub.to_excel(writer, sheet_name=sheet, index=False)
            except Exception as e:
                logger.warning("Excel 保存失败: %s", e)
                out = os.path.join(BASE_DIR, fname.replace(".xlsx", ".csv"))
                df.to_csv(out, index=False, encoding="utf-8-sig")

            self.log("\n💾 已保存: {}".format(out))

            # 统计
            self.log("\n统计：")
            for et in ["合股/拆股", "公司更名", "增发"]:
                n = len(df[df["事件类型"] == et])
                self.log("  {}: {} 条".format(et, n))
            self.log("  总计: {} 条".format(len(df)))

            self.set_status("扫描完成：{} 条，已保存 Excel".format(len(df)))

            # 5. 自动打开
            try:
                if sys.platform == "win32":
                    os.startfile(out)
                elif sys.platform == "darwin":
                    os.system('open "{}"'.format(out))
                else:
                    os.system('xdg-open "{}"'.format(out))
            except Exception as e:
                logger.warning("自动打开失败: %s", e)

        except Exception as e:
            self.log("❌ 扫描出错: {}".format(e))
            self.set_status("扫描出错")
        finally:
            self.running = False
            self.root.after(0, lambda: self.scan_btn.configure(
                state="normal", bg="#3498db"))


# ==================== 入口 ====================
def main():
    root = tk.Tk()
    app = App(root)
    root.mainloop()


if __name__ == "__main__":
    main()