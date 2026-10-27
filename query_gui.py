# -*- coding: utf-8 -*-
"""
美股事件查询工具（图形界面版）
- 输入股票代码，一键查询
- 可调查询天数
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
TICKER_CACHE = os.path.join(BASE_DIR, "sec_tickers_cache.json")
SEC_TICKERS_URL = "https://www.sec.gov/files/company_tickers_exchange.json"

FORMS_OF_INTEREST = {"8-K", "8-K/A", "6-K", "6-K/A",
                     "424B2", "424B3", "424B5",
                     "S-1", "S-1/A", "S-3", "S-3/A"}

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


# ==================== 数据层 ====================
def load_ticker_map(force_refresh=False):
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


def find_cik(ticker, ticker_map):
    ticker = ticker.upper().strip()
    for cik, info in ticker_map.items():
        if info.get("ticker", "").upper() == ticker:
            return cik, info
    return None, None


def get_submissions(cik):
    url = "https://data.sec.gov/submissions/CIK{}.json".format(cik)
    r = requests.get(url, headers={"User-Agent": SEC_USER_AGENT}, timeout=30)
    r.raise_for_status()
    return r.json()


def check_document(url):
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


def query_one(ticker, ticker_map, days_back, log_fn):
    cik, info = find_cik(ticker, ticker_map)
    if not cik:
        log_fn("  ❌ 未找到股票代码: {}".format(ticker))
        return []

    log_fn("  📌 {} (CIK {})".format(ticker, cik))
    log_fn("     公司名: {}".format(info.get("name")))
    log_fn("     交易所: {}".format(info.get("exchange") or "（未登记）"))
    log_fn("     检查近 {} 天的文件...".format(days_back))

    try:
        data = get_submissions(cik)
    except Exception as e:
        log_fn("  ❌ 拉取文件列表失败: {}".format(e))
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

        events = []
        if items:
            if "5.03" in items:
                events.append("公司更名")
            if "3.03" in items:
                events.append("合股/拆股")
            if "3.02" in items:
                events.append("增发")

        detail = ""
        if not events:
            time.sleep(0.3)
            events, detail = check_document(url)

        if events:
            log_fn("  ✅ {} | {:<6} | {:<12} | {}".format(
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
        log_fn("  ⚠️ 近 {} 天没有找到相关事件".format(days_back))
    else:
        log_fn("  → 共命中 {} 条".format(len(results)))
    return results


# ==================== GUI ====================
class App:

    def __init__(self, root):
        self.root = root
        self.root.title("美股事件查询工具")
        self.root.geometry("780x580")
        self.root.configure(bg="#f5f6fa")

        self.ticker_map = None
        self.querying = False

        self._build_ui()

    def _build_ui(self):
        # ============ 顶部标题 ============
        header = tk.Frame(self.root, bg="#2c3e50", height=70)
        header.pack(fill="x")
        header.pack_propagate(False)

        tk.Label(
            header, text="美股事件查询工具",
            font=("微软雅黑", 18, "bold"),
            bg="#2c3e50", fg="white",
        ).pack(side="left", padx=25)
        tk.Label(
            header, text="更名 · 合股 · 增发",
            font=("微软雅黑", 11),
            bg="#2c3e50", fg="#95a5a6",
        ).pack(side="left")

        # ============ 输入区 ============
        input_frame = tk.Frame(self.root, bg="#f5f6fa")
        input_frame.pack(fill="x", padx=25, pady=(20, 10))

        # 股票代码
        tk.Label(
            input_frame, text="股票代码",
            font=("微软雅黑", 11),
            bg="#f5f6fa", fg="#2c3e50",
        ).grid(row=0, column=0, sticky="w", pady=5)

        self.ticker_entry = tk.Entry(
            input_frame, font=("微软雅黑", 12),
            width=30, relief="flat", bd=1,
            highlightthickness=1, highlightbackground="#dcdde1",
        )
        self.ticker_entry.grid(row=1, column=0, sticky="w", padx=(0, 20))
        self.ticker_entry.insert(0, "")
        self.ticker_entry.bind("<Return>", lambda e: self.on_query())

        tk.Label(
            input_frame, text="多个代码用逗号分隔，如 BDAI,SBFM",
            font=("微软雅黑", 9),
            bg="#f5f6fa", fg="#7f8c8d",
        ).grid(row=2, column=0, sticky="w", pady=(2, 0))

        # 天数
        tk.Label(
            input_frame, text="查询天数",
            font=("微软雅黑", 11),
            bg="#f5f6fa", fg="#2c3e50",
        ).grid(row=0, column=1, sticky="w", pady=5)

        self.days_var = tk.StringVar(value="365")
        days_combo = ttk.Combobox(
            input_frame, textvariable=self.days_var,
            values=["30", "90", "180", "365", "730", "1095"],
            width=10, font=("微软雅黑", 11), state="readonly",
        )
        days_combo.grid(row=1, column=1, sticky="w")

        tk.Label(
            input_frame, text="近 N 天内的文件",
            font=("微软雅黑", 9),
            bg="#f5f6fa", fg="#7f8c8d",
        ).grid(row=2, column=1, sticky="w", pady=(2, 0))

        # 按钮
        btn_frame = tk.Frame(input_frame, bg="#f5f6fa")
        btn_frame.grid(row=1, column=2, sticky="w", padx=(30, 0))

        self.query_btn = tk.Button(
            btn_frame, text="开始查询",
            font=("微软雅黑", 11, "bold"),
            bg="#3498db", fg="white",
            activebackground="#2980b9", activeforeground="white",
            relief="flat", padx=25, pady=8, cursor="hand2",
            command=self.on_query,
        )
        self.query_btn.pack(side="left")

        tk.Button(
            btn_frame, text="清空",
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
            log_frame, text="查询日志",
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

    # ============ 日志输出 ============
    def log(self, msg):
        def _append():
            self.log_text.configure(state="normal")
            self.log_text.insert("end", msg + "\n")
            self.log_text.see("end")
            self.log_text.configure(state="disabled")
        self.root.after(0, _append)

    def set_status(self, msg):
        self.root.after(0, lambda: self.status_var.set(msg))

    # ============ 按钮事件 ============
    def on_clear(self):
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.configure(state="disabled")
        self.ticker_entry.delete(0, "end")
        self.status_var.set("就绪")

    def on_query(self):
        if self.querying:
            messagebox.showinfo("提示", "正在查询中，请稍候...")
            return

        raw = self.ticker_entry.get().strip()
        if not raw:
            messagebox.showwarning("提示", "请输入股票代码")
            return

        tickers = [t.strip().upper() for t in re.split(r"[,;\s]+", raw) if t.strip()]
        if not tickers:
            messagebox.showwarning("提示", "请输入有效的股票代码")
            return

        try:
            days_back = int(self.days_var.get())
        except Exception:
            days_back = 365

        self.querying = True
        self.query_btn.configure(state="disabled", bg="#95a5a6")

        # 清空日志
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.configure(state="disabled")

        # 后台线程
        t = threading.Thread(
            target=self._do_query, args=(tickers, days_back), daemon=True
        )
        t.start()

    def _do_query(self, tickers, days_back):
        try:
            # 加载股票列表
            if self.ticker_map is None:
                self.set_status("正在加载股票列表...")
                self.log("正在加载股票列表...")
                self.ticker_map = load_ticker_map()
                self.log("已加载 {} 只股票\n".format(len(self.ticker_map)))

            all_results = []
            for tk_code in tickers:
                self.set_status("正在查询 {}...".format(tk_code))
                self.log("=" * 55)
                results = query_one(tk_code, self.ticker_map, days_back, self.log)
                all_results.extend(results)
                self.log("")

            if not all_results:
                self.log("⚠️ 本次查询无结果")
                self.set_status("查询完成：无结果")
                return

            df = pd.DataFrame(all_results)
            df = df.sort_values("日期", ascending=False).reset_index(drop=True)

            # 保存
            if len(tickers) == 1:
                fname = "查询结果_{}.xlsx".format(tickers[0])
            else:
                fname = "查询结果_{}.xlsx".format("_".join(tickers[:3]))

            out = os.path.join(BASE_DIR, fname)
            try:
                df.to_excel(out, index=False, engine="openpyxl")
            except Exception as e:
                logger.warning("Excel 保存失败: %s", e)
                out = os.path.join(BASE_DIR, fname.replace(".xlsx", ".csv"))
                df.to_csv(out, index=False, encoding="utf-8-sig")

            self.log("💾 已保存: {}".format(out))
            self.set_status("查询完成：{} 条，已保存 Excel".format(len(df)))

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

        except Exception as e:
            self.log("❌ 查询出错: {}".format(e))
            self.set_status("查询出错")
        finally:
            self.querying = False
            self.root.after(0, lambda: self.query_btn.configure(
                state="normal", bg="#3498db"))


# ==================== 入口 ====================
def main():
    root = tk.Tk()
    app = App(root)
    root.mainloop()


if __name__ == "__main__":
    main()