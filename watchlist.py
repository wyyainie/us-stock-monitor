# -*- coding: utf-8 -*-
"""
诊断脚本：查指定 ticker 在 SEC 的所有近期文件
用法: python diagnose.py HKIT
"""

import sys
import json
import requests

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

SEC_USER_AGENT = "StockEventBot/1.0 (your_email@example.com)"  # 改成你的邮箱
TICKER_CACHE = os.path.join(BASE_DIR, "sec_tickers_cache.json")



def diagnose(ticker: str):
    ticker = ticker.upper()

    # 1. 从缓存找 CIK
    with open(TICKER_CACHE, "r", encoding="utf-8") as f:
        us_map = json.load(f)

    cik = None
    for k, v in us_map.items():
        if v.get("ticker", "").upper() == ticker:
            cik = k
            break

    if not cik:
        print(f"❌ 在股票列表里没找到 {ticker}")
        return

    print(f"✅ {ticker} 的 CIK: {cik}")

    # 2. 查 submissions
    url = f"https://data.sec.gov/submissions/CIK{cik}.json"
    r = requests.get(url, headers={"User-Agent": SEC_USER_AGENT}, timeout=30)
    if r.status_code != 200:
        print(f"❌ 请求失败: {r.status_code}")
        return

    data = r.json()
    print(f"公司名: {data.get('name')}")
    print(f"注册地: {data.get('stateOfIncorporation')}")
    print(f"财年: {data.get('fiscalYearEnd')}")

    filings = data.get("filings", {}).get("recent", {})
    forms = filings.get("form", [])
    dates = filings.get("filingDate", [])
    accessions = filings.get("accessionNumber", [])
    primary_docs = filings.get("primaryDocument", [])
    descriptions = filings.get("primaryDocDescription", [])

    print(f"\n近期所有 8-K / 6-K / S 系列 / 424B 文件（最近 60 条）：")
    print("-" * 120)
    cik_int = str(int(cik))
    count = 0
    for i, form in enumerate(forms):
        if count >= 60:
            break
        if form in ("8-K", "8-K/A", "6-K", "6-K/A",
                    "S-1", "S-1/A", "S-3", "S-3/A",
                    "424B2", "424B3", "424B5"):
            date = dates[i]
            acc = accessions[i].replace("-", "")
            doc = primary_docs[i]
            desc = descriptions[i] if i < len(descriptions) else ""
            link = f"https://www.sec.gov/Archives/edgar/data/{cik_int}/{acc}/{doc}"
            print(f"{date}  {form:<8}  {desc[:45]:<45}  {link}")
            count += 1

    if count == 0:
        print("（无）")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("用法: python diagnose.py HKIT")
        sys.exit(1)
    diagnose(sys.argv[1])