"""為每檔CB補上「轉換價」「投標結束日現股收盤價」以及「掛牌後前六個交易日」的股價/CB市價。

- 轉換價 (ConversionPrice):
  1. 優先用 FinMind TaiwanStockConvertibleBondDailyOverview，取該CB最早一筆的轉換價
     (轉換價在掛牌前即已依發行條件訂定，掛牌後除非除權息調整否則不變，用最早可查到的一筆估算拍賣當時的轉換價)
  2. 若CB尚未掛牌、FinMind還查不到，改用統一證券CBAS資訊網「預計發行CB資料/最近掛牌」
     (https://cbas16889.pscnet.com.tw/marketInfo/expectedRelease/) 的 conversion_price 當備援
- 投標結束日現股收盤價: FinMind TaiwanStockPrice，母公司股票代號=CB代號前4碼
- 投標結束日前75日股價波動率: 投標結束日往前75個交易日(含當天)母公司現股收盤價的變異係數
  (標準差 ÷ 平均值)，不是用報酬率算的年化波動率
- 掛牌後現股/CB價格: 以「撥券日期(上市、上櫃日期)」為掛牌第一天，往後抓母公司現股(TaiwanStockPrice)
  與CB自己的市場價格(TaiwanStockConvertibleBondDaily)各6個交易日；此欄位公式皆已用原始
  「CB資料表 的副本.xlsx」逐筆反推驗證過(見開發過程)，只有CB還沒掛牌時才會是空的。

有快取機制(cache/finmind_cb.json)，已完整解析的CB不會重複打API。
"""
import json
import time
from datetime import date, datetime, timedelta
from pathlib import Path

import requests

BASE_DIR = Path(__file__).resolve().parent.parent
CACHE_DIR = BASE_DIR / "cache"
RAW_CACHE = CACHE_DIR / "twse_auction_raw.json"
FINMIND_CACHE = CACHE_DIR / "finmind_cb.json"
CBAS_CACHE = CACHE_DIR / "cbas_recently_listed.json"

TOKEN = (
    "eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzI1NiJ9."
    "eyJ1c2VyX2lkIjoiU2FtdWVsLUZpbm1pbmQiLCJlbWFpbCI6Imt1bnNlbi5saWFvQGdtYWlsLmNvbSIsInRva2VuX3ZlcnNpb24iOjF9."
    "0M2SsYqe4BkjT7540-rSX7rMLTFjiLdoZHplblQ8h9g"
)
API_URL = "https://api.finmindtrade.com/api/v4/data"
CBAS_URL = "https://cbas16889.pscnet.com.tw/api/CbasQuote/GetRecentlyListed"


def fetch_cbas_conversion_prices() -> dict:
    """抓統一證券CBAS「最近掛牌」清單，回傳 {cb_code: conversion_price}

    該站憑證鏈缺少 Subject Key Identifier，Python的嚴格SSL驗證會擋下(curl不會)，
    這裡只針對這個已知有此憑證問題的公開唯讀資料網域關閉憑證驗證。
    """
    try:
        import urllib3

        urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
        resp = requests.get(CBAS_URL, timeout=20, verify=False)
        resp.raise_for_status()
        payload = resp.json()
    except (requests.RequestException, ValueError):
        return {}
    out = {}
    for item in payload.get("result", []):
        cb_code = item.get("cb_code")
        price = item.get("conversion_price")
        if cb_code and price not in (None, ""):
            try:
                out[cb_code] = float(price)
            except ValueError:
                pass
    CBAS_CACHE.write_text(json.dumps(out, ensure_ascii=False, indent=2), encoding="utf-8")
    return out


def _get(dataset: str, data_id: str, start_date: str, end_date: str):
    for attempt in range(3):
        resp = requests.get(
            API_URL,
            params={
                "dataset": dataset,
                "data_id": data_id,
                "start_date": start_date,
                "end_date": end_date,
                "token": TOKEN,
            },
            timeout=20,
        )
        if resp.status_code == 200:
            payload = resp.json()
            if payload.get("msg") == "success":
                return payload.get("data", [])
            return []
        if resp.status_code in (429, 402):
            time.sleep(5 * (attempt + 1))
            continue
        return []
    return []


def load_cache() -> dict:
    if FINMIND_CACHE.exists():
        return json.loads(FINMIND_CACHE.read_text(encoding="utf-8"))
    return {}


def save_cache(cache: dict):
    FINMIND_CACHE.write_text(
        json.dumps(cache, ensure_ascii=False, indent=2), encoding="utf-8"
    )


def _price_volatility(closes: list[float]) -> float | None:
    """變異係數 = 標準差(樣本) ÷ 平均值，至少要有2筆資料才算得出來"""
    if len(closes) < 2:
        return None
    mean = sum(closes) / len(closes)
    if not mean:
        return None
    variance = sum((c - mean) ** 2 for c in closes) / (len(closes) - 1)
    return (variance ** 0.5) / mean


def fetch_one(cb_code: str, bid_end_date: str, listing_date: str, cbas_prices: dict) -> dict:
    """回傳轉換價、投標結束日現股收盤價、投標結束日前75日股價波動率，
    以及掛牌後6個交易日的現股/CB價格序列"""
    result = {
        "轉換價": None,
        "投標結束日現股收盤價": None,
        "投標結束日前75日股價波動率": None,
        "掛牌現股序列": None,  # [[date, open, close], ...] 最多6筆
        "掛牌CB序列": None,  # [[date, open, close, max], ...] 最多6筆
    }

    overview = _get(
        "TaiwanStockConvertibleBondDailyOverview",
        cb_code,
        "2000-01-01",
        "2030-12-31",
    )
    if overview:
        result["轉換價"] = overview[0].get("ConversionPrice")
    elif cb_code in cbas_prices:
        result["轉換價"] = cbas_prices[cb_code]

    parent_code = cb_code[:4]
    if bid_end_date:
        prices = _get("TaiwanStockPrice", parent_code, bid_end_date, bid_end_date)
        if prices:
            result["投標結束日現股收盤價"] = prices[0].get("close")

        vol_end = datetime.strptime(bid_end_date, "%Y-%m-%d").date()
        vol_start = vol_end - timedelta(days=150)  # 150個日曆天足夠涵蓋75個交易日(含假日緩衝)
        vol_rows = _get("TaiwanStockPrice", parent_code, vol_start.isoformat(), vol_end.isoformat())
        vol_closes = [r["close"] for r in vol_rows if r.get("close")][-75:]
        result["投標結束日前75日股價波動率"] = _price_volatility(vol_closes)

    if listing_date:
        start = datetime.strptime(listing_date, "%Y-%m-%d").date()
        end = start + timedelta(days=20)  # 20個日曆天足夠涵蓋6個交易日(含假日)
        if start <= date.today():  # 掛牌日還沒到就一定沒資料，不用浪費API呼叫；掛牌當天起就先抓，能拿到幾天算幾天
            stock_rows = _get("TaiwanStockPrice", parent_code, start.isoformat(), end.isoformat())
            stock_rows = [r for r in stock_rows if r.get("close")]
            result["掛牌現股序列"] = [
                [r["date"], r.get("open"), r.get("close")] for r in stock_rows[:6]
            ] or None

            cb_rows = _get("TaiwanStockConvertibleBondDaily", cb_code, start.isoformat(), end.isoformat())
            # FinMind偶爾會回傳收盤價全為0的假交易日(遇到連假等情況)，這種不是真的交易日要濾掉，
            # 不然「第六個交易日」會算錯(用現股序列已排除的方式驗證過)
            cb_rows = [r for r in cb_rows if r.get("close")]
            result["掛牌CB序列"] = [
                [r["date"], r.get("open"), r.get("close"), r.get("max")] for r in cb_rows[:6]
            ] or None

    return result


def _is_resolved(entry: dict) -> bool:
    if entry.get("轉換價") is None or entry.get("投標結束日現股收盤價") is None:
        return False
    # 舊快取沒有這個欄位(新增功能前抓的)，強制重抓一次補上；沒抓到值(母公司歷史資料不足75天)
    # 就算resolved，不用每次重試
    if "投標結束日前75日股價波動率" not in entry:
        return False
    # 掛牌現股/CB序列要滿6筆才算真的解決；還沒掛牌滿20天前，每次都會重試(fetch_one內部會自動跳過還沒到的)
    for key in ("掛牌現股序列", "掛牌CB序列"):
        seq = entry.get(key)
        if seq is None or len(seq) < 6:
            return False
    return True


def main():
    if not RAW_CACHE.exists():
        print("找不到 twse_auction_raw.json，請先執行 fetch_twse.py")
        return

    rows = json.loads(RAW_CACHE.read_text(encoding="utf-8"))
    cache = load_cache()
    cbas_prices = fetch_cbas_conversion_prices()
    print(f"統一證券CBAS「最近掛牌」轉換價備援資料: {len(cbas_prices)} 檔")

    updated = 0
    for row in rows:
        cb_code = str(row["證券代號"])
        bid_end = row["投標結束日"]
        listing_date = row["撥券日期(上市、上櫃日期)"]
        if cb_code in cache and _is_resolved(cache[cb_code]):
            continue
        data = fetch_one(cb_code, bid_end, listing_date, cbas_prices)
        cache[cb_code] = data
        updated += 1
        if updated % 20 == 0:
            save_cache(cache)
            print(f"...已處理 {updated} 檔")
        time.sleep(0.3)

    save_cache(cache)
    print(f"完成，本次新抓/補齊 {updated} 檔，快取共 {len(cache)} 檔，存於 {FINMIND_CACHE}")


if __name__ == "__main__":
    main()
