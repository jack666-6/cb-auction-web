"""抓取證交所競價拍賣公告(auction.html)資料，篩選出可轉換公司債(CB)並存成本地快取。

資料來源: https://www.twse.com.tw/rwd/zh/announcement/auction?date=YYYY0101&response=json
證交所僅提供民國105年(西元2016年)起的資料。
"""
import json
import re
from datetime import datetime, date
from pathlib import Path

import requests

BASE_DIR = Path(__file__).resolve().parent.parent
CACHE_DIR = BASE_DIR / "cache"
CACHE_DIR.mkdir(exist_ok=True)
RAW_CACHE = CACHE_DIR / "twse_auction_raw.json"

START_YEAR = 2026  # 使用者只需要「今年」開始競拍的CB，不用回溯歷史年度
API_URL = "https://www.twse.com.tw/rwd/zh/announcement/auction"


def _to_number(s):
    s = (s or "").replace(",", "").strip()
    if s == "":
        return None
    try:
        if re.fullmatch(r"-?\d+", s):
            return int(s)
        return float(s)
    except ValueError:
        return None


def _to_date(s):
    s = (s or "").strip()
    if not s:
        return None
    return datetime.strptime(s, "%Y/%m/%d").date()


def fetch_year(year: int) -> list[dict]:
    resp = requests.get(
        API_URL,
        params={"date": f"{year}0101", "response": "json"},
        timeout=20,
    )
    resp.raise_for_status()
    payload = resp.json()
    if payload.get("stat") != "OK":
        return []
    rows = []
    for r in payload.get("data", []):
        rows.append(
            {
                "開標日期": _to_date(r[1]),
                "證券名稱": r[2],
                "證券代號": int(r[3]) if r[3].isdigit() else r[3],
                "發行市場": r[4],
                "發行性質": r[5],
                "競拍方式": r[6],
                "投標開始日": _to_date(r[7]),
                "投標結束日": _to_date(r[8]),
                "競拍數量(張)": _to_number(r[9]),
                "最低投標價格(元)": _to_number(r[10]),
                "最低每標單投標數量(張)": _to_number(r[11]),
                "最高投(得)標數量(張)": _to_number(r[12]),
                "保證金成數(%)": _to_number(r[13]),
                "每一投標單投標處理費(元)": _to_number(r[14]),
                "撥券日期(上市、上櫃日期)": _to_date(r[15]),
                "主辦券商": r[16],
                "得標總金額(元)": _to_number(r[17]),
                "得標手續費率(%)": _to_number(r[18]),
                "總合格件": _to_number(r[19]),
                "合格投標數量(張)": _to_number(r[20]),
                "最低得標價格(元)": _to_number(r[21]),
                "最高得標價格(元)": _to_number(r[22]),
                "得標加權平均價格(元)": _to_number(r[23]),
                "承銷價格(元)": _to_number(r[24]),
                "取消競價拍賣(流標或取消)": r[25],
            }
        )
    return rows


def fetch_all(end_year: int | None = None) -> list[dict]:
    end_year = end_year or date.today().year
    all_rows = []
    for year in range(START_YEAR, end_year + 1):
        all_rows.extend(fetch_year(year))
    return all_rows


def is_cb(row: dict) -> bool:
    return "轉換公司債" in (row.get("發行性質") or "")


def main():
    rows = fetch_all()
    cb_rows = [r for r in rows if is_cb(r)]

    def _default(o):
        if isinstance(o, date):
            return o.isoformat()
        raise TypeError

    RAW_CACHE.write_text(
        json.dumps(cb_rows, ensure_ascii=False, indent=2, default=_default),
        encoding="utf-8",
    )
    print(f"抓到 {len(rows)} 筆競拍公告，其中 CB {len(cb_rows)} 筆，已存入 {RAW_CACHE}")


if __name__ == "__main__":
    main()
