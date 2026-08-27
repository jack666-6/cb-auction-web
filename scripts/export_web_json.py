"""把 build_excel.py 算好的資料轉成網頁用的 {headers, rows} JSON，不用真的產生xlsx。"""
import json
from datetime import date
from pathlib import Path

import build_excel

BASE_DIR = Path(__file__).resolve().parent.parent
OUTPUT_JSON = BASE_DIR / "cache" / "web_data.json"

# 這些欄位存的是「小數形式比率」(例如0.1857代表18.57%)，網頁端會再乘以100顯示成%。
# 如果跟其他價格欄位一樣統一四捨五入到小數點後2位，等於把精度砍到只剩整數百分比
# (0.1857 -> 0.19 -> 顯示19.00%，而不是18.57%)，所以這些欄位要保留更多小數位。
FRACTION_PCT_COLUMNS = {
    "最低得標成本溢價率", "得標加權平均成本溢價率",
    "投標結束日至掛牌第一天現股漲跌幅", "投標結束日至掛牌第六天現股漲跌幅",
    "掛牌第一天開盤價最低得標成本報酬率", "掛牌第一天開盤價得標加權平均成本報酬率",
    "掛牌第一天收盤價最低得標成本報酬率", "掛牌第一天收盤價得標加權平均成本報酬率",
    "掛牌第六天最低得標成本報酬率", "掛牌第六天得標加權平均成本報酬率",
}


def main():
    rows = build_excel.load_rows()
    out_rows = []
    for r in rows:
        row = []
        for h in build_excel.HEADERS:
            v = r.get(h)
            if isinstance(v, date):
                v = v.strftime("%Y/%m/%d")
            elif isinstance(v, float):
                v = round(v, 6) if h in FRACTION_PCT_COLUMNS else round(v, 2)
            row.append(v)
        out_rows.append(row)
    OUTPUT_JSON.write_text(
        json.dumps({"headers": build_excel.HEADERS, "rows": out_rows}, ensure_ascii=False),
        encoding="utf-8",
    )
    print(f"已寫出網頁用資料 {OUTPUT_JSON}，共 {len(out_rows)} 筆")


if __name__ == "__main__":
    main()
