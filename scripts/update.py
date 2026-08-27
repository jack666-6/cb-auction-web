"""一鍵更新: 抓TWSE競拍公告 -> 補FinMind轉換價/現股價 -> 重新產生 CB競拍統計.xlsx

用法:
    python3 scripts/update.py
"""
import fetch_twse
import fetch_finmind
import build_excel
import export_web_json


def main():
    fetch_twse.main()
    fetch_finmind.main()
    build_excel.main()
    export_web_json.main()


if __name__ == "__main__":
    main()
