"""把 cache/web_data.json 套進 web_template.html，寫到 repo 根目錄的 index.html，
並 commit + push（Render 接了這個repo，push後會自動重新部署）。

跟本地端的 build_web_deploy.py 不同的地方：這支腳本假設自己就跑在這個repo(web-deploy)
裡面，BASE_DIR 就是 repo 根目錄本身，不是另外指到一個 web-deploy 子資料夾。
這樣同一支腳本本地、雲端都能用：本地端從主專案clone一份到 web-deploy/ 執行；
雲端排程則是每次都重新 git clone 這個repo到 /tmp 執行。
"""
import json
import subprocess
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent.parent
TEMPLATE = BASE_DIR / "scripts" / "web_template.html"
DATA_JSON = BASE_DIR / "cache" / "web_data.json"
INDEX_HTML = BASE_DIR / "index.html"


def build_html() -> str:
    template = TEMPLATE.read_text(encoding="utf-8")
    data = json.loads(DATA_JSON.read_text(encoding="utf-8"))
    data_json = json.dumps(data, ensure_ascii=False)
    if "__DATA_JSON__" not in template:
        raise ValueError("模板裡找不到 __DATA_JSON__ 佔位符")
    return template.replace("__DATA_JSON__", data_json)


def git(*args):
    return subprocess.run(
        ["git", *args], cwd=BASE_DIR, capture_output=True, text=True
    )


def main():
    html = build_html()
    INDEX_HTML.write_text(html, encoding="utf-8")

    git("config", "user.email", "diego.gogogo8@gmail.com")
    git("config", "user.name", "jack666-6")
    git("add", "index.html")

    diff = git("diff", "--cached", "--quiet")
    if diff.returncode == 0:
        print("網頁內容沒有變化，不需要推送")
        return

    commit = git("commit", "-m", "自動更新CB競拍網頁")
    if commit.returncode != 0:
        print("commit失敗：", commit.stdout, commit.stderr)
        return

    push = git("push", "origin", "main")
    if push.returncode != 0:
        print("push失敗：", push.stdout, push.stderr)
        return

    print("已推送到GitHub，Render會自動重新部署 (https://cb-auction-web.onrender.com)")


if __name__ == "__main__":
    main()
