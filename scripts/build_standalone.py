"""社内共有（LINE WORKS 等）用に、ダッシュボードを1つのHTMLファイルにまとめる。

公開版の index.html は CSS・グラフ部品（Chart.js）・データ（data/*.json）を別ファイルから読み込むため、
index.html だけを配っても表示されない（PCで直接開くとブラウザがデータ読込を止める）。
このスクリプトはそれらをすべて1ファイルに埋め込み、ダウンロードしてブラウザで開くだけで
公開版と同じ操作（データセット・年度・期間の切替、会社の絞り込み等）ができるファイルを作る。
インターネット接続は不要（各社IR資料へのリンクを開くときだけ必要）。

  埋め込むもの
    assets/tailwind.css                       … <style> に展開
    assets/vendor/chart.umd.min.js            … Chart.js 4.4.1（公開版と同じ版。SRIで照合済み）
    assets/vendor/chartjs-plugin-datalabels.min.js … datalabels 2.2.0（同上）
    data/*.json                               … window.__EMBEDDED_DATA__ に格納し、fetch("data/…") をここから返す

usage: .venv/Scripts/python.exe scripts/build_standalone.py [出力先.html]
  既定の出力先: <プロジェクト>/配布用_LINEWORKS/家電量販7社_業績ダッシュボード.html
  （毎回同じファイル名で上書き。LINE WORKS でも同じ名前で差し替える運用を想定）
"""
import base64
import hashlib
import json
import re
import sys
from datetime import date
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
PROJECT = REPO.parent.parent
DEFAULT_OUT = PROJECT / "配布用_LINEWORKS" / "家電量販7社_業績ダッシュボード.html"

# 公開版 index.html の <script integrity> と同じ値。同梱ファイルが差し替わっていないことを毎回確かめる
VENDOR = [
    ("assets/vendor/chart.umd.min.js",
     "https://cdn.jsdelivr.net/npm/chart.js@4.4.1/dist/chart.umd.min.js"),
    ("assets/vendor/chartjs-plugin-datalabels.min.js",
     "https://cdn.jsdelivr.net/npm/chartjs-plugin-datalabels@2.2.0"),
]

# fetch("data/xxx.json?v=…") を埋め込みデータから返す。データに無いファイルは 404 扱い（公開版と同じ挙動）
FETCH_SHIM = r"""
(function () {
  const D = window.__EMBEDDED_DATA__ || {};
  const realFetch = window.fetch ? window.fetch.bind(window) : null;
  window.fetch = function (input, init) {
    const url = String(typeof input === "string" ? input : input.url).split("?")[0].replace(/^\.\//, "");
    if (url.startsWith("data/")) {
      if (Object.prototype.hasOwnProperty.call(D, url)) {
        return Promise.resolve(new Response(JSON.stringify(D[url]),
          { status: 200, headers: { "Content-Type": "application/json" } }));
      }
      return Promise.resolve(new Response("", { status: 404 }));
    }
    return realFetch ? realFetch(input, init) : Promise.reject(new Error("fetch unavailable"));
  };
})();
"""


def sri(path):
    return "sha384-" + base64.b64encode(hashlib.sha384(path.read_bytes()).digest()).decode()


def script_safe(text):
    """<script> 内に埋め込んでも途中で閉じられないようにする。"""
    return text.replace("</", "<\\/")


def replace_once(html, old, new, what):
    n = html.count(old)
    if n != 1:
        sys.exit(f"ERROR: index.html の {what} が見つかりません（{n}件）。index.html の構造が変わった可能性があります")
    return html.replace(old, new)


def main():
    out = Path(sys.argv[1]) if len(sys.argv) > 1 else DEFAULT_OUT
    html = (REPO / "index.html").read_text(encoding="utf-8")

    # ---- CSS ----
    css = (REPO / "assets" / "tailwind.css").read_text(encoding="utf-8")
    html, n = re.subn(r'<link rel="stylesheet" href="assets/tailwind\.css[^"]*">',
                      lambda _: f"<style>\n{css}\n</style>", html)
    if n != 1:
        sys.exit("ERROR: tailwind.css の読込タグが見つかりません")

    # ---- Chart.js（SRIを公開版の値と照合してから埋め込む）----
    for rel, cdn in VENDOR:
        path = REPO / rel
        m = re.search(r'<script src="' + re.escape(cdn) + r'"\s+integrity="([^"]+)"[^>]*></script>', html)
        if not m:
            sys.exit(f"ERROR: {cdn} の読込タグが index.html に見つかりません")
        if sri(path) != m.group(1):
            sys.exit(f"ERROR: {rel} が公開版の SRI と一致しません（版の不一致または改変）")
        js = path.read_text(encoding="utf-8")
        html = html.replace(m.group(0), f"<script>/* {cdn} */\n{script_safe(js)}\n</script>")

    # ---- データ ----
    data = {}
    for f in sorted((REPO / "data").glob("*.json")):
        data[f"data/{f.name}"] = json.loads(f.read_text(encoding="utf-8"))
    payload = json.dumps(data, ensure_ascii=False, separators=(",", ":"))
    head_inject = (f"<script>window.__EMBEDDED_DATA__ = {script_safe(payload)};</script>\n"
                   f"<script>{FETCH_SHIM}</script>\n")
    html = replace_once(html, "</head>", head_inject + "</head>", "</head>")

    # ---- 配布版の表示 ----
    built = date.today().isoformat()
    html = replace_once(
        html,
        '<a href="data/companies.json" class="text-sky-700 underline">data/companies.json</a>',
        f'<span class="font-mono">社内配布版（{built} 作成）</span>',
        "ヘッダーのデータリンク")
    html = replace_once(html, "<title>家電量販店7社 業績ダッシュボード</title>",
                        f"<title>家電量販店7社 業績ダッシュボード（社内配布版 {built}）</title>", "<title>")

    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(html, encoding="utf-8")
    print(f"wrote {out}")
    print(f"  size: {out.stat().st_size / 1024 / 1024:.2f} MB / data files: {len(data)} / built: {built}")


if __name__ == "__main__":
    main()
