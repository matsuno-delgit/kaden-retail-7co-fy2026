# Tailwind CSS の再生成手順

`assets/tailwind.css` は `index.html` で使用しているクラスのみを含む **静的・最小**のTailwind CSSです。
新しいクラスを `index.html` に追加した場合は、以下の手順で再生成してください。

## なぜ自己ホスト？

セキュリティ審査でCDN供給網リスクを抑える目的（SRIが効かない動的CSS生成への依存をなくす）。
詳細は `SECURITY.md` 参照。

---

## 方法 1: Tailwind Play CDN から抽出（推奨・ツール不要）

1. ローカルでHTTPサーバを立ち上げる
   ```bash
   cd electronics-retail-dashboard
   python -m http.server 8000
   ```

2. ブラウザで http://localhost:8000/ を開く
   （`index.html` の `<link rel="stylesheet" href="assets/tailwind.css">` を一時的にコメントアウトし、
   `<script src="https://cdn.tailwindcss.com"></script>` を一時的に挿入しておく）

3. ブラウザのDevTools (F12) → コンソールで以下を実行：
   ```js
   document.querySelectorAll('style')[1].textContent
   ```
   （Tailwind Play CDN が生成した CSS が表示される。`style[0]` はカスタム`<style>`、`style[1]`が生成CSS）

4. 出力を `assets/tailwind.css` に上書き保存。先頭にこの自動生成ヘッダコメントを残しておくこと

5. `index.html` を元に戻す（`<link>` を有効化、Play CDN `<script>` を削除）

## 方法 2: Tailwind standalone CLI（推奨だがツールDLが必要）

1. Tailwind 公式のスタンドアロンCLI（Windows x64）をダウンロード（約50MB）:
   https://github.com/tailwindlabs/tailwindcss/releases/latest

2. 一時ディレクトリに配置して実行：
   ```bash
   ./tailwindcss-windows-x64.exe \
     -i scripts/tailwind-input.css \
     -o assets/tailwind.css \
     --content "index.html" \
     --minify
   ```

   `tailwind-input.css` の最小例：
   ```css
   @tailwind base;
   @tailwind components;
   @tailwind utilities;
   ```

3. 出力を確認してコミット。CLI バイナリはコミットしないこと（重い & 不要）。

## 方法 3: Node.js + npx（クリーンビルド派）

```bash
# 1度だけ
npm init -y
npm install -D tailwindcss

# 毎回
npx tailwindcss -i scripts/tailwind-input.css -o assets/tailwind.css --content "index.html" --minify
```

`package.json` `package-lock.json` `node_modules` を含めると依存が増えるので、CIで生成する運用が良い。

---

## 検証チェックリスト

再生成後、以下を確認：

- [ ] `index.html` を `python -m http.server` 経由で開いてスタイル崩れがない
- [ ] DevToolsで未定義クラス警告が出ていない
- [ ] `git diff assets/tailwind.css` で予期した差分のみ
- [ ] サイズが大きく増えていない（数百KB級になっていたらPurgeが効いていない）
