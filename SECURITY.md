# セキュリティ方針 / Security Notes

このリポジトリは静的ダッシュボード（GitHub Pages公開）であり、サーバ側ロジックや認証情報を持ちません。
ただし、開発・運用上のセキュリティ留意点を以下にまとめます。

## 含まれるデータの性質

- 公開済みの一次情報（例: 法定開示資料・公的統計）から抽出した数値のみ
- 本リポジトリは数値（JSON/CSV）とそれを可視化するHTMLのみを含む
- 個人情報、認証情報、機微情報は**一切含まない**

## 監査済み項目

- リポジトリ内に API key / token / secret / password 文字列が含まれないこと
- コミット作者メールがGitHub noreply形式（`<id>+<username>@users.noreply.github.com`）であること
- GitHub Secret Scanning + Push Protection が有効であること

## 外部依存

`index.html` から参照しているJSライブラリは以下のみ。
**バージョンを固定 + SRI（Subresource Integrity）ハッシュ**を指定し、
CDNが改ざんされた場合は実行を拒否する。

| ライブラリ | バージョン | CDN | SRI |
|---|---|---|---|
| Chart.js | <VERSION> | cdn.jsdelivr.net | sha384-<HASH> |
| <他のライブラリ> | ... | ... | ... |

CSSは Tailwind CSS の利用クラスのみを抽出した静的ファイル（`assets/tailwind.css`）を自己ホストしており、
CDN動的生成（`cdn.tailwindcss.com`）に依存していません。

## GitHub CLI トークンスコープの最小化

このリポジトリの運用に必要な GitHub CLI スコープは `repo` のみです。
他のスコープ（`gist` / `workflow` 等）を付与している場合は、最小化することを推奨します。

### 確認方法

```powershell
gh auth status
```

### 最小化手順

GitHub CLI の `auth refresh` は**スコープを追加するだけ**で削減はできないため、
一度ログアウトして再ログインします。

```powershell
gh auth logout --hostname github.com
gh auth login --hostname github.com --scopes repo --web
gh auth status   # Token scopes: 'repo' になっていればOK
```

## 失効した token の取り扱い

万一トークンが漏洩した・漏洩した可能性がある場合：

1. https://github.com/settings/tokens でただちに当該トークンをRevoke
2. もしOAuth経由なら https://github.com/settings/applications で GitHub CLI を Revoke
3. `gh auth login` で再ログイン
4. 直近のコミット内容を確認し、機微情報の漏洩がないか検査
5. 必要に応じて `git filter-repo` で履歴から削除

## データダウンロード時のTLS設定

データ取得スクリプトは **証明書失効チェックを有効にしたまま** で動作する想定です。
`--ssl-no-revoke` 等のチェック無効化フラグは原則使用しません。

## 報告先

本リポジトリのセキュリティ上の問題を発見された場合は、GitHub Issues ではなく、
リポジトリ所有者にプライベート連絡でご報告ください。
