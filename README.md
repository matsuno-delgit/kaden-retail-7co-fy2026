# 家電量販店7社 業績比較ダッシュボード

家電量販店主要7社（ヤマダHD・ケーズHD・エディオン・上新電機・ノジマ・ビックカメラ・コジマ）の
業績データを、各社公式IR決算短信・決算説明会資料から抽出・正規化したダッシュボード。
通期のほか、1Q/2Q単独/上期/3Q単独/3Q累計/4Q単独/下期の各期タブと、
8月決算2社（ビックカメラ・コジマ）を3月決算社の期間に揃えた「組替」データセット
トグルを備える。

## 公開ページ

GitHub Pages: `https://matsuno-delgit.github.io/kaden-retail-7co-fy2026/`

## 構成

```
.
├── index.html                       # ダッシュボード本体（単一HTML, Chart.js + 自己ホストTailwind）
├── assets/
│   └── tailwind.css                 # Tailwind 利用クラス抽出版（自己ホスト）
├── data/
│   ├── companies.json               # 通期（実績+前々期+計画）
│   ├── companies_{q1,q2,h1,q3,q3cum,q4,h2}.json      # 四半期・上下期 各タブ
│   └── companies_kumikae_{fy,q1,q2,h1,q3,q3cum,q4,h2}.json  # 組替（期間揃え）
├── scripts/
│   ├── consolidate.py               # 通期実績Excel → companies.json
│   ├── consolidate_q1.py            # 1Q Excel → companies_q1.json
│   ├── consolidate_quarter.py       # 2Q/上期/3Q/3Q累計/4Q/下期 → companies_*.json
│   ├── consolidate_kumikae.py       # 期間揃え比較Excel → companies_kumikae_*.json
│   ├── xlsx_utils.py                # 共通: 最新ver自動検出・データシート特定
│   └── regenerate-tailwind.md       # tailwind.css 再生成手順
├── SECURITY.md                      # セキュリティ方針
├── LICENSE                          # MIT
└── README.md
```

入力Excelは親フォルダ群（`01_通期実績_2026.03/` ほか）の**最新版（`_ver.N` 最大）を
自動検出**する。Excel更新時にスクリプトの書き換えは不要。

## 対象データ・出典

| 会社 | 証券コード | 期 | 連結区分 | 出典 |
|---|---|---|---|---|
| ヤマダHD | 9831 | 2026年3月期 通期 | 連結 | TDnet (2026/5/8) |
| ケーズHD | 8282 | 2026年3月期 通期 | 連結 | TDnet (2026/5/8) |
| エディオン | 2730 | 2026年3月期 通期 | 連結 | TDnet (2026/5/11) |
| 上新電機 | 8173 | 2026年3月期 通期 | 連結 | TDnet (2026/5/8) |
| ノジマ | 7419 | 2026年3月期 通期 | 連結 | TDnet (2026/5/7) |
| ビックカメラ | 3048 | 2025年8月期 通期 | 連結 | TDnet (2025/10/10) |
| コジマ | 7513 | 2025年8月期 通期 | 非連結 | TDnet (2025/10/9) |

ビックカメラ・コジマは8月決算のため、表示時期が他社（3月決算）と異なります。横並び比較は参考値として扱ってください。

## データ検証ステータス

- 通期: 各社の主要20項目（売上高/営業利益/経常利益/純利益/総資産/自己資本 等）
  ×当期/前期=140セルすべて短信PDFと**完全一致**を確認済み
- 四半期・上下期・前々期・組替の各データセットも、PDF突合＋差引クロスフット＋
  組替整合検証（`00_共通スクリプト/verify_retro.py`、2026-07-12 実施）で全件一致を確認済み
- セグメント情報（デンキセグ・ノジマ家電専門店・ビック単体）も決算説明会資料と一致確認済
- 詳細はプロジェクトフォルダ内の「検証レポート_20260522.md」（GitHubには含まれない
  ローカル文書）および各期フォルダの `verify_*_out.txt` を参照

## ローカルでの確認

```bash
cd kaden-retail-7co-fy2026
python -m http.server 8000
# ブラウザで http://localhost:8000/ を開く
```

## データ更新手順

```bash
# 1. 新しい四半期/通期PDFをエビデンスフォルダに配置
# 2. extract系スクリプトで抽出・verify系スクリプトで検証（親フォルダ側の作業）
# 3. Excel を新しい _ver.N+1 として保存（スクリプトは最新verを自動検出）
# 4. consolidate 4本で data/*.json を再生成
cd kaden-retail-7co-fy2026
uv run --no-project python scripts/consolidate.py
uv run --no-project python scripts/consolidate_q1.py
uv run --no-project python scripts/consolidate_quarter.py
uv run --no-project python scripts/consolidate_kumikae.py
git diff data/          # 差分がデータとして妥当か確認
git add data/
git commit -m "Update <FY/期>"
git push
```

push後、GitHub Pages が1〜2分で自動再ビルドされ反映されます。

## セキュリティ

- 外部CDN依存（Chart.js, chartjs-plugin-datalabels）は **SRIハッシュ付き** で改ざん検知
- CSSは Tailwind の利用クラスを抽出した **自己ホスト版**（cdn.tailwindcss.com 等の動的生成CDN依存なし）
- 詳細は [SECURITY.md](SECURITY.md) を参照

## ライセンス

- ソースコード: MIT License
- データ: 各社公式IR開示資料の引用。投資判断に用いる場合は必ず原資料を参照してください

## 関連

- 抽出スキーム: `extract_excerpt.py` / `extract_v2.py` / `verify.py` / `verify_retro.py`（このリポジトリの親フォルダ群）
- 元Excel: `【経営企画部】各社業績対比フォーマット（…）_ver.N.xlsx`（各期フォルダの最新verを自動検出）
- 生成補助: [github-dashboard-publisher](https://github.com/) スキル（Claude Code）
