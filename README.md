# 家電量販店7社 業績比較ダッシュボード (FY2026/3 通期)

家電量販店主要7社（ヤマダHD・ケーズHD・エディオン・上新電機・ノジマ・ビックカメラ・コジマ）の
通期業績データを、各社公式IR決算短信・決算説明会資料から抽出・正規化したダッシュボード。

## 公開ページ

GitHub Pages: `https://matsuno-delgit.github.io/kaden-retail-7co-fy2026/`

## 構成

```
.
├── index.html                       # ダッシュボード本体（単一HTML, Chart.js + 自己ホストTailwind）
├── assets/
│   └── tailwind.css                 # Tailwind 利用クラス抽出版（自己ホスト）
├── data/
│   └── companies.json               # ダッシュボード用統合データ（7社）
├── scripts/
│   ├── consolidate.py               # Excel _ver.4 → companies.json 生成
│   └── regenerate-tailwind.md       # tailwind.css 再生成手順
├── SECURITY.md                      # セキュリティ方針
├── LICENSE                          # MIT
└── README.md
```

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

- 各社の主要20項目（売上高/営業利益/経常利益/純利益/総資産/自己資本 等）×当期/前期=140セル
- すべて短信PDFと **完全一致**を確認済み（[検証レポート](../検証レポート_20260522.md) 参照）
- セグメント情報（デンキセグ・ノジマ家電専門店・ビック単体）も決算説明会資料と一致確認済

## ローカルでの確認

```bash
cd kaden-retail-7co-fy2026
python -m http.server 8000
# ブラウザで http://localhost:8000/ を開く
```

## データ更新手順

```bash
# 1. 新しい四半期/通期PDFをエビデンスフォルダに配置
# 2. extract_v2.py で抽出・verify.py で検証
# 3. Excel _ver.X を更新
# 4. consolidate.py で companies.json 再生成
cd kaden-retail-7co-fy2026
uv run --no-project python scripts/consolidate.py
git add data/companies.json
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

- 抽出スキーム: `extract_excerpt.py` / `extract_v2.py` / `verify.py`（このリポジトリの親フォルダ）
- 元Excel: `【経営企画部】各社業績対比フォーマット（2026.03通期）_..._ver.4.xlsx`
- 生成補助: [github-dashboard-publisher](https://github.com/) スキル（Claude Code）
