"""Excel ファイル(_ver.4.xlsx)から、ダッシュボード data/companies.json を生成する。

リファレンス index.html が期待するスキーマ:
{
  "schema_version": "1.0",
  "generated_at": "YYYY-MM-DD",
  "companies": [
    {
      "key": "yamada",
      "label": "ヤマダHD",
      "ticker": "9831",
      "fiscal_year_end": "03",
      "consolidation": "consolidated",
      "current_period": "FY2026",
      "previous_period": "FY2025",
      "forecast_period": "FY2027",
      "tanshin_url": "...",
      "metrics": {
        "Revenue": {"current":, "previous":, "forecast":, "unit": "百万円"},
        ...
      },
      "yoy": {"Revenue": {"current_yoy_pct":, "forecast_yoy_pct":}, ...},
      "ratios": {"operating_margin_pct":, "equity_ratio_pct":}
    },
    ...
  ]
}

メイン7社のみを対象とする（家電量販店7社業績比較）。
セグメント単位（デンキセグ・ノジマ家電専門店・ビックカメラ単体）は補足セクションに別途。
"""
import json
from datetime import date
from pathlib import Path
from openpyxl import load_workbook

# プロジェクトルートからの相対パス
ROOT = Path(__file__).parent.parent.parent  # kaden-retail-7co-fy2026 の親
XLSX = ROOT / "【経営企画部】各社業績対比フォーマット（2026.03通期）_エディオン2026.3期反映_20260521_ver.4.xlsx"
OUT_DIR = Path(__file__).parent.parent / "data"
OUT_DIR.mkdir(exist_ok=True)

# 7社の定義（Excel列レイアウト + メタ）
COMPANIES_MAIN = [
    {
        "key": "yamada", "label": "ヤマダHD", "ticker": "9831",
        "fiscal_year_end": "03", "consolidation": "consolidated",
        "current_period": "FY2026", "previous_period": "FY2025", "forecast_period": "FY2027",
        "col_curr": 4, "col_prev": 5,
        "tanshin_url": "https://www.yamada-holdings.jp/ir/",
    },
    {
        "key": "ks", "label": "ケーズHD", "ticker": "8282",
        "fiscal_year_end": "03", "consolidation": "consolidated",
        "current_period": "FY2026", "previous_period": "FY2025", "forecast_period": "FY2027",
        "col_curr": 8, "col_prev": 9,
        "tanshin_url": "https://www.ksdenki.co.jp/ir/",
    },
    {
        "key": "edion", "label": "エディオン", "ticker": "2730",
        "fiscal_year_end": "03", "consolidation": "consolidated",
        "current_period": "FY2026", "previous_period": "FY2025", "forecast_period": "FY2027",
        "col_curr": 10, "col_prev": 11,
        "tanshin_url": "https://www.edion.co.jp/ir/",
    },
    {
        "key": "joshin", "label": "上新電機", "ticker": "8173",
        "fiscal_year_end": "03", "consolidation": "consolidated",
        "current_period": "FY2026", "previous_period": "FY2025", "forecast_period": "FY2027",
        "col_curr": 12, "col_prev": 13,
        "tanshin_url": "https://www.joshin.co.jp/ir/",
    },
    {
        "key": "nojima", "label": "ノジマ", "ticker": "7419",
        "fiscal_year_end": "03", "consolidation": "consolidated",
        "current_period": "FY2026", "previous_period": "FY2025", "forecast_period": "FY2027",
        "col_curr": 14, "col_prev": 15,
        "tanshin_url": "https://www.nojima.co.jp/ir/",
    },
    {
        "key": "bic", "label": "ビックカメラ", "ticker": "3048",
        "fiscal_year_end": "08", "consolidation": "consolidated",
        "current_period": "FY2025", "previous_period": "FY2024", "forecast_period": "FY2026",
        "col_curr": 22, "col_prev": 23,
        "tanshin_url": "https://www.biccamera.co.jp/ir/",
    },
    {
        "key": "kojima", "label": "コジマ", "ticker": "7513",
        "fiscal_year_end": "08", "consolidation": "non_consolidated",
        "current_period": "FY2025", "previous_period": "FY2024", "forecast_period": "FY2026",
        "col_curr": 18, "col_prev": 19,
        "tanshin_url": "https://www.kojima.net/corporation/ir/",
    },
]

# Excel行 → メトリクスキー (リファレンスindex.htmlで使うキー名)
ROW_TO_METRIC = {
    7:  "Revenue",
    8:  "GrossProfit",
    9:  "SGA",
    82: "OperatingIncome",
    83: "OrdinaryIncome",
    84: "NetIncome",
    88: "InterestBearingDebt",
    89: "TotalEquity",         # 自己資本
    91: "DividendPerShare",
    92: "DividendTotal",
    93: "PayoutRatio",
    96: "EPS",
    100: "TotalAssets",
    101: "Inventory",
}

# 業績予想（短信表紙から）
FORECAST = {
    "yamada": {"Revenue": 1780000, "OperatingIncome": 51500, "OrdinaryIncome": 52600, "NetIncome": 27800},
    "ks":     {"Revenue": 785000,  "OperatingIncome": 30500, "OrdinaryIncome": 33500, "NetIncome": 20000},
    "edion":  {"Revenue": None,    "OperatingIncome": None,  "OrdinaryIncome": None,  "NetIncome": None},   # 未確認
    "joshin": {"Revenue": None,    "OperatingIncome": None,  "OrdinaryIncome": None,  "NetIncome": None},   # 未確認
    "nojima": {"Revenue": 1000000, "OperatingIncome": 59000, "OrdinaryIncome": 76000, "NetIncome": 48000},
    "bic":    {"Revenue": None,    "OperatingIncome": None,  "OrdinaryIncome": None,  "NetIncome": None},   # 8月決算で記載なし
    "kojima": {"Revenue": 294000,  "OperatingIncome": 7600,  "OrdinaryIncome": 7900,  "NetIncome": 4900},
}


def to_num(v):
    if v is None or (isinstance(v, str) and v.startswith("=")):
        return None
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def main():
    wb = load_workbook(XLSX, data_only=False)
    ws = wb["PL・BSデータ (2026.通期)"]

    companies_out = []
    for co in COMPANIES_MAIN:
        metrics = {}
        for row, key in ROW_TO_METRIC.items():
            cur = to_num(ws.cell(row=row, column=co["col_curr"]).value)
            prev = to_num(ws.cell(row=row, column=co["col_prev"]).value)
            fc = FORECAST.get(co["key"], {}).get(key)
            metrics[key] = {
                "current": cur,
                "previous": prev,
                "forecast": fc,
                "unit": "百万円" if key not in ("DividendPerShare", "EPS", "PayoutRatio") else (
                    "円" if key in ("DividendPerShare", "EPS") else "%"
                ),
            }

        # 派生: YoY % (current vs previous), forecast YoY %
        yoy = {}
        for key in ("Revenue", "OperatingIncome", "OrdinaryIncome", "NetIncome"):
            m = metrics.get(key, {})
            cur = m.get("current")
            prev = m.get("previous")
            fc = m.get("forecast")
            yoy[key] = {
                "current_yoy_pct": round((cur / prev - 1) * 100, 2) if cur and prev else None,
                "forecast_yoy_pct": round((fc / cur - 1) * 100, 2) if fc and cur else None,
            }

        # 比率
        rev = metrics["Revenue"]["current"]
        op = metrics["OperatingIncome"]["current"]
        ta = metrics["TotalAssets"]["current"]
        te = metrics["TotalEquity"]["current"]
        ratios = {
            "operating_margin_pct": round(op / rev * 100, 2) if op and rev else None,
            "equity_ratio_pct": round(te / ta * 100, 2) if te and ta else None,
            "gross_margin_pct": round(metrics["GrossProfit"]["current"] / rev * 100, 2)
                if metrics["GrossProfit"]["current"] and rev else None,
        }

        co_out = {
            "key": co["key"],
            "label": co["label"],
            "ticker": co["ticker"],
            "fiscal_year_end": co["fiscal_year_end"],
            "consolidation": co["consolidation"],
            "current_period": co["current_period"],
            "previous_period": co["previous_period"],
            "forecast_period": co["forecast_period"],
            "tanshin_url": co["tanshin_url"],
            "metrics": metrics,
            "yoy": yoy,
            "ratios": ratios,
        }
        companies_out.append(co_out)

    output = {
        "schema_version": "1.0",
        "generated_at": date.today().isoformat(),
        "source": "各社決算短信・決算説明会資料 (TDnet公開資料、2026/5月発表分)",
        "note": "ヤマダ・ケーズ・エディオン・上新・ノジマ=2026年3月期通期。ビックカメラ・コジマ=2025年8月期通期。コジマは非連結開示。",
        "companies": companies_out,
    }

    out_path = OUT_DIR / "companies.json"
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(output, f, ensure_ascii=False, indent=2)
    print(f"Wrote {out_path}")
    print(f"Companies: {len(companies_out)}")
    for c in companies_out:
        rev = c["metrics"]["Revenue"]["current"]
        op = c["metrics"]["OperatingIncome"]["current"]
        om = c["ratios"]["operating_margin_pct"]
        print(f"  {c['label']:10s} ({c['current_period']}): 売上 {rev:>12,.0f} / 営利 {op:>8,.0f} / 営利率 {om}%")


if __name__ == "__main__":
    main()
