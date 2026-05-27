"""組替版ダッシュボード用 companies_kumikae_<period>.json を一括生成。

入力Excel (期間揃え比較フォルダ):
  当期版    1_組替通期_2026.02 etc.    (1_-8_)
  前々期版  9_組替通期_2025.02 etc.   (9_-16_)

期間ラベル統一: 全社「FY2026 (=2025/3〜2026/2 系)」「FY2025」「FY2024」を採用。
forecast/forecast_annual は組替版では使わない (null)。
"""
import json
from datetime import date
from pathlib import Path
from openpyxl import load_workbook

ROOT = Path(__file__).parent.parent.parent  # 01_通期実績_2026.03
PROJ = ROOT.parent  # 競合各社業績比較_20260520
KUMIKAE = PROJ / "期間揃え比較"
OUT_DIR = Path(__file__).parent.parent / "data"
OUT_DIR.mkdir(exist_ok=True)

# 期間別 (curr_xlsx, pp_xlsx) ペア
PERIODS = {
    "fy": {
        "label": "組替通期 (2025/3〜2026/2)",
        "curr": KUMIKAE / "1_組替通期_2026.02" / "【組替通期】各社業績対比フォーマット（2025.03_2026.02 vs 2024.03_2025.02）_ver.1.xlsx",
        "pp":   KUMIKAE / "9_組替通期_2025.02"  / "【組替通期】各社業績対比フォーマット（2024.03_2025.02 vs 2023.03_2024.02）_ver.1.xlsx",
    },
    "q1": {
        "label": "組替1Q (3〜5月)",
        "curr": KUMIKAE / "2_組替1Q_2025.05"   / "【組替1Q】各社業績対比フォーマット（2025.03_2025.05 vs 2024.03_2024.05）_ver.1.xlsx",
        "pp":   KUMIKAE / "10_組替1Q_2024.05"  / "【組替1Q】各社業績対比フォーマット（2024.03_2024.05 vs 2023.03_2023.05）_ver.1.xlsx",
    },
    "q2": {
        "label": "組替2Q単独 (6〜8月)",
        "curr": KUMIKAE / "3_組替2Q単独_2025.08" / "【組替2Q単独】各社業績対比フォーマット（2025.06_2025.08 vs 2024.06_2024.08）_ver.1.xlsx",
        "pp":   KUMIKAE / "11_組替2Q単独_2024.08" / "【組替2Q単独】各社業績対比フォーマット（2024.06_2024.08 vs 2023.06_2023.08）_ver.1.xlsx",
    },
    "h1": {
        "label": "組替上期 (3〜8月)",
        "curr": KUMIKAE / "4_組替上期_2025.08" / "【組替上期】各社業績対比フォーマット（2025.03_2025.08 vs 2024.03_2024.08）_ver.1.xlsx",
        "pp":   KUMIKAE / "12_組替上期_2024.08" / "【組替上期】各社業績対比フォーマット（2024.03_2024.08 vs 2023.03_2023.08）_ver.1.xlsx",
    },
    "q3": {
        "label": "組替3Q単独 (9〜11月)",
        "curr": KUMIKAE / "5_組替3Q単独_2025.11" / "【組替3Q単独】各社業績対比フォーマット（2025.09_2025.11 vs 2024.09_2024.11）_ver.1.xlsx",
        "pp":   KUMIKAE / "13_組替3Q単独_2024.11" / "【組替3Q単独】各社業績対比フォーマット（2024.09_2024.11 vs 2023.09_2023.11）_ver.1.xlsx",
    },
    "q3cum": {
        "label": "組替3Q累計 (3〜11月)",
        "curr": KUMIKAE / "6_組替3Q累計_2025.11" / "【組替3Q累計】各社業績対比フォーマット（2025.03_2025.11 vs 2024.03_2024.11）_ver.1.xlsx",
        "pp":   KUMIKAE / "14_組替3Q累計_2024.11" / "【組替3Q累計】各社業績対比フォーマット（2024.03_2024.11 vs 2023.03_2023.11）_ver.1.xlsx",
    },
    "q4": {
        "label": "組替4Q単独 (12〜翌2月)",
        "curr": KUMIKAE / "7_組替4Q単独_2026.02" / "【組替4Q単独】各社業績対比フォーマット（2025.12_2026.02 vs 2024.12_2025.02）_ver.1.xlsx",
        "pp":   KUMIKAE / "15_組替4Q単独_2025.02" / "【組替4Q単独】各社業績対比フォーマット（2024.12_2025.02 vs 2023.12_2024.02）_ver.1.xlsx",
    },
    "h2": {
        "label": "組替下期 (9〜翌2月)",
        "curr": KUMIKAE / "8_組替下期_2026.02" / "【組替下期】各社業績対比フォーマット（2025.09_2026.02 vs 2024.09_2025.02）_ver.1.xlsx",
        "pp":   KUMIKAE / "16_組替下期_2025.02" / "【組替下期】各社業績対比フォーマット（2024.09_2025.02 vs 2023.09_2024.02）_ver.1.xlsx",
    },
}

COMPANIES_MAIN = [
    {"key": "yamada", "label": "ヤマダHD", "ticker": "9831", "fy_end": "03", "consol": "consolidated", "col_curr": 4, "col_prev": 5, "url": "https://www.yamada-holdings.jp/ir/"},
    {"key": "yamada_denki", "label": "ヤマダ（デンキセグメント）", "ticker": "9831-DK", "fy_end": "03", "consol": "segment", "col_curr": 6, "col_prev": 7, "url": "https://www.yamada-holdings.jp/ir/", "is_segment": True},
    {"key": "ks", "label": "ケーズHD", "ticker": "8282", "fy_end": "03", "consol": "consolidated", "col_curr": 8, "col_prev": 9, "url": "https://www.ksdenki.co.jp/ir/"},
    {"key": "edion", "label": "エディオン", "ticker": "2730", "fy_end": "03", "consol": "consolidated", "col_curr": 10, "col_prev": 11, "url": "https://www.edion.co.jp/ir/"},
    {"key": "joshin", "label": "上新電機", "ticker": "8173", "fy_end": "03", "consol": "consolidated", "col_curr": 12, "col_prev": 13, "url": "https://www.joshin.co.jp/ir/"},
    {"key": "nojima", "label": "ノジマ", "ticker": "7419", "fy_end": "03", "consol": "consolidated", "col_curr": 14, "col_prev": 15, "url": "https://www.nojima.co.jp/ir/"},
    {"key": "bic", "label": "ビックカメラ（連結）", "ticker": "3048", "fy_end": "08", "consol": "consolidated", "col_curr": 22, "col_prev": 23, "url": "https://www.biccamera.co.jp/ir/"},
    {"key": "kojima", "label": "コジマ", "ticker": "7513", "fy_end": "08", "consol": "non_consolidated", "col_curr": 18, "col_prev": 19, "url": "https://www.kojima.net/corporation/ir/"},
]

ROW_TO_METRIC = {
    7: "Revenue", 8: "GrossProfit", 9: "SGA",
    82: "OperatingIncome", 83: "OrdinaryIncome", 84: "NetIncome",
    88: "InterestBearingDebt", 89: "TotalEquity",
    100: "TotalAssets", 101: "Inventory", 86: "Tax",
}


def to_num(v):
    if v is None or (isinstance(v, str) and v.startswith("=")):
        return None
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def safe_div(a, b):
    try:
        return a / b if a is not None and b not in (None, 0) else None
    except (TypeError, ZeroDivisionError):
        return None


def build_one(pk, conf, is_annual):
    print(f"\n========== {pk} ({conf['label']}) ==========")
    wb_curr = load_workbook(conf["curr"], data_only=True); ws_curr = wb_curr.active
    wb_pp = load_workbook(conf["pp"], data_only=True); ws_pp = wb_pp.active

    companies_out = []
    for co in COMPANIES_MAIN:
        # 統一期間ラベル: 全社FY2026/FY2025/FY2024
        current_period = "FY2026"
        previous_period = "FY2025"
        prev_previous_period = "FY2024"
        forecast_period = "FY2027"

        metrics = {}
        for row, key in ROW_TO_METRIC.items():
            cur = to_num(ws_curr.cell(row=row, column=co["col_curr"]).value)
            prev = to_num(ws_curr.cell(row=row, column=co["col_prev"]).value)
            # 前々期: 9_-16_の前期列 (= 2023/3-2024/2 期)
            prev_prev = to_num(ws_pp.cell(row=row, column=co["col_prev"]).value)
            metrics[key] = {
                "prev_previous": prev_prev,
                "current": cur,
                "previous": prev,
                "forecast": None,
                "forecast_annual": None,
                "unit": "百万円",
            }

        yoy = {}
        for key in ("Revenue", "OperatingIncome", "OrdinaryIncome", "NetIncome"):
            m = metrics.get(key, {})
            cur = m.get("current"); prev = m.get("previous"); pp = m.get("prev_previous")
            yoy[key] = {
                "current_yoy_pct": round((cur / prev - 1) * 100, 2) if cur and prev else None,
                "previous_yoy_pct": round((prev / pp - 1) * 100, 2) if prev and pp else None,
                "forecast_yoy_pct": None,
            }

        rev = metrics["Revenue"]["current"]
        op = metrics["OperatingIncome"]["current"]
        ta = metrics["TotalAssets"]["current"]
        te = metrics["TotalEquity"]["current"]
        ord_c = metrics["OrdinaryIncome"]["current"]
        rev_p = metrics["Revenue"]["previous"]
        ord_p = metrics["OrdinaryIncome"]["previous"]
        rev_pp = metrics["Revenue"]["prev_previous"]
        ord_pp = metrics["OrdinaryIncome"]["prev_previous"]
        gp_c = metrics["GrossProfit"]["current"]
        gp_p = metrics["GrossProfit"]["previous"]
        gp_pp = metrics["GrossProfit"]["prev_previous"]
        ratios = {
            "operating_margin_pct":              round(op / rev * 100, 2) if op and rev else None,
            "ordinary_margin_pct":               round(ord_c / rev * 100, 2) if ord_c and rev else None,
            "ordinary_margin_pct_previous":      round(ord_p / rev_p * 100, 2) if ord_p and rev_p else None,
            "ordinary_margin_pct_prev_previous": round(ord_pp / rev_pp * 100, 2) if ord_pp and rev_pp else None,
            "equity_ratio_pct":                  round(te / ta * 100, 2) if te and ta else None,
            "gross_margin_pct":                  round(gp_c / rev * 100, 2) if gp_c and rev else None,
            "gross_margin_pct_previous":         round(gp_p / rev_p * 100, 2) if gp_p and rev_p else None,
            "gross_margin_pct_prev_previous":    round(gp_pp / rev_pp * 100, 2) if gp_pp and rev_pp else None,
        }

        # trend_ratios (通期のみ計算、その他はnull)
        EFFECTIVE_TAX_RATE = 0.35

        def roe(ni, eq):
            v = safe_div(ni, eq)
            return round(v * 100, 2) if v is not None else None

        def roic(op_v, debt_v, eq_v):
            ic = (debt_v or 0) + (eq_v or 0)
            if not op_v or ic == 0:
                return None
            return round(op_v * (1 - EFFECTIVE_TAX_RATE) / ic * 100, 2)

        def turnover(num, den):
            v = safe_div(num, den)
            return round(v, 3) if v is not None else None

        trend_ratios = {"roe": {}, "roic": {}, "asset_turnover": {}, "inventory_turnover": {}}
        if is_annual:
            bs_periods = {
                "prev_previous": {"ta": metrics["TotalAssets"]["prev_previous"], "te": metrics["TotalEquity"]["prev_previous"], "inv": metrics["Inventory"]["prev_previous"], "debt": metrics["InterestBearingDebt"]["prev_previous"]},
                "previous":      {"ta": metrics["TotalAssets"]["previous"],      "te": metrics["TotalEquity"]["previous"],      "inv": metrics["Inventory"]["previous"],      "debt": metrics["InterestBearingDebt"]["previous"]},
                "current":       {"ta": metrics["TotalAssets"]["current"],       "te": metrics["TotalEquity"]["current"],       "inv": metrics["Inventory"]["current"],       "debt": metrics["InterestBearingDebt"]["current"]},
            }
            for period in ("prev_previous", "previous", "current"):
                rev_p2 = metrics["Revenue"][period]
                op_p2 = metrics["OperatingIncome"][period]
                ni_p2 = metrics["NetIncome"][period]
                bs = bs_periods[period]
                trend_ratios["roe"][period] = roe(ni_p2, bs["te"])
                trend_ratios["roic"][period] = roic(op_p2, bs["debt"], bs["te"])
                trend_ratios["asset_turnover"][period] = turnover(rev_p2, bs["ta"])
                trend_ratios["inventory_turnover"][period] = turnover(rev_p2, bs["inv"])
            for k in trend_ratios.keys():
                trend_ratios[k]["forecast"] = None
        else:
            for k in trend_ratios.keys():
                trend_ratios[k] = {"prev_previous": None, "previous": None, "current": None, "forecast": None}

        # デンキセグ: BS非開示、在庫回転率POSベース固定値
        if co.get("is_segment"):
            for k in ("roe", "roic", "asset_turnover"):
                trend_ratios[k] = {"prev_previous": None, "previous": None, "current": None, "forecast": None}
            if is_annual:
                trend_ratios["inventory_turnover"] = {
                    "prev_previous": 3.6, "previous": 4.0, "current": 4.5, "forecast": None,
                }
            ratios["equity_ratio_pct"] = None

        companies_out.append({
            "key": co["key"], "label": co["label"], "ticker": co["ticker"],
            "fiscal_year_end": co["fy_end"], "consolidation": co["consol"],
            "current_period": current_period, "previous_period": previous_period,
            "prev_previous_period": prev_previous_period, "forecast_period": forecast_period,
            "tanshin_url": co["url"], "metrics": metrics, "yoy": yoy, "ratios": ratios,
            "trend_ratios": trend_ratios, "is_segment": bool(co.get("is_segment")),
        })

    output = {
        "schema_version": "1.0",
        "generated_at": date.today().isoformat(),
        "source": f"組替版 ({conf['label']})",
        "period_type": f"kumikae_{pk}",
        "dataset_type": "kumikae",
        "companies": companies_out,
    }
    out_path = OUT_DIR / f"companies_kumikae_{pk}.json"
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(output, f, ensure_ascii=False, indent=2)
    print(f"  → {out_path.name}")
    for c in companies_out:
        rev = c["metrics"]["Revenue"]["current"]
        ord_v = c["metrics"]["OrdinaryIncome"]["current"]
        om = c["ratios"]["ordinary_margin_pct"]
        rev_s = f"{rev:,.0f}" if isinstance(rev,(int,float)) else "—"
        ord_s = f"{ord_v:,.0f}" if isinstance(ord_v,(int,float)) else "—"
        om_s = f"{om}%" if om is not None else "—"
        print(f"    {c['label']:20s}: 売上={rev_s:>11s} / 経常={ord_s:>8s} / 経常率={om_s}")


def main():
    for pk, conf in PERIODS.items():
        build_one(pk, conf, is_annual=(pk == "fy"))


if __name__ == "__main__":
    main()
