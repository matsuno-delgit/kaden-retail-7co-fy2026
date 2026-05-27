"""四半期/上期/下期ダッシュボード用 companies_<period>.json を一括生成。

サポート期間: q2 (2Q単独), h1 (上期), q3 (3Q単独), q3cum (3Q累計), q4 (4Q単独), h2 (下期)

データソース:
  当期/前期: 各 04/05/06/07/08/09_*_2026.03 Excel の最新ver
  前々期:    各 12/13/14/15/16/17_*_2024.03 Excel の最新ver

usage:
  python consolidate_quarter.py q2 h1 q3 q3cum q4 h2
  または引数なし → 全期間生成
"""
import json
import sys
from datetime import date
from pathlib import Path
from openpyxl import load_workbook

ROOT = Path(__file__).parent.parent.parent  # 01_通期実績_2026.03
PROJ = ROOT.parent  # 競合各社業績比較_20260520
OUT_DIR = Path(__file__).parent.parent / "data"
OUT_DIR.mkdir(exist_ok=True)

# (period_key, label, ja_short, 当期Excelパターン, 前々期Excelパターン, current_period suffix(3月), suffix(8月))
PERIODS = {
    "q2": {
        "label": "2Q単独",
        "xlsx_curr": PROJ / "04_2Q単独_2026.03" / "【経営企画部】各社業績対比フォーマット（2026.03第2四半期単独 vs 2025.03第2四半期単独）_ver.7.xlsx",
        "xlsx_pp":   PROJ / "12_2Q単独_2024.03" / "【経営企画部】各社業績対比フォーマット（2024.03第2四半期単独 vs 2023.03第2四半期単独）_ver.3.xlsx",
        "curr_suffix_3": "-Q2", "curr_suffix_8": "-Q2",
        "prev_suffix_3": "-Q2", "prev_suffix_8": "-Q2",
        "pp_suffix_3":   "-Q2", "pp_suffix_8":   "-Q2",
    },
    "h1": {
        "label": "上期",
        "xlsx_curr": PROJ / "05_上期実績_2026.03" / "【経営企画部】各社業績対比フォーマット（2026.03上期実績 vs 2025.03上期実績）_ver.8.xlsx",
        "xlsx_pp":   PROJ / "13_上期実績_2024.03" / "【経営企画部】各社業績対比フォーマット（2024.03上期実績 vs 2023.03上期実績）_ver.4.xlsx",
        "curr_suffix_3": "-H1", "curr_suffix_8": "-H1",
        "prev_suffix_3": "-H1", "prev_suffix_8": "-H1",
        "pp_suffix_3":   "-H1", "pp_suffix_8":   "-H1",
    },
    "q3": {
        "label": "3Q単独",
        "xlsx_curr": PROJ / "06_3Q単独_2026.03" / "【経営企画部】各社業績対比フォーマット（2026.03第3四半期単独 vs 2025.03第3四半期単独）_ver.6.xlsx",
        "xlsx_pp":   PROJ / "14_3Q単独_2024.03" / "【経営企画部】各社業績対比フォーマット（2024.03第3四半期単独 vs 2023.03第3四半期単独）_ver.3.xlsx",
        "curr_suffix_3": "-Q3", "curr_suffix_8": "-Q3",
        "prev_suffix_3": "-Q3", "prev_suffix_8": "-Q3",
        "pp_suffix_3":   "-Q3", "pp_suffix_8":   "-Q3",
    },
    "q3cum": {
        "label": "3Q累計",
        "xlsx_curr": PROJ / "07_3Q累計_2026.03" / "【経営企画部】各社業績対比フォーマット（2026.03第3四半期累計 vs 2025.03第3四半期累計）_ver.6.xlsx",
        "xlsx_pp":   PROJ / "15_3Q累計_2024.03" / "【経営企画部】各社業績対比フォーマット（2024.03第3四半期累計 vs 2023.03第3四半期累計）_ver.5.xlsx",
        "curr_suffix_3": "-Q3CUM", "curr_suffix_8": "-Q3CUM",
        "prev_suffix_3": "-Q3CUM", "prev_suffix_8": "-Q3CUM",
        "pp_suffix_3":   "-Q3CUM", "pp_suffix_8":   "-Q3CUM",
    },
    "q4": {
        "label": "4Q単独",
        "xlsx_curr": PROJ / "08_4Q単独_2026.03" / "【経営企画部】各社業績対比フォーマット（2026.03第4四半期単独 vs 2025.03第4四半期単独）_ver.5.xlsx",
        "xlsx_pp":   PROJ / "16_4Q単独_2024.03" / "【経営企画部】各社業績対比フォーマット（2024.03第4四半期単独 vs 2023.03第4四半期単独）_ver.4.xlsx",
        "curr_suffix_3": "-Q4", "curr_suffix_8": "-Q4",
        "prev_suffix_3": "-Q4", "prev_suffix_8": "-Q4",
        "pp_suffix_3":   "-Q4", "pp_suffix_8":   "-Q4",
    },
    "h2": {
        "label": "下期",
        "xlsx_curr": PROJ / "09_下期実績_2026.03" / "【経営企画部】各社業績対比フォーマット（2026.03下期実績 vs 2025.03下期実績）_ver.5.xlsx",
        "xlsx_pp":   PROJ / "17_下期実績_2024.03" / "【経営企画部】各社業績対比フォーマット（2024.03下期実績 vs 2023.03下期実績）_ver.3.xlsx",
        "curr_suffix_3": "-H2", "curr_suffix_8": "-H2",
        "prev_suffix_3": "-H2", "prev_suffix_8": "-H2",
        "pp_suffix_3":   "-H2", "pp_suffix_8":   "-H2",
    },
}

# 各社の列レイアウト
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

# 通期計画（既存）— 末尾「次期通期業績予想」表用にforecast_annualで保持
FORECAST = {
    "yamada":       {"Revenue": 1780000, "OperatingIncome": 51500, "OrdinaryIncome": 52600, "NetIncome": 27800, "GrossProfit": 503100},
    "ks":           {"Revenue": 785000,  "OperatingIncome": 30500, "OrdinaryIncome": 33500, "NetIncome": 20000, "GrossProfit": 219500},
    "edion":        {"Revenue": 816000,  "OperatingIncome": 27000, "OrdinaryIncome": 27000, "NetIncome": 15700, "GrossProfit": 235800},
    "joshin":       {"Revenue": 438000,  "OperatingIncome": 6000,  "OrdinaryIncome": 5500,  "NetIncome": 3500,  "GrossProfit": 114500},
    "nojima":       {"Revenue": 1000000, "OperatingIncome": 59000, "OrdinaryIncome": 76000, "NetIncome": 48000, "GrossProfit": None},
    "bic":          {"Revenue": 1013000, "OperatingIncome": 30500, "OrdinaryIncome": 31500, "NetIncome": 17500, "GrossProfit": 271484},
    "kojima":       {"Revenue": 294000,  "OperatingIncome": 7600,  "OrdinaryIncome": 7900,  "NetIncome": 4900,  "GrossProfit": 79968},
    "yamada_denki": {"Revenue": 1407400, "OperatingIncome": 34500, "OrdinaryIncome": 37100, "NetIncome": None,  "GrossProfit": 411500},
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


def build_one(period_key, conf):
    print(f"\n========== {period_key} ({conf['label']}) ==========")
    if not conf["xlsx_curr"].exists():
        print(f"  ERR xlsx_curr not found: {conf['xlsx_curr'].name}")
        return
    if not conf["xlsx_pp"].exists():
        print(f"  ERR xlsx_pp not found: {conf['xlsx_pp'].name}")
        return
    wb = load_workbook(conf["xlsx_curr"], data_only=False)
    ws = wb.active
    wb_pp = load_workbook(conf["xlsx_pp"], data_only=True)
    ws_pp = wb_pp.active

    companies_out = []
    for co in COMPANIES_MAIN:
        is_8 = (co["fy_end"] == "08")
        # 3月決算社: FY2026, 8月決算社: FY2025
        cp_base = "FY2025" if is_8 else "FY2026"
        pp_base = "FY2024" if is_8 else "FY2025"
        ppp_base = "FY2023" if is_8 else "FY2024"
        suf = conf["curr_suffix_8"] if is_8 else conf["curr_suffix_3"]
        sufp = conf["prev_suffix_8"] if is_8 else conf["prev_suffix_3"]
        sufpp = conf["pp_suffix_8"] if is_8 else conf["pp_suffix_3"]
        current_period = cp_base + suf
        previous_period = pp_base + sufp
        prev_previous_period = ppp_base + sufpp
        forecast_period = "FY2026" if is_8 else "FY2027"

        metrics = {}
        for row, key in ROW_TO_METRIC.items():
            cur = to_num(ws.cell(row=row, column=co["col_curr"]).value)
            prev = to_num(ws.cell(row=row, column=co["col_prev"]).value)
            prev_prev = to_num(ws_pp.cell(row=row, column=co["col_curr"]).value)
            fc_annual = FORECAST.get(co["key"], {}).get(key)
            metrics[key] = {
                "prev_previous": prev_prev,
                "current": cur,
                "previous": prev,
                "forecast": None,           # 四半期では推移グラフ4点目を描画しない
                "forecast_annual": fc_annual,
                "unit": "百万円",
            }

        # YoY
        yoy = {}
        for key in ("Revenue", "OperatingIncome", "OrdinaryIncome", "NetIncome"):
            m = metrics.get(key, {})
            cur = m.get("current"); prev = m.get("previous"); pp = m.get("prev_previous")
            yoy[key] = {
                "current_yoy_pct": round((cur / prev - 1) * 100, 2) if cur and prev else None,
                "previous_yoy_pct": round((prev / pp - 1) * 100, 2) if prev and pp else None,
                "forecast_yoy_pct": None,
            }

        # ratios
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
            "operating_margin_pct":             round(op / rev * 100, 2) if op and rev else None,
            "ordinary_margin_pct":              round(ord_c / rev * 100, 2) if ord_c and rev else None,
            "ordinary_margin_pct_previous":     round(ord_p / rev_p * 100, 2) if ord_p and rev_p else None,
            "ordinary_margin_pct_prev_previous":round(ord_pp / rev_pp * 100, 2) if ord_pp and rev_pp else None,
            "equity_ratio_pct":                 round(te / ta * 100, 2) if te and ta else None,
            "gross_margin_pct":                 round(gp_c / rev * 100, 2) if gp_c and rev else None,
            "gross_margin_pct_previous":        round(gp_p / rev_p * 100, 2) if gp_p and rev_p else None,
            "gross_margin_pct_prev_previous":   round(gp_pp / rev_pp * 100, 2) if gp_pp and rev_pp else None,
        }

        # trend_ratios (推移グラフ描画ロジックでparseされる) 四半期では推移グラフは通期のみ表示なので
        # 値はあっても無くてもダッシュボード上は表示しない。
        # ただ、エラー回避のため空dictを返す。
        trend_ratios = {"roe": {}, "roic": {}, "asset_turnover": {}, "inventory_turnover": {}}
        for key in trend_ratios.keys():
            trend_ratios[key] = {"prev_previous": None, "previous": None, "current": None, "forecast": None}

        if co.get("is_segment"):
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
        "source": f"各社決算短信・決算説明会資料 ({conf['label']})",
        "period_type": f"quarterly_{period_key}",
        "companies": companies_out,
    }
    out_path = OUT_DIR / f"companies_{period_key}.json"
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(output, f, ensure_ascii=False, indent=2)
    print(f"  → {out_path.name}")
    for c in companies_out:
        rev = c["metrics"]["Revenue"]["current"]
        op = c["metrics"]["OperatingIncome"]["current"]
        om = c["ratios"]["ordinary_margin_pct"]
        rev_s = f"{rev:,.0f}" if isinstance(rev,(int,float)) else "—"
        op_s = f"{op:,.0f}" if isinstance(op,(int,float)) else "—"
        om_s = f"{om}%" if om is not None else "—"
        print(f"    {c['label']:20s} ({c['current_period']:14s}): 売上 {rev_s:>11s} / 営利 {op_s:>8s} / 経常利益率 {om_s}")


def main():
    args = sys.argv[1:]
    if not args:
        args = list(PERIODS.keys())
    for pk in args:
        if pk not in PERIODS:
            print(f"  unknown period: {pk}")
            continue
        build_one(pk, PERIODS[pk])


if __name__ == "__main__":
    main()
