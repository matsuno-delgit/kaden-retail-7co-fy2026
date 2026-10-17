"""次期通期の会社計画（「次期通期業績予想」表・推移グラフの計画点）の一元管理。

対象は 2027年3月期（8月決算2社は四半期番号を合わせた通常版で 2026年8月期）。
期初計画を INITIAL に、その後の修正を REVISIONS に持ち、plan_for() は修正があれば修正後を返す
（原則として最新の会社計画を表示する。2026-10-17 方針）。修正した会社は表の社名下に「○○修正」の印を出す。

単位: 百万円。GrossProfit は計画額の開示が無い会社は説明会資料の「売上総利益 売上比」（小数1桁）× 売上高。
"""

# ---- 期初計画 ----
# GrossProfit: 各社の2026/3期通期決算説明会資料の来期計画ページから
#   ヤマダHD: D8=503,100  デンキセグ: F8=411,500  ケーズ: H8=219,500  エディオン: J8=235,800
#   上新: 説明会P19=1,145億円  コジマ: 説明会P15 売上比27.2%→79,968  ビック連結: 説明会P18 売上比26.8%→271,484
#   ノジマ: 説明会資料に売上総利益計画記載なし→None
# ビック単体: 2025年8月期 決算説明会資料の単体業績予想（02_通期計画 T列）。売上総利益の計画は非開示
# デンキセグメントは ヤマダ説明会資料202603 P36 連結計画 から
INITIAL = {
    "yamada":       {"Revenue": 1780000, "OperatingIncome": 51500, "OrdinaryIncome": 52600, "NetIncome": 27800, "GrossProfit": 503100},
    "ks":           {"Revenue": 785000,  "OperatingIncome": 30500, "OrdinaryIncome": 33500, "NetIncome": 20000, "GrossProfit": 219500},
    "edion":        {"Revenue": 816000,  "OperatingIncome": 27000, "OrdinaryIncome": 27000, "NetIncome": 15700, "GrossProfit": 235800},
    "joshin":       {"Revenue": 438000,  "OperatingIncome": 6000,  "OrdinaryIncome": 5500,  "NetIncome": 3500,  "GrossProfit": 114500},
    "nojima":       {"Revenue": 1000000, "OperatingIncome": 59000, "OrdinaryIncome": 76000, "NetIncome": 48000, "GrossProfit": None},
    "bic":          {"Revenue": 1013000, "OperatingIncome": 30500, "OrdinaryIncome": 31500, "NetIncome": 17500, "GrossProfit": 271484},
    "kojima":       {"Revenue": 294000,  "OperatingIncome": 7600,  "OrdinaryIncome": 7900,  "NetIncome": 4900,  "GrossProfit": 79968},
    "bic_s":        {"Revenue": 481000,  "OperatingIncome": 10000, "OrdinaryIncome": 14300, "NetIncome": 11100, "GrossProfit": None},
    "yamada_denki": {"Revenue": 1407400, "OperatingIncome": 34500, "OrdinaryIncome": 37100, "NetIncome": None,  "GrossProfit": 411500},
}

# ---- 期初計画からの修正（通期） ----
# 2027年3月期 第1四半期決算短信（2026年7〜8月）で確認: ヤマダ・ケーズ・エディオンは修正なし。
#   上新は「修正有」だが第2四半期累計のみで、通期は期初どおり。
REVISIONS = {
    # 2027年3月期 第1四半期決算短信 p2（2026/7/30）。売上高のみ 1,000,000→1,030,000、利益は据え置き
    "nojima": {"date": "2026-07-30",
               "plan": {"Revenue": 1030000, "OperatingIncome": 59000, "OrdinaryIncome": 76000, "NetIncome": 48000, "GrossProfit": None}},
    # 2026年8月期 第2四半期決算短信 表紙（2026/4/10）＋補足説明資料 p12（売上総利益 売上比27.1%）
    "bic":    {"date": "2026-04-10",
               "plan": {"Revenue": 1022000, "OperatingIncome": 34400, "OrdinaryIncome": 35700, "NetIncome": 18400, "GrossProfit": 276962}},
    # 2026年8月期 第2四半期決算 補足説明資料 p13 単体業績予想 修正予想（売上総利益 売上比25.1%）
    "bic_s":  {"date": "2026-04-10",
               "plan": {"Revenue": 497000, "OperatingIncome": 11900, "OrdinaryIncome": 16600, "NetIncome": 11700, "GrossProfit": 124747}},
    # 2026年8月期 第2四半期決算短信 表紙（2026/4/9）＋補足説明資料 p15（売上総利益 売上比27.4%）。売上高は据え置き
    "kojima": {"date": "2026-04-09",
               "plan": {"Revenue": 294000, "OperatingIncome": 8200, "OrdinaryIncome": 8500, "NetIncome": 5300, "GrossProfit": 80556}},
}


def plan_for(key):
    """表示に使う計画（修正があれば修正後）。"""
    if key in REVISIONS:
        return dict(REVISIONS[key]["plan"])
    return dict(INITIAL.get(key, {}))


def _fmt(v):
    return "—" if v is None else f"{v:,.0f}"


def revision_note(date, initial, revised, prefix=""):
    """表の印（badge）とツールチップ用の説明（期初→修正）。"""
    y, m, d = (int(x) for x in date.split("-"))
    names = (("Revenue", "売上高"), ("OperatingIncome", "営業利益"), ("NetIncome", "純利益"))
    parts = [f"{n} {_fmt(initial.get(k))}→{_fmt(revised.get(k))}" for k, n in names
             if initial.get(k) is not None and initial.get(k) != revised.get(k)]
    return {
        "date": date,
        "badge": f"{prefix}{m}月計画修正",
        "text": (f"{y}年{m}月{d}日に{prefix or '通期'}計画を修正"
                 f"（{'合算値の' if prefix else ''}期初計画→修正後: {'、'.join(parts) or '経常利益のみ'}）"),
    }


def revision_for(key):
    """通期計画の修正情報（無ければ None）。"""
    if key not in REVISIONS:
        return None
    r = REVISIONS[key]
    return revision_note(r["date"], INITIAL[key], r["plan"])
