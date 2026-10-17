"""次期の会社計画（上期・下期・通期）の一元管理と data/plans_fy2027.json の生成。

対象は 2027年3月期（8月決算2社は四半期番号を合わせた通常版で 2026年8月期）。
期初計画と、その後の修正を分けて持ち、表示には原則として最新（修正後）の計画を使う（2026-10-17 方針）。
修正した計画は表で「○月計画修正」の印を出し、期初→修正後の差を注に書く。

  通期: INITIAL / REVISIONS
  上期: H1_INITIAL / H1_REVISIONS
  下期: H2_INITIAL / H2_REVISIONS（会社が下期計画を開示している分）。開示が無い会社、
        または上期・通期だけが修正されて開示済みの下期が古くなった会社は「最新の通期−最新の上期」で算出（印「算出」）。
        ※上期の実績確定後に通期を修正した場合（下期計画＝修正通期−上期実績）は、資料の下期修正予想を
          H2_REVISIONS に入れる（例: ビック・コジマの2026年4月修正）。
  組替版の8月決算2社（暦月を3月決算に合わせる）:
        組替上期(3〜8月)＝8月期の下期計画、組替下期(9〜2月)＝翌8月期の上期計画（NEXT_H1_08）、組替通期＝その合計。

単位: 百万円。GrossProfit は計画額の開示が無い会社は説明会資料の「売上総利益 売上比」（小数1桁）× 売上高。

usage: .venv/Scripts/python.exe scripts/forecast_plans.py   → data/plans_fy2027.json
"""
import json
from datetime import date
from pathlib import Path

ITEMS = ("Revenue", "GrossProfit", "OperatingIncome", "OrdinaryIncome", "NetIncome")
KEYS = ("yamada", "yamada_denki", "ks", "edion", "joshin", "nojima", "bic", "bic_s", "kojima")
KEYS_08 = ("bic", "bic_s", "kojima")


def P(rev, op, ord_, ni, gp=None, gm=None):
    """計画1件。gp が無く gm（売上総利益の売上比%）があれば売上高×gm で算出。"""
    if gp is None and gm is not None:
        gp = round(rev * gm / 100)
    return {"Revenue": rev, "GrossProfit": gp, "OperatingIncome": op, "OrdinaryIncome": ord_, "NetIncome": ni}


# ======== 通期 ========
# 期初計画。GrossProfit: 各社の2026/3期通期決算説明会資料の来期計画ページから
#   ヤマダHD: P33=503,100  デンキセグ: P36=411,500  ケーズ: P20=219,500  エディオン: P16=235,800
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

# 期初計画からの修正（通期）
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

# ======== 上期 ========
# 期初計画（決算短信表紙の「第2四半期（累計）」＋説明会資料の売上総利益）
#   ヤマダ: 2026/3期 通期決算説明会資料 P34【上期・下期】  ケーズ: 説明会資料 P20 27/3上期計画
#   エディオン: 4Q決算説明会 P16  上新: 決算説明資料 P13  ノジマ: 決算短信表紙（売上総利益の計画なし）
#   ビック連結: 2025年8月期 決算説明資料 P18（売上比）  ビック単体: 同 P15  コジマ: 同 P16
#   デンキセグメントは上期・下期の計画を開示していない
H1_INITIAL = {
    "yamada": P(862000, 27500, 28400, 15500, gp=247700),
    "ks":     P(398000, 17500, 18800, 13100, gp=111500),
    "edion":  P(400000, 14700, 14700, 9500, gp=119800),
    "joshin": P(211000, 2500, 2300, 1900, gp=55900),
    "nojima": P(470000, 28000, 44000, 27000),
    "bic":    P(497500, 14440, 14900, 9610, gm=26.7),
    "bic_s":  P(238500, 4200, 8140, 7950, gm=24.7),
    "kojima": P(143000, 3500, 3700, 2500, gm=27.2),
}
H1_REVISIONS = {
    # 2027年3月期 第1四半期決算説明資料 p9 上期予想（今回予想）。通期は据え置き
    "joshin": {"date": "2026-08-04", "plan": P(218000, 4500, 4300, 2800, gp=56400)},
    # 2027年3月期 第1四半期決算短信 p2 第2四半期(累計)。売上高のみ 470,000→490,000
    "nojima": {"date": "2026-07-30", "plan": P(490000, 28000, 44000, 27000)},
}

# ======== 下期 ========
# 会社が開示した期初の下期計画（ケーズ・ノジマは開示なし→算出）
#   ヤマダ: P34  エディオン: P16  上新: P13  ビック連結: P18  ビック単体: P15  コジマ: P16
H2_INITIAL = {
    "yamada": P(918000, 24000, 24200, 12300, gp=255400),
    "edion":  P(416000, 12300, 12300, 6200, gp=116000),
    "joshin": P(227000, 3500, 3200, 1600, gp=58600),
    "bic":    P(515500, 16060, 16600, 7890, gm=27.0),
    "bic_s":  P(242500, 5800, 6160, 3150, gm=25.8),
    "kojima": P(151000, 4100, 4200, 2400, gm=27.2),
}
# 下期計画の修正（2026年8月期 第2四半期決算 補足説明資料「下期 修正予想」: ビック連結 p12 / ビック単体 p13 / コジマ p15）
H2_REVISIONS = {
    "bic":    {"date": "2026-04-10", "plan": P(513570, 15672, 16278, 7301, gm=27.6)},
    "bic_s":  {"date": "2026-04-10", "plan": P(249238, 5703, 5909, 3273, gm=25.8)},
    "kojima": {"date": "2026-04-09", "plan": P(150062, 4125, 4281, 2480, gm=27.6)},
}

# ======== 組替版の8月決算2社: 翌8月期（2027年8月期）の上期計画 ========
# 2026年8月期決算 補足説明資料の「2027年8月期 上期予想」: ビック連結 p14 / ビック単体 p15 / コジマ p15
NEXT_H1_08 = {
    "bic":    P(536500, 18100, 18500, 10500, gm=27.0),
    "bic_s":  P(266400, 6100, 11550, 9300, gm=24.7),
    "kojima": P(148000, 4100, 4300, 2900, gm=27.4),
}


def _fmt(v):
    return "—" if v is None else f"{v:,.0f}"


def _ymd(d):
    y, m, dd = (int(x) for x in d.split("-"))
    return y, m, dd


def revision_note(date_, initial, revised, what="通期", combined=False):
    """表の印（badge）と注の説明（期初→修正後）。"""
    y, m, d = _ymd(date_)
    names = (("Revenue", "売上高"), ("OperatingIncome", "営業利益"), ("NetIncome", "純利益"))
    parts = [f"{n} {_fmt(initial.get(k))}→{_fmt(revised.get(k))}" for k, n in names
             if initial.get(k) is not None and initial.get(k) != revised.get(k)]
    return {
        "date": date_,
        "badge": f"{m}月計画修正",
        "text": (f"{y}年{m}月{d}日に{what}計画を修正"
                 f"（{'合算値の' if combined else ''}期初計画→修正後: {'、'.join(parts) or '経常利益のみ'}）"),
    }


def plan_for(key):
    """通期の表示用計画（修正があれば修正後）。"""
    if key in REVISIONS:
        return dict(REVISIONS[key]["plan"])
    return dict(INITIAL.get(key, {}))


def revision_for(key):
    """通期計画の修正情報（無ければ None）。"""
    if key not in REVISIONS:
        return None
    r = REVISIONS[key]
    return revision_note(r["date"], INITIAL[key], r["plan"])


def _latest(initial, revisions, key):
    if key in revisions:
        return dict(revisions[key]["plan"])
    return dict(initial[key]) if key in initial else None


def _diff(a, b):
    return {k: (a[k] - b[k] if a.get(k) is not None and b.get(k) is not None else None) for k in ITEMS}


def _sum(a, b):
    return {k: (a[k] + b[k] if a.get(k) is not None and b.get(k) is not None else None) for k in ITEMS}


def h1_for(key):
    return _latest(H1_INITIAL, H1_REVISIONS, key)


def h2_for(key):
    """(下期計画, 算出の説明 or None)。"""
    if key in H2_REVISIONS:
        return dict(H2_REVISIONS[key]["plan"]), None
    h1 = h1_for(key)
    if key in H2_INITIAL and key not in H1_REVISIONS and key not in REVISIONS:
        return dict(H2_INITIAL[key]), None
    if h1 is None:
        return None, None
    why = []
    if key not in H2_INITIAL:
        why.append("下期計画の開示なし")
    if key in H1_REVISIONS:
        y, m, d = _ymd(H1_REVISIONS[key]["date"])
        why.append(f"{y}年{m}月{d}日の上期計画修正を反映")
    if key in REVISIONS:
        y, m, d = _ymd(REVISIONS[key]["date"])
        why.append(f"{y}年{m}月{d}日の通期計画修正を反映")
    return _diff(plan_for(key), h1), f"通期計画−上期計画で算出（{'、'.join(why)}）"


# ---- 組替版 ----
FORECAST_LABEL_08 = "2026/3〜2027/2（26/8期下期＋27/8期上期計画）"


def kumikae_forecast(key):
    """組替通期の計画。3月決算社は通常版と同じ。8月決算社は 8月期下期計画＋翌8月期上期計画。"""
    if key not in KEYS_08:
        return plan_for(key)
    return _sum(h2_for(key)[0], NEXT_H1_08[key])


def kumikae_revision(key):
    if key not in KEYS_08:
        return revision_for(key)
    if key not in H2_REVISIONS:
        return None
    r = H2_REVISIONS[key]
    return revision_note(r["date"], _sum(H2_INITIAL[key], NEXT_H1_08[key]), _sum(r["plan"], NEXT_H1_08[key]),
                         what="26/8期下期", combined=True)


def _row(label, plan, revision=None, derived=None):
    if plan is None:
        return {"label": label, **{k: None for k in ITEMS}, "revision": None, "derived": None, "undisclosed": True}
    return {"label": label, **{k: plan.get(k) for k in ITEMS}, "revision": revision, "derived": derived,
            "undisclosed": False}


def _half_rev(revs, initial, key, what):
    if key not in revs:
        return None
    return revision_note(revs[key]["date"], initial[key], revs[key]["plan"], what=what)


def build_rows(dataset):
    out = {}
    for key in KEYS:
        is08 = key in KEYS_08
        if dataset == "kumikae" and is08:
            h2, _ = h2_for(key)
            out[key] = {
                "h1": _row("26/8期 下期計画", h2, _half_rev(H2_REVISIONS, H2_INITIAL, key, "26/8期下期")),
                "h2": _row("27/8期 上期計画", NEXT_H1_08[key]),
                "fy": _row("2026/3〜2027/2", kumikae_forecast(key), kumikae_revision(key)),
            }
            continue
        label = "2026/8期" if is08 else "2027/3期"
        h2, derived = h2_for(key)
        out[key] = {
            "h1": _row(label, h1_for(key), _half_rev(H1_REVISIONS, H1_INITIAL, key, "上期")),
            "h2": _row(label, h2, _half_rev(H2_REVISIONS, H2_INITIAL, key, "下期"), derived),
            "fy": _row(label, plan_for(key), revision_for(key)),
        }
    return out


def main():
    out_path = Path(__file__).parent.parent / "data" / "plans_fy2027.json"
    data = {
        "schema_version": "1.0",
        "generated_at": date.today().isoformat(),
        "note": ("2027年3月期の会社計画（通常版の8月決算2社は2026年8月期、組替版は2026年3月〜2027年2月）。"
                 "原則として最新（修正後）の計画。単位: 百万円。生成: scripts/forecast_plans.py"),
        "datasets": {ds: build_rows(ds) for ds in ("normal", "kumikae")},
    }
    out_path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"wrote {out_path}")
    for ds, rows in data["datasets"].items():
        for key, r in rows.items():
            print(f"  {ds:8s} {key:13s} " + " / ".join(
                f"{p}:{_fmt(r[p]['Revenue'])}{'*' if r[p]['revision'] else ''}{'^' if r[p]['derived'] else ''}"
                for p in ("h1", "h2", "fy")))


if __name__ == "__main__":
    main()
