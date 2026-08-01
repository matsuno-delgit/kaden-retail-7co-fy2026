# -*- coding: utf-8 -*-
"""consolidate系スクリプト共通ユーティリティ。

- find_latest_xlsx: フォルダ内の `_ver.N.xlsx` の最大Nを自動検出。
  Excelを更新するたびにconsolidate側のver番号を書き換える運用を廃止するためのもの。
- data_sheet: 「PL・BSデータ」シートを名前で特定。wb.active は Excel で最後に
  選択していたシートに依存して壊れるため使わない。
"""
import re
from pathlib import Path

_VER_RE = re.compile(r"ver\.(\d+)")


def find_latest_xlsx(folder, keyword=""):
    """folder直下で、名前にkeywordを含み `ver.N` を持つxlsxのうちNが最大のものを返す。

    ~$で始まるExcelロック残骸は除外。該当なしはFileNotFoundError。
    """
    folder = Path(folder)
    cands = [
        p for p in folder.glob("*.xlsx")
        if not p.name.startswith("~") and keyword in p.name and _VER_RE.search(p.name)
    ]
    if not cands:
        raise FileNotFoundError(f"versioned xlsx not found in {folder} (keyword={keyword!r})")
    return max(cands, key=lambda p: int(_VER_RE.search(p.name).group(1)))


def data_sheet(wb):
    """「PL・BSデータ」で始まる名前のシートを返す（無ければ先頭シート）。"""
    for sn in wb.sheetnames:
        if sn.startswith("PL・BSデータ"):
            return wb[sn]
    return wb[wb.sheetnames[0]]


# --- 直近四半期（LTM）回転率ブロック ---------------------------------------
# Excel側 add_ltm_turnover.py が全ファイル共通の行番号で書き込む。
ROW_LTM_REVENUE = 117      # 直近12ヶ月売上高（LTM）
ROW_LTM_AVG_TA = 118       # 平均総資産（2期平均）
ROW_LTM_AVG_INV = 119      # 平均商品（棚卸資産）（2期平均）
ROW_LTM_OP = 122           # 直近12ヶ月営業利益（LTM）
ROW_LTM_NI = 123           # 直近12ヶ月当期純利益（LTM）
ROW_LTM_AVG_EQ = 124       # 平均自己資本（2期平均）
ROW_LTM_AVG_DEBT = 125     # 平均有利子負債（2期平均）
# R120/R121/R98/R99 は数式セルのため openpyxl では値を取得できない。上記の実数行から再計算する。
EFFECTIVE_TAX_RATE = 0.35  # ROIC算定に使う実効税率（Excel R113 と同じ固定値）


def _num(v):
    if v is None or (isinstance(v, str) and v.startswith("=")):
        return None
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def _div(a, b):
    if a is None or not b:
        return None
    return round(a / b, 3)


_LTM_EMPTY = {"revenue": None, "avg_total_assets": None, "avg_inventory": None,
              "asset_turnover": None, "inventory_turnover": None,
              "operating_income": None, "net_income": None,
              "avg_equity": None, "avg_debt": None,
              "roe_pct": None, "roic_pct": None}


def read_ltm(ws, col):
    """1列分の直近四半期(LTM)指標を返す。値が無ければ各キー None。

    すべて「直近12ヶ月（直近4四半期）の損益 ÷ 2期平均残高」で統一している。
      asset_turnover     = 直近12ヶ月売上高 ÷ 平均総資産（当期末・前年同期末の平均）
      inventory_turnover = 直近12ヶ月売上高 ÷ 平均棚卸資産（同上）
      roe_pct            = 直近12ヶ月当期純利益 ÷ 平均自己資本 ×100
      roic_pct           = 直近12ヶ月営業利益 ×(1−35%) ÷（平均有利子負債＋平均自己資本）×100
    """
    if ws is None or not col:
        return dict(_LTM_EMPTY)
    g = lambda r: _num(ws.cell(row=r, column=col).value)
    rev, ta, inv = g(ROW_LTM_REVENUE), g(ROW_LTM_AVG_TA), g(ROW_LTM_AVG_INV)
    op, ni = g(ROW_LTM_OP), g(ROW_LTM_NI)
    eq, debt = g(ROW_LTM_AVG_EQ), g(ROW_LTM_AVG_DEBT)
    # 投下資本は「有利子負債＋自己資本」。自己資本が未取得の期に負債だけを分母にすると
    # ROICが過大に出るため、自己資本が揃っている場合のみ算定する。
    ic = (eq + (debt or 0)) if eq else None
    roe = _div(ni, eq)
    roic = _div(op * (1 - EFFECTIVE_TAX_RATE), ic) if op is not None and ic else None
    return {
        "revenue": rev,
        "avg_total_assets": ta,
        "avg_inventory": inv,
        "asset_turnover": _div(rev, ta),
        "inventory_turnover": _div(rev, inv),
        "operating_income": op,
        "net_income": ni,
        "avg_equity": eq,
        "avg_debt": debt,
        "roe_pct": round(roe * 100, 2) if roe is not None else None,
        "roic_pct": round(roic * 100, 2) if roic is not None else None,
    }


def build_ltm(ws_curr, ws_pp, co, pp_col_key="col_curr"):
    """当期/前期/前々期のLTM指標をまとめて返す。

    当期・前期は当期Excelの当期列/前期列から取る。
    前々期は前々期Excelから取るが、参照する列は系列によって異なるため
    pp_col_key で指定する（通常系列は当期列、組替系列は前期列）。
    """
    return {
        "current": read_ltm(ws_curr, co["col_curr"]),
        "previous": read_ltm(ws_curr, co["col_prev"]),
        "prev_previous": read_ltm(ws_pp, co[pp_col_key]),
    }


def ltm_trend(ltm, key):
    """build_ltm() の結果を推移グラフ用の {前々期/前期/当期/次期予想} 形式に変換する。

    次期予想はBS非開示のため常に None。
    """
    return {
        "prev_previous": ltm["prev_previous"][key],
        "previous": ltm["previous"][key],
        "current": ltm["current"][key],
        "forecast": None,
    }


def margin(num, den):
    """利益率(%)。分母が0/Noneなら None。"""
    if num is None or not den:
        return None
    return round(num / den * 100, 2)


# ---------------------------------------------------------------------------
# 期間ラベル
# ---------------------------------------------------------------------------
SUFFIX_LABEL = {
    "": "", "-Q1": "1Q", "-Q2": "2Q単独", "-H1": "上期",
    "-Q3": "3Q単独", "-Q3CUM": "3Q累計", "-Q4": "4Q単独", "-H2": "下期",
}

# 組替期間キー → 8月決算社における実際の期間。
#   (FYオフセット, サフィックス)  … 自社の四半期と1対1で対応する場合
#   ("SPAN", テンプレ)            … 自社の会計年度をまたぎ1対1対応しない場合
# 組替FY(3月決算社基準)を base とすると、組替期間の暦月は
#   通期 base-1年3月〜base年2月 / 1Q 3-5月 / 2Q 6-8月 / 3Q 9-11月 / 4Q 12-2月。
KUMIKAE_08_PERIOD = {
    "fy":    ("SPAN", "{p}年3月〜{c}年2月"),
    "q1":    (-1, "-Q3"),      # base-1年3〜5月   = 8月期3Q
    "q2":    (-1, "-Q4"),      # base-1年6〜8月   = 8月期4Q
    "h1":    (-1, "-H2"),      # base-1年3〜8月   = 8月期下期
    "q3":    (0,  "-Q1"),      # base-1年9〜11月  = 翌8月期1Q
    "q3cum": ("SPAN", "{p}年3月〜11月"),
    "q4":    (0,  "-Q2"),      # base-1年12〜base年2月 = 翌8月期2Q
    "h2":    (0,  "-H1"),      # base-1年9〜base年2月  = 翌8月期上期
}


def period_label(code, fy_end):
    """期間コード ("FY2026-Q3CUM" 等) を日本語ラベルに変換する。"""
    if not code:
        return ""
    fy, _, suffix = code.partition("-")
    base = f"{fy.replace('FY', '')}年{fy_end}月期"
    lab = SUFFIX_LABEL.get(f"-{suffix}" if suffix else "", suffix)
    return f"{base} {lab}".strip()


def kumikae_labels_08(period_key, base_fy):
    """組替版で8月決算社の (当期, 前期, 前々期) 実期間ラベルを返す。

    base_fy は3月決算社基準の組替FY(int)。組替FY2026なら2025/3〜2026/2。
    """
    spec = KUMIKAE_08_PERIOD[period_key]
    out = []
    for shift in (0, -1, -2):
        b = base_fy + shift
        if spec[0] == "SPAN":
            out.append(spec[1].format(p=b - 1, c=b))
        else:
            off, suf = spec
            out.append(period_label(f"FY{b + off}{suf}", "08"))
    return tuple(out)
