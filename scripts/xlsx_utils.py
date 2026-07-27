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
# R120/R121 は数式セルのため openpyxl では値を取得できない。117〜119から再計算する。


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


def read_ltm(ws, col):
    """1列分の直近四半期(LTM)指標を返す。値が無ければ各キー None。

    asset_turnover     = 直近12ヶ月売上高 ÷ 平均総資産（当期末・前年同期末の2期平均）
    inventory_turnover = 直近12ヶ月売上高 ÷ 平均棚卸資産（同上）
    """
    if ws is None or not col:
        return {"revenue": None, "avg_total_assets": None, "avg_inventory": None,
                "asset_turnover": None, "inventory_turnover": None}
    rev = _num(ws.cell(row=ROW_LTM_REVENUE, column=col).value)
    ta = _num(ws.cell(row=ROW_LTM_AVG_TA, column=col).value)
    inv = _num(ws.cell(row=ROW_LTM_AVG_INV, column=col).value)
    return {
        "revenue": rev,
        "avg_total_assets": ta,
        "avg_inventory": inv,
        "asset_turnover": _div(rev, ta),
        "inventory_turnover": _div(rev, inv),
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


def margin(num, den):
    """利益率(%)。分母が0/Noneなら None。"""
    if num is None or not den:
        return None
    return round(num / den * 100, 2)
