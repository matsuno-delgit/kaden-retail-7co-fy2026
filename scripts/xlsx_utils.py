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
