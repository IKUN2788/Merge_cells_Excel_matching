from __future__ import annotations

import os
from dataclasses import dataclass
from typing import Dict, Iterable, List, Optional, Sequence, Tuple, Union

from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet


@dataclass(frozen=True)
class SourceConfig:
    xlsx_path: str
    sheet_name: str
    key_columns: Tuple[str, ...]
    value_column: str
    accumulate: bool = False


@dataclass(frozen=True)
class TargetConfig:
    xlsx_path: str
    sheet_name: str
    key_columns: Tuple[str, ...]
    write_to_column: Optional[str] = None


def list_sheets(xlsx_path: str) -> List[str]:
    wb = load_workbook(xlsx_path, read_only=True, data_only=True)
    try:
        return list(wb.sheetnames)
    finally:
        wb.close()


def list_columns(xlsx_path: str, sheet_name: str) -> List[str]:
    wb = load_workbook(xlsx_path, read_only=True, data_only=True)
    try:
        ws = wb[sheet_name]
        headers: List[str] = []
        for idx, cell in enumerate(ws[1], start=1):
            v = cell.value
            name = str(v).strip() if v is not None else ""
            if name == "":
                name = f"列{idx}"
            headers.append(name)
        return headers
    finally:
        wb.close()


def build_source_mapping(src_cfg: SourceConfig) -> Dict[str, Union[str, int, float]]:
    wb = load_workbook(src_cfg.xlsx_path, read_only=True, data_only=True)
    try:
        ws = wb[src_cfg.sheet_name]
        header_to_col = _header_to_col(ws)

        key_col_indices = _resolve_columns(header_to_col, src_cfg.key_columns)
        value_col_index = _resolve_column(header_to_col, src_cfg.value_column)
        merged_lookup = _build_merged_lookup(ws)

        mapping: Dict[str, Union[str, int, float]] = {}
        for row in range(2, ws.max_row + 1):
            key_parts = []
            bad_key = False
            for c in key_col_indices:
                v = _get_cell_value(ws, merged_lookup, row, c)
                s = _normalize_text(v)
                if s == "":
                    bad_key = True
                    break
                key_parts.append(s)
            if bad_key:
                continue

            key = "_".join(key_parts)
            val_raw = _get_cell_value(ws, merged_lookup, row, value_col_index)
            val = _normalize_value(val_raw)
            if val is None:
                continue

            if not src_cfg.accumulate:
                mapping[key] = val
                continue

            if key not in mapping:
                mapping[key] = val
                continue

            mapping[key] = _accumulate(mapping[key], val)

        return mapping
    finally:
        wb.close()


def apply_mapping_to_target(
    *,
    src_cfg: SourceConfig,
    mapping: Dict[str, Union[str, int, float]],
    tgt_cfg: TargetConfig,
    output_path: str,
    output_header: Optional[str] = None,
) -> Tuple[int, int]:
    wb = load_workbook(tgt_cfg.xlsx_path, read_only=False, data_only=True)
    try:
        ws = wb[tgt_cfg.sheet_name]
        header_to_col = _header_to_col(ws)

        key_col_indices = _resolve_columns(header_to_col, tgt_cfg.key_columns)

        write_col_index: int
        if tgt_cfg.write_to_column:
            write_col_index = _resolve_column(header_to_col, tgt_cfg.write_to_column)
        else:
            write_col_index = ws.max_column + 1
            header = (output_header or "").strip() or f"匹配_{src_cfg.value_column}"
            ws.cell(row=1, column=write_col_index).value = header

        merged_lookup = _build_merged_lookup(ws)

        matched = 0
        total = 0
        for row in range(2, ws.max_row + 1):
            key_parts = []
            bad_key = False
            for c in key_col_indices:
                v = _get_cell_value(ws, merged_lookup, row, c)
                s = _normalize_text(v)
                if s == "":
                    bad_key = True
                    break
                key_parts.append(s)
            if bad_key:
                continue

            total += 1
            key = "_".join(key_parts)
            if key not in mapping:
                continue

            ws.cell(row=row, column=write_col_index).value = mapping[key]
            matched += 1

        folder = os.path.dirname(os.path.abspath(output_path))
        if folder and not os.path.exists(folder):
            os.makedirs(folder, exist_ok=True)
        wb.save(output_path)
        return matched, total
    finally:
        wb.close()


def _header_to_col(ws: Worksheet) -> Dict[str, int]:
    header_to_col: Dict[str, int] = {}
    for idx, cell in enumerate(ws[1], start=1):
        v = cell.value
        name = str(v).strip() if v is not None else ""
        if name == "":
            name = f"列{idx}"
        if name not in header_to_col:
            header_to_col[name] = idx
    return header_to_col


def _resolve_column(header_to_col: Dict[str, int], col_name: str) -> int:
    name = col_name.strip()
    if name in header_to_col:
        return header_to_col[name]
    raise KeyError(f"找不到列：{col_name}")


def _resolve_columns(header_to_col: Dict[str, int], cols: Sequence[str]) -> List[int]:
    return [_resolve_column(header_to_col, c) for c in cols]


def _build_merged_lookup(ws: Worksheet) -> Dict[Tuple[int, int], Tuple[int, int]]:
    lookup: Dict[Tuple[int, int], Tuple[int, int]] = {}
    for r in ws.merged_cells.ranges:
        min_row = r.min_row
        min_col = r.min_col
        for row in range(r.min_row, r.max_row + 1):
            for col in range(r.min_col, r.max_col + 1):
                lookup[(row, col)] = (min_row, min_col)
    return lookup


def _get_cell_value(
    ws: Worksheet, merged_lookup: Dict[Tuple[int, int], Tuple[int, int]], row: int, col: int
):
    cell = ws.cell(row=row, column=col)
    if cell.value is not None:
        return cell.value
    anchor = merged_lookup.get((row, col))
    if not anchor:
        return None
    ar, ac = anchor
    return ws.cell(row=ar, column=ac).value


def _normalize_text(v) -> str:
    if v is None:
        return ""
    s = str(v).strip()
    return s


def _normalize_value(v) -> Optional[Union[str, int, float]]:
    if v is None:
        return None
    if isinstance(v, bool):
        return str(v)
    if isinstance(v, (int, float)):
        return v
    s = str(v).strip()
    if s == "":
        return None
    return s


def _accumulate(old: Union[str, int, float], new: Union[str, int, float]) -> Union[str, int, float]:
    if isinstance(old, (int, float)) and isinstance(new, (int, float)):
        if isinstance(old, int) and isinstance(new, int):
            return old + new
        return float(old) + float(new)

    old_s = _normalize_text(old)
    new_s = _normalize_text(new)
    parts: List[str] = []
    seen = set()
    for p in _split_parts(old_s):
        if p not in seen:
            seen.add(p)
            parts.append(p)
    for p in _split_parts(new_s):
        if p not in seen:
            seen.add(p)
            parts.append(p)
    return ";".join(parts)


def _split_parts(s: str) -> Iterable[str]:
    for p in s.split(";"):
        p = p.strip()
        if p != "":
            yield p
