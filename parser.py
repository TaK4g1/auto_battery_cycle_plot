from __future__ import annotations

import math
import re
from dataclasses import dataclass
from typing import Any

import numpy as np
import pandas as pd

from config import DEFAULT_MODE_RULES


class DataParsingError(ValueError):
    """Raised when the Excel sheet cannot be converted into plottable point data."""


HEADER_ALIASES: dict[str, list[str]] = {
    "cycle": [
        "循环序号",
        "循环号",
        "循环",
        "cycle",
        "cycle index",
        "cyc no",
        "cycno",
    ],
    "mode": [
        "工作模式",
        "模式",
        "状态",
        "工步类型",
        "step type",
        "mode",
        "status",
    ],
    "voltage": [
        "电压/v",
        "电压",
        "voltage/v",
        "voltage",
        "volt",
    ],
    "specific_capacity": [
        "比容量/mah/g",
        "比容量",
        "specific capacity",
        "specificcapacity",
        "sp. capacity",
        "spcapacity",
    ],
    "capacity": [
        "容量/mah",
        "容量",
        "capacity/mah",
        "capacity",
    ],
    "current": [
        "电流/ma",
        "电流",
        "current/ma",
        "current",
    ],
    "step": [
        "工步序号",
        "工步",
        "step",
        "step index",
        "step no",
    ],
    "record": [
        "记录序号",
        "记录",
        "record",
        "record no",
        "point",
    ],
}

NORMALIZED_ALIASES = {
    canonical: [re.sub(r"[^0-9a-zA-Z\u4e00-\u9fff]+", "", alias).lower() for alias in aliases]
    for canonical, aliases in HEADER_ALIASES.items()
}


@dataclass
class HeaderDetectionResult:
    sheet_name: str
    header_row: int
    mapping: dict[str, str]
    score: int


@dataclass
class ParsedDataset:
    data: pd.DataFrame
    sheet_name: str
    header_row: int
    source_columns: dict[str, str]


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and math.isnan(value):
        return ""
    return re.sub(r"[^0-9a-zA-Z\u4e00-\u9fff]+", "", str(value)).lower()


def _match_header_name(header: str) -> tuple[str | None, int]:
    normalized = normalize_text(header)
    if not normalized:
        return None, 0

    best_name: str | None = None
    best_score = 0
    for canonical, aliases in NORMALIZED_ALIASES.items():
        for alias in aliases:
            if normalized == alias:
                score = 100 + len(alias)
            elif alias in normalized or normalized in alias:
                score = 10 + len(alias)
            else:
                continue
            if score > best_score:
                best_name = canonical
                best_score = score
    return best_name, best_score


def infer_column_mapping(headers: list[str]) -> dict[str, str]:
    matches: dict[str, tuple[str, int]] = {}
    for header in headers:
        canonical, score = _match_header_name(header)
        if not canonical:
            continue
        current = matches.get(canonical)
        if current is None or score > current[1]:
            matches[canonical] = (header, score)
    return {canonical: original for canonical, (original, _) in matches.items()}


def detect_header_row(sheet_name: str, preview_df: pd.DataFrame) -> HeaderDetectionResult | None:
    best_result: HeaderDetectionResult | None = None

    for row_index in range(len(preview_df.index)):
        row_values = preview_df.iloc[row_index].tolist()
        headers = ["" if pd.isna(item) else str(item).strip() for item in row_values]
        non_empty = sum(1 for item in headers if item)
        if non_empty < 3:
            continue

        mapping = infer_column_mapping(headers)
        score = 0
        if "voltage" in mapping:
            score += 8
        if "specific_capacity" in mapping:
            score += 8
        if "cycle" in mapping:
            score += 5
        if "mode" in mapping:
            score += 5
        if "record" in mapping:
            score += 2
        if "step" in mapping:
            score += 2

        if score < 16:
            continue

        candidate = HeaderDetectionResult(
            sheet_name=sheet_name,
            header_row=row_index,
            mapping=mapping,
            score=score,
        )
        if best_result is None or candidate.score > best_result.score:
            best_result = candidate

    return best_result


def build_mode_lookup(mode_overrides: dict[str, str] | None = None) -> dict[str, str]:
    lookup: dict[str, str] = {}
    for canonical, aliases in DEFAULT_MODE_RULES.items():
        for alias in aliases:
            lookup[normalize_text(alias)] = canonical

    if mode_overrides:
        for alias, canonical in mode_overrides.items():
            lookup[normalize_text(alias)] = canonical.lower()
    return lookup


def classify_mode(raw_value: Any, mode_lookup: dict[str, str]) -> str:
    normalized = normalize_text(raw_value)
    if not normalized:
        return "unknown"
    if normalized in mode_lookup:
        return mode_lookup[normalized]

    for alias, canonical in mode_lookup.items():
        if alias and alias in normalized:
            return canonical
    return "unknown"


def _coerce_cycle_number(series: pd.Series) -> pd.Series:
    numeric = pd.to_numeric(series, errors="coerce")
    rounded = numeric.round()
    is_valid = numeric.notna() & np.isfinite(numeric) & ((numeric - rounded).abs() < 1e-6)
    result = pd.Series(pd.NA, index=series.index, dtype="Int64")
    result.loc[is_valid] = rounded.loc[is_valid].astype("Int64")
    return result


def standardize_columns(raw_df: pd.DataFrame, header_mapping: dict[str, str]) -> pd.DataFrame:
    rename_map = {
        header_mapping[canonical]: canonical
        for canonical in header_mapping
        if canonical in {"cycle", "mode", "voltage", "specific_capacity", "current", "capacity", "step", "record"}
    }
    standardized = raw_df.rename(columns=rename_map).copy()
    standardized.columns = [str(column).strip() for column in standardized.columns]
    return standardized


def clean_dataset(
    raw_df: pd.DataFrame,
    sheet_name: str,
    header_row: int,
    header_mapping: dict[str, str],
    cycles: list[int] | None = None,
    mode_overrides: dict[str, str] | None = None,
    auto_sort: bool = True,
    absolute_specific_capacity: bool = True,
) -> ParsedDataset:
    required = {"cycle", "mode", "voltage", "specific_capacity"}
    missing = required.difference(header_mapping)
    if missing:
        missing_names = ", ".join(sorted(missing))
        raise DataParsingError(f"缺少关键列: {missing_names}")

    data = standardize_columns(raw_df, header_mapping)
    data = data.loc[:, ~data.columns.duplicated()].copy()
    data["_row_order"] = np.arange(len(data))

    mode_lookup = build_mode_lookup(mode_overrides)

    data["cycle"] = _coerce_cycle_number(data["cycle"])
    data["mode_text"] = data["mode"].astype(str)
    data["curve_type"] = data["mode_text"].map(lambda item: classify_mode(item, mode_lookup))
    data["voltage"] = pd.to_numeric(data["voltage"], errors="coerce")
    data["specific_capacity"] = pd.to_numeric(data["specific_capacity"], errors="coerce")

    if "step" in data.columns:
        data["step"] = pd.to_numeric(data["step"], errors="coerce")
    if "record" in data.columns:
        data["record"] = pd.to_numeric(data["record"], errors="coerce")
    if "current" in data.columns:
        data["current"] = pd.to_numeric(data["current"], errors="coerce")
    if "capacity" in data.columns:
        data["capacity"] = pd.to_numeric(data["capacity"], errors="coerce")

    if absolute_specific_capacity:
        data["specific_capacity"] = data["specific_capacity"].abs()

    data = data[
        data["cycle"].notna()
        & data["voltage"].notna()
        & data["specific_capacity"].notna()
        & np.isfinite(data["voltage"])
        & np.isfinite(data["specific_capacity"])
    ].copy()

    data = data[data["curve_type"].isin(["charge", "discharge"])].copy()
    data = data[(data["specific_capacity"] >= 0) & (data["voltage"] > -10) & (data["voltage"] < 10)].copy()
    data["cycle"] = data["cycle"].astype(int)

    if cycles:
        selected = set(cycles)
        data = data[data["cycle"].isin(selected)].copy()

    sort_columns = [column for column in ["cycle", "step", "record", "specific_capacity", "_row_order"] if column in data.columns]
    if auto_sort and sort_columns:
        data = data.sort_values(sort_columns, kind="mergesort")
    else:
        data = data.sort_values(["cycle", "_row_order"], kind="mergesort")

    data = data.reset_index(drop=True)
    if data.empty:
        raise DataParsingError("没有找到可绘制的充放电逐点数据，请检查 sheet、表头或循环范围。")

    return ParsedDataset(
        data=data,
        sheet_name=sheet_name,
        header_row=header_row,
        source_columns={canonical: header_mapping[canonical] for canonical in header_mapping},
    )
