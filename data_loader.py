from __future__ import annotations

from pathlib import Path

import numpy as np
import pandas as pd

from config import SCAN_PREVIEW_ROWS
from parser import DataParsingError, ParsedDataset, clean_dataset, detect_header_row


class DataLoaderError(RuntimeError):
    """Raised when the input file cannot be loaded as a battery cycling dataset."""


def _resolve_sheet_name(excel_file: pd.ExcelFile, requested: str | None) -> str | None:
    if requested is None or not requested.strip():
        return None

    requested_text = requested.strip()
    for sheet_name in excel_file.sheet_names:
        if sheet_name == requested_text:
            return sheet_name
    for sheet_name in excel_file.sheet_names:
        if sheet_name.lower() == requested_text.lower():
            return sheet_name

    available = ", ".join(excel_file.sheet_names)
    raise DataLoaderError(f"未找到 sheet: {requested_text}。可用 sheet: {available}")


def _read_preview(file_path: Path, sheet_name: str) -> pd.DataFrame:
    return pd.read_excel(
        file_path,
        sheet_name=sheet_name,
        header=None,
        nrows=SCAN_PREVIEW_ROWS,
        dtype=str,
    )


def _detect_best_sheet(file_path: Path, excel_file: pd.ExcelFile, requested_sheet: str | None) -> tuple[str, int, dict[str, str]]:
    if requested_sheet is not None:
        preview = _read_preview(file_path, requested_sheet)
        detected = detect_header_row(requested_sheet, preview)
        if detected is None:
            raise DataLoaderError(f"在 sheet '{requested_sheet}' 中没有识别到包含逐点曲线数据的表头。")
        return detected.sheet_name, detected.header_row, detected.mapping

    best_detection = None
    for sheet_name in excel_file.sheet_names:
        preview = _read_preview(file_path, sheet_name)
        detected = detect_header_row(sheet_name, preview)
        if detected is None:
            continue
        if best_detection is None or detected.score > best_detection.score:
            best_detection = detected

    if best_detection is None:
        raise DataLoaderError("未能在任何 sheet 中找到包含 电压/V 和 比容量/mAh/g 的逐点数据表头。")
    return best_detection.sheet_name, best_detection.header_row, best_detection.mapping


def load_battery_dataset(
    file_path: Path,
    sheet_name: str | None = None,
    cycles: list[int] | None = None,
    mode_overrides: dict[str, str] | None = None,
    auto_sort: bool = True,
    absolute_specific_capacity: bool = True,
) -> ParsedDataset:
    suffix = file_path.suffix.lower()
    if suffix == ".ccs":
        raise DataLoaderError("暂未稳定支持 .ccs 原始文件，请先在设备软件中导出为 .xlsx 再处理。")

    try:
        excel_file = pd.ExcelFile(file_path)
    except Exception as exc:
        raise DataLoaderError(f"无法打开 Excel 文件: {file_path}") from exc

    resolved_sheet = _resolve_sheet_name(excel_file, sheet_name)
    selected_sheet, header_row, header_mapping = _detect_best_sheet(file_path, excel_file, resolved_sheet)

    try:
        raw_df = pd.read_excel(file_path, sheet_name=selected_sheet, header=header_row, dtype=str)
    except Exception as exc:
        raise DataLoaderError(f"读取 sheet '{selected_sheet}' 失败。") from exc

    try:
        return clean_dataset(
            raw_df=raw_df,
            sheet_name=selected_sheet,
            header_row=header_row,
            header_mapping=header_mapping,
            cycles=cycles,
            mode_overrides=mode_overrides,
            auto_sort=auto_sort,
            absolute_specific_capacity=absolute_specific_capacity,
        )
    except DataParsingError as exc:
        raise DataLoaderError(str(exc)) from exc


def generate_demo_dataset(cycle_count: int = 4, points_per_segment: int = 120) -> ParsedDataset:
    rows: list[dict[str, object]] = []
    for cycle in range(1, cycle_count + 1):
        charge_capacity = np.linspace(0, 160 - cycle * 3, points_per_segment)
        discharge_capacity = np.linspace(0, 158 - cycle * 3.5, points_per_segment)
        charge_voltage = 2.8 + 1.25 * (1 - np.exp(-charge_capacity / 55)) + cycle * 0.01
        discharge_voltage = 4.15 - 1.05 * (discharge_capacity / discharge_capacity.max()) ** 0.8 - cycle * 0.015

        for record, (capacity, voltage) in enumerate(zip(charge_capacity, charge_voltage), start=1):
            rows.append(
                {
                    "cycle": cycle,
                    "curve_type": "charge",
                    "mode_text": "恒流充电",
                    "specific_capacity": float(capacity),
                    "voltage": float(voltage),
                    "record": record,
                }
            )
        for record, (capacity, voltage) in enumerate(zip(discharge_capacity, discharge_voltage), start=1):
            rows.append(
                {
                    "cycle": cycle,
                    "curve_type": "discharge",
                    "mode_text": "恒流放电",
                    "specific_capacity": float(capacity),
                    "voltage": float(voltage),
                    "record": record,
                }
            )

    demo_df = pd.DataFrame(rows)
    return ParsedDataset(
        data=demo_df,
        sheet_name="demo",
        header_row=0,
        source_columns={
            "cycle": "cycle",
            "mode": "mode_text",
            "voltage": "voltage",
            "specific_capacity": "specific_capacity",
        },
    )
