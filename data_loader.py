from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path, PurePath, PurePosixPath, PureWindowsPath

import numpy as np
import pandas as pd

from config import SCAN_PREVIEW_ROWS
from parser import DataParsingError, HeaderDetectionResult, ParsedDataset, clean_dataset, detect_header_row, detect_header_rows, normalize_text


class DataLoaderError(RuntimeError):
    """Raised when the input file cannot be loaded as a battery cycling dataset."""


@dataclass
class DatasetLoadItem:
    dataset: ParsedDataset
    output_stem: str
    source_path: str | None = None


@dataclass
class _SourceMetadata:
    output_stem: str
    source_path: str | None
    source_stem: str


@dataclass
class _DetectedBlock:
    detection: HeaderDetectionResult
    raw_df: pd.DataFrame
    block_index_in_sheet: int
    total_blocks_in_sheet: int


_SOURCE_PATH_PATTERN = re.compile(r"\.(?:ccs|xlsx|xlsm)$", re.IGNORECASE)
_INFO_HEADER_ALIASES: dict[str, list[str]] = {
    "file_name": ["文件名", "filename", "file name", "源文件名"],
    "path": ["路径", "path", "文件路径", "源路径", "保存路径"],
}


def _build_alias_lookup(aliases: dict[str, list[str]]) -> dict[str, str]:
    lookup: dict[str, str] = {}
    for canonical, items in aliases.items():
        for item in items:
            lookup[normalize_text(item)] = canonical
    return lookup


_INFO_ALIAS_LOOKUP = _build_alias_lookup(_INFO_HEADER_ALIASES)


def _make_pure_path(text: str) -> PurePath:
    cleaned = text.strip()
    if re.match(r"^[A-Za-z]:", cleaned) or "\\" in cleaned:
        return PureWindowsPath(cleaned)
    return PurePosixPath(cleaned)


def _sanitize_file_stem(name: str) -> str:
    cleaned = re.sub(r'[<>:"/\\|?*]+', "_", name).strip().rstrip(".")
    return cleaned or "dataset"


def _derive_output_stem(source_path: str | None, fallback: str) -> str:
    if source_path:
        pure_path = _make_pure_path(source_path)
        stem = pure_path.stem or pure_path.name
        parent_name = pure_path.parent.name.strip()
        if parent_name:
            return _sanitize_file_stem(f"{parent_name}_{stem}")
        if stem:
            return _sanitize_file_stem(stem)
    return _sanitize_file_stem(fallback)


def _compose_source_path(path_text: str | None, file_name: str | None) -> str | None:
    raw_path = (path_text or "").strip()
    raw_file_name = (file_name or "").strip()

    if raw_path and _SOURCE_PATH_PATTERN.search(raw_path):
        return raw_path
    if raw_file_name and _SOURCE_PATH_PATTERN.search(raw_file_name):
        if raw_path:
            return str(_make_pure_path(raw_path) / raw_file_name)
        return raw_file_name
    if raw_path:
        return raw_path
    if raw_file_name:
        return raw_file_name
    return None


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


def _read_sheet_raw(file_path: Path, sheet_name: str) -> pd.DataFrame:
    return pd.read_excel(
        file_path,
        sheet_name=sheet_name,
        header=None,
        dtype=str,
    )


def _build_block_raw_df(sheet_df: pd.DataFrame, header_row: int, next_header_row: int | None) -> pd.DataFrame:
    header_values = ["" if pd.isna(item) else str(item).strip() for item in sheet_df.iloc[header_row].tolist()]
    end_row = next_header_row if next_header_row is not None else len(sheet_df.index)
    block_df = sheet_df.iloc[header_row + 1 : end_row].copy().reset_index(drop=True)
    block_df.columns = header_values
    return block_df


def _detect_sheet_blocks(
    file_path: Path,
    excel_file: pd.ExcelFile,
    requested_sheet: str | None,
) -> list[_DetectedBlock]:
    required_fields = {"cycle", "mode", "voltage", "specific_capacity"}
    target_sheets = [requested_sheet] if requested_sheet is not None else list(excel_file.sheet_names)
    detected_blocks: list[_DetectedBlock] = []

    for sheet_name in target_sheets:
        try:
            sheet_df = _read_sheet_raw(file_path, sheet_name)
        except Exception as exc:
            raise DataLoaderError(f"读取 sheet '{sheet_name}' 失败。") from exc

        header_candidates = detect_header_rows(
            sheet_name=sheet_name,
            preview_df=sheet_df,
            required_fields=required_fields,
        )

        if requested_sheet is not None and not header_candidates:
            raise DataLoaderError(f"在 sheet '{requested_sheet}' 中没有识别到包含逐点曲线数据的表头。")

        for index, detection in enumerate(header_candidates):
            next_header_row = header_candidates[index + 1].header_row if index + 1 < len(header_candidates) else None
            detected_blocks.append(
                _DetectedBlock(
                    detection=detection,
                    raw_df=_build_block_raw_df(sheet_df, detection.header_row, next_header_row),
                    block_index_in_sheet=index,
                    total_blocks_in_sheet=len(header_candidates),
                )
            )

    if not detected_blocks:
        raise DataLoaderError("未能在任何 sheet 中找到包含 电压/V 和 比容量/mAh/g 的逐点数据表头。")
    return detected_blocks


def _detect_candidate_sheets(
    file_path: Path,
    excel_file: pd.ExcelFile,
    requested_sheet: str | None,
) -> list[HeaderDetectionResult]:
    if requested_sheet is not None:
        preview = _read_preview(file_path, requested_sheet)
        detected = detect_header_row(requested_sheet, preview)
        if detected is None:
            raise DataLoaderError(f"在 sheet '{requested_sheet}' 中没有识别到包含逐点曲线数据的表头。")
        return [detected]

    detections: list[HeaderDetectionResult] = []
    for sheet_name in excel_file.sheet_names:
        preview = _read_preview(file_path, sheet_name)
        detected = detect_header_row(sheet_name, preview)
        if detected is not None:
            detections.append(detected)

    if not detections:
        raise DataLoaderError("未能在任何 sheet 中找到包含 电压/V 和 比容量/mAh/g 的逐点数据表头。")
    return detections


def _detect_best_sheet(file_path: Path, excel_file: pd.ExcelFile, requested_sheet: str | None) -> tuple[str, int, dict[str, str]]:
    detections = _detect_candidate_sheets(file_path, excel_file, requested_sheet)
    best_detection = max(detections, key=lambda item: item.score)
    return best_detection.sheet_name, best_detection.header_row, best_detection.mapping


def _resolve_info_columns(headers: list[str]) -> dict[str, str]:
    resolved: dict[str, str] = {}
    for header in headers:
        canonical = _INFO_ALIAS_LOOKUP.get(normalize_text(header))
        if canonical and canonical not in resolved:
            resolved[canonical] = header
    return resolved


def _read_source_metadata(file_path: Path, excel_file: pd.ExcelFile) -> list[_SourceMetadata]:
    for sheet_name in excel_file.sheet_names:
        preview = _read_preview(file_path, sheet_name)
        max_rows = min(len(preview.index), 8)
        for row_index in range(max_rows):
            headers = ["" if pd.isna(item) else str(item).strip() for item in preview.iloc[row_index].tolist()]
            resolved_columns = _resolve_info_columns(headers)
            if "file_name" not in resolved_columns and "path" not in resolved_columns:
                continue

            try:
                info_df = pd.read_excel(file_path, sheet_name=sheet_name, header=row_index, dtype=str)
            except Exception:
                continue

            metadata_rows: list[_SourceMetadata] = []
            for _, row in info_df.iterrows():
                file_name = row.get(resolved_columns["file_name"]) if "file_name" in resolved_columns else None
                path_text = row.get(resolved_columns["path"]) if "path" in resolved_columns else None
                source_path = _compose_source_path(
                    None if pd.isna(path_text) else str(path_text),
                    None if pd.isna(file_name) else str(file_name),
                )
                if not source_path:
                    continue

                pure_source = _make_pure_path(source_path)
                source_stem = pure_source.stem or pure_source.name
                metadata_rows.append(
                    _SourceMetadata(
                        output_stem=_derive_output_stem(source_path, source_stem),
                        source_path=source_path,
                        source_stem=normalize_text(source_stem),
                    )
                )

            if metadata_rows:
                return metadata_rows
    return []


def _match_metadata_for_sheet(
    sheet_name: str,
    metadata_rows: list[_SourceMetadata],
    fallback_index: int,
    total_datasets: int,
    used_indices: set[int],
) -> _SourceMetadata | None:
    normalized_sheet = normalize_text(sheet_name)

    exact_matches: list[tuple[int, _SourceMetadata]] = []
    fuzzy_matches: list[tuple[int, _SourceMetadata]] = []
    for index, record in enumerate(metadata_rows):
        if index in used_indices:
            continue
        if not record.source_stem:
            continue
        if record.source_stem == normalized_sheet or record.source_stem.endswith(normalized_sheet):
            exact_matches.append((index, record))
        elif normalized_sheet and normalized_sheet in record.source_stem:
            fuzzy_matches.append((index, record))

    matched: tuple[int, _SourceMetadata] | None = None
    if exact_matches:
        matched = exact_matches[0]
    elif fuzzy_matches:
        matched = fuzzy_matches[0]
    elif len(metadata_rows) == total_datasets and fallback_index < len(metadata_rows):
        matched = (fallback_index, metadata_rows[fallback_index])

    if matched is None:
        return None

    matched_index, record = matched
    used_indices.add(matched_index)
    return record


def _build_fallback_output_stem(sheet_name: str, block_index_in_sheet: int, total_blocks_in_sheet: int) -> str:
    if total_blocks_in_sheet <= 1:
        return _derive_output_stem(None, sheet_name)
    return _derive_output_stem(None, f"{sheet_name}_part{block_index_in_sheet + 1}")


def _merge_parsed_dataset_group(group: list[tuple[_DetectedBlock, ParsedDataset]]) -> tuple[_DetectedBlock, ParsedDataset]:
    first_block, first_dataset = group[0]
    if len(group) == 1:
        return first_block, first_dataset

    merged_data = pd.concat([dataset.data for _, dataset in group], ignore_index=True)
    merged_dataset = ParsedDataset(
        data=merged_data,
        sheet_name=first_dataset.sheet_name,
        header_row=first_dataset.header_row,
        source_columns=first_dataset.source_columns,
    )
    return first_block, merged_dataset


def _collapse_parsed_items_to_expected_count(
    parsed_items: list[tuple[_DetectedBlock, ParsedDataset]],
    expected_count: int,
) -> list[tuple[_DetectedBlock, ParsedDataset]]:
    if expected_count <= 0 or len(parsed_items) <= expected_count:
        return parsed_items

    grouped_items: list[list[tuple[_DetectedBlock, ParsedDataset]]] = []
    total_items = len(parsed_items)
    start = 0
    for group_index in range(expected_count):
        remaining_items = total_items - start
        remaining_groups = expected_count - group_index
        chunk_size = -(-remaining_items // remaining_groups)
        grouped_items.append(parsed_items[start : start + chunk_size])
        start += chunk_size

    return [_merge_parsed_dataset_group(group) for group in grouped_items if group]


def _deduplicate_output_stems(items: list[DatasetLoadItem]) -> list[DatasetLoadItem]:
    total_counts: dict[str, int] = {}
    for item in items:
        total_counts[item.output_stem] = total_counts.get(item.output_stem, 0) + 1

    seen_counts: dict[str, int] = {}
    for item in items:
        stem = item.output_stem
        if total_counts[stem] <= 1:
            continue
        seen_counts[stem] = seen_counts.get(stem, 0) + 1
        item.output_stem = f"{stem}_part{seen_counts[stem]}"
    return items


def load_battery_dataset(
    file_path: Path,
    sheet_name: str | None = None,
    cycles: list[int] | None = None,
    mode_overrides: dict[str, str] | None = None,
    auto_sort: bool = True,
    absolute_specific_capacity: bool = True,
) -> ParsedDataset:
    items = load_battery_datasets(
        file_path=file_path,
        sheet_name=sheet_name,
        cycles=cycles,
        mode_overrides=mode_overrides,
        auto_sort=auto_sort,
        absolute_specific_capacity=absolute_specific_capacity,
    )
    return items[0].dataset


def load_battery_datasets(
    file_path: Path,
    sheet_name: str | None = None,
    cycles: list[int] | None = None,
    mode_overrides: dict[str, str] | None = None,
    auto_sort: bool = True,
    absolute_specific_capacity: bool = True,
) -> list[DatasetLoadItem]:
    suffix = file_path.suffix.lower()
    if suffix == ".ccs":
        raise DataLoaderError("暂未稳定支持 .ccs 原始文件，请先在设备软件中导出为 .xlsx 再处理。")

    try:
        excel_file = pd.ExcelFile(file_path)
    except Exception as exc:
        raise DataLoaderError(f"无法打开 Excel 文件: {file_path}") from exc

    resolved_sheet = _resolve_sheet_name(excel_file, sheet_name)
    detected_blocks = _detect_sheet_blocks(file_path, excel_file, resolved_sheet)
    parsed_items: list[tuple[_DetectedBlock, ParsedDataset]] = []
    last_parse_error: DataParsingError | None = None

    for block in detected_blocks:
        detection = block.detection
        try:
            dataset = clean_dataset(
                raw_df=block.raw_df,
                sheet_name=detection.sheet_name,
                header_row=detection.header_row,
                header_mapping=detection.mapping,
                cycles=cycles,
                mode_overrides=mode_overrides,
                auto_sort=auto_sort,
                absolute_specific_capacity=absolute_specific_capacity,
            )
        except DataParsingError as exc:
            if resolved_sheet is not None:
                raise DataLoaderError(str(exc)) from exc
            last_parse_error = exc
            continue

        parsed_items.append((block, dataset))

    if not parsed_items:
        if last_parse_error is not None:
            raise DataLoaderError(str(last_parse_error)) from last_parse_error
        raise DataLoaderError("没有找到可绘制的充放电子表格。")

    metadata_rows = _read_source_metadata(file_path, excel_file)
    if metadata_rows:
        parsed_items = _collapse_parsed_items_to_expected_count(parsed_items, len(metadata_rows))
    used_metadata_indices: set[int] = set()
    loaded_items: list[DatasetLoadItem] = []

    for detection_index, (block, dataset) in enumerate(parsed_items):
        detection = block.detection
        matched_metadata = _match_metadata_for_sheet(
            sheet_name=detection.sheet_name,
            metadata_rows=metadata_rows,
            fallback_index=detection_index,
            total_datasets=len(parsed_items),
            used_indices=used_metadata_indices,
        )
        output_stem = (
            matched_metadata.output_stem
            if matched_metadata is not None
            else _build_fallback_output_stem(
                sheet_name=detection.sheet_name,
                block_index_in_sheet=block.block_index_in_sheet,
                total_blocks_in_sheet=block.total_blocks_in_sheet,
            )
        )
        source_path = matched_metadata.source_path if matched_metadata is not None else None
        loaded_items.append(
            DatasetLoadItem(
                dataset=dataset,
                output_stem=output_stem,
                source_path=source_path,
            )
        )

    return _deduplicate_output_stems(loaded_items)


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
