from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path

SUPPORTED_EXCEL_SUFFIXES = {".xlsx", ".xlsm"}
SCAN_PREVIEW_ROWS = 80
DEFAULT_OUTPUT_FORMAT = "png"
DEFAULT_COLORMAP = "tab10"
DEFAULT_THEME = "paper"
DEFAULT_FONT_CANDIDATES = [
    "Microsoft YaHei",
    "SimHei",
    "Noto Sans CJK SC",
    "Source Han Sans SC",
    "Arial Unicode MS",
    "DejaVu Sans",
]

DEFAULT_MODE_RULES: dict[str, list[str]] = {
    "charge": [
        "charge",
        "charging",
        "cc_charge",
        "cv_charge",
        "恒流充电",
        "恒压充电",
        "充电",
        "充",
    ],
    "discharge": [
        "discharge",
        "discharging",
        "cc_discharge",
        "dc_discharge",
        "恒流放电",
        "放电",
        "放",
    ],
    "rest": [
        "rest",
        "pause",
        "idle",
        "ocv",
        "静置",
        "搁置",
        "开路",
    ],
}


@dataclass
class AppConfig:
    input_path: Path | None = None
    sheet_name: str | None = None
    output_path: Path | None = None
    output_format: str = DEFAULT_OUTPUT_FORMAT
    dpi: int = 300
    cycles: list[int] | None = None
    show_legend: bool = True
    legend_loc: str = "best"
    title: str = "Battery Charge-Discharge Curves"
    x_label: str = "Specific Capacity (mAh/g)"
    y_label: str = "Voltage (V)"
    title_fontsize: int = 16
    label_fontsize: int = 13
    tick_fontsize: int = 11
    font_family: list[str] = field(default_factory=lambda: list(DEFAULT_FONT_CANDIDATES))
    line_width: float = 1.8
    line_style_charge: str = "-"
    line_style_discharge: str = "--"
    line_color: str = "#1f77b4"
    grid: bool = True
    grid_style: str = "--"
    figure_width: float = 8.0
    figure_height: float = 6.0
    color_by_cycle: bool = True
    colormap: str = DEFAULT_COLORMAP
    x_lim: tuple[float, float] | None = None
    y_lim: tuple[float, float] | None = None
    auto_sort: bool = True
    show_after_save: bool = False
    transparent_background: bool = False
    theme: str = DEFAULT_THEME
    mode_overrides: dict[str, str] = field(default_factory=dict)
    absolute_specific_capacity: bool = True
    sheet_auto_detect: bool = True

    def normalized_output_format(self) -> str:
        return self.output_format.lower().lstrip(".") or DEFAULT_OUTPUT_FORMAT

    def figure_size(self) -> tuple[float, float]:
        return (self.figure_width, self.figure_height)
