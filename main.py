from __future__ import annotations

import argparse
import sys
from pathlib import Path

from config import AppConfig
from data_loader import DataLoaderError, generate_demo_dataset, load_battery_dataset
from plotter import PlottingError, plot_voltage_specific_capacity
from utils import (
    UserInputError,
    build_output_path,
    parse_axis_limits,
    parse_bool_text,
    parse_cycle_expression,
    parse_figure_size,
    parse_float_text,
    parse_font_family,
    parse_int_text,
    parse_mode_overrides,
    parse_optional_color,
    parse_output_format,
    prompt_text,
    prompt_until_valid,
    validate_input_file,
)


def build_argument_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="绘制电池充放电曲线（Voltage vs Specific Capacity）。默认直接运行会进入交互式输入。"
    )
    parser.add_argument("--input", help="输入 Excel 文件路径（.xlsx / .xlsm）")
    parser.add_argument("--sheet", help="要读取的 sheet 名")
    parser.add_argument("--output", help="输出图片路径")
    parser.add_argument("--format", dest="output_format", help="输出格式：png / svg / pdf")
    parser.add_argument("--cycles", nargs="*", help="循环编号，例如 1 2 3 或 1-5")
    parser.add_argument("--dpi", type=int, help="导出 DPI")
    parser.add_argument("--legend-loc", help="图例位置，例如 best / upper right")
    parser.add_argument("--title", help="图标题")
    parser.add_argument("--x-label", help="X 轴标题")
    parser.add_argument("--y-label", help="Y 轴标题")
    parser.add_argument("--title-fontsize", type=int, help="标题字号")
    parser.add_argument("--label-fontsize", type=int, help="坐标轴标题字号")
    parser.add_argument("--tick-fontsize", type=int, help="刻度字号")
    parser.add_argument("--font-family", help="字体列表，逗号分隔")
    parser.add_argument("--line-width", type=float, help="线宽")
    parser.add_argument("--line-style-charge", help="充电曲线线型")
    parser.add_argument("--line-style-discharge", help="放电曲线线型")
    parser.add_argument("--line-color", help="统一颜色，例如 #1f77b4")
    parser.add_argument("--grid-style", help="网格线型")
    parser.add_argument("--figsize", nargs=2, metavar=("WIDTH", "HEIGHT"), help="图尺寸，单位英寸")
    parser.add_argument("--colormap", help="colormap 名称，例如 tab10 / viridis")
    parser.add_argument("--theme", choices=["paper", "default"], help="绘图主题")
    parser.add_argument("--x-lim", nargs=2, metavar=("MIN", "MAX"), help="X 轴范围")
    parser.add_argument("--y-lim", nargs=2, metavar=("MIN", "MAX"), help="Y 轴范围")
    parser.add_argument("--mode-map", nargs="*", help="工作模式映射补充，例如 恒流充电=charge 静置=rest")
    parser.add_argument("--interactive", action="store_true", help="强制进入交互式输入")
    parser.add_argument("--demo", action="store_true", help="使用内置示例数据生成演示图")
    parser.add_argument("--show-legend", dest="show_legend", action=argparse.BooleanOptionalAction, default=None)
    parser.add_argument("--grid", dest="grid", action=argparse.BooleanOptionalAction, default=None)
    parser.add_argument("--color-by-cycle", dest="color_by_cycle", action=argparse.BooleanOptionalAction, default=None)
    parser.add_argument("--auto-sort", dest="auto_sort", action=argparse.BooleanOptionalAction, default=None)
    parser.add_argument("--show-after-save", dest="show_after_save", action=argparse.BooleanOptionalAction, default=None)
    parser.add_argument("--transparent", dest="transparent_background", action=argparse.BooleanOptionalAction, default=None)
    parser.add_argument(
        "--absolute-specific-capacity",
        dest="absolute_specific_capacity",
        action=argparse.BooleanOptionalAction,
        default=None,
        help="是否对比容量取绝对值",
    )
    return parser


def sync_output_settings(config: AppConfig) -> None:
    if config.output_path is None:
        return
    target_suffix = f".{config.normalized_output_format()}"
    if config.output_path.suffix.lower() != target_suffix:
        config.output_path = config.output_path.with_suffix(target_suffix)


def apply_cli_arguments(args: argparse.Namespace) -> AppConfig:
    config = AppConfig()

    if args.input:
        config.input_path = validate_input_file(args.input)
    if args.sheet is not None:
        config.sheet_name = args.sheet.strip() or None
    if args.output_format:
        config.output_format = parse_output_format(args.output_format, config.output_format)
    if args.output:
        config.output_path = build_output_path(args.output, config.output_format)
        if args.output_format is None and config.output_path.suffix:
            config.output_format = config.output_path.suffix.lstrip(".").lower()
    if args.cycles is not None:
        config.cycles = parse_cycle_expression(args.cycles)
    if args.dpi is not None:
        config.dpi = args.dpi
    if args.show_legend is not None:
        config.show_legend = args.show_legend
    if args.legend_loc:
        config.legend_loc = args.legend_loc
    if args.title:
        config.title = args.title
    if args.x_label:
        config.x_label = args.x_label
    if args.y_label:
        config.y_label = args.y_label
    if args.title_fontsize is not None:
        config.title_fontsize = args.title_fontsize
    if args.label_fontsize is not None:
        config.label_fontsize = args.label_fontsize
    if args.tick_fontsize is not None:
        config.tick_fontsize = args.tick_fontsize
    if args.font_family is not None:
        config.font_family = parse_font_family(args.font_family)
    if args.line_width is not None:
        config.line_width = args.line_width
    if args.line_style_charge:
        config.line_style_charge = args.line_style_charge
    if args.line_style_discharge:
        config.line_style_discharge = args.line_style_discharge
    if args.line_color:
        config.line_color = args.line_color
    if args.grid is not None:
        config.grid = args.grid
    if args.grid_style:
        config.grid_style = args.grid_style
    if args.figsize:
        config.figure_width = float(args.figsize[0])
        config.figure_height = float(args.figsize[1])
    if args.color_by_cycle is not None:
        config.color_by_cycle = args.color_by_cycle
    if args.colormap:
        config.colormap = args.colormap
    if args.theme:
        config.theme = args.theme
    if args.x_lim:
        config.x_lim = (float(args.x_lim[0]), float(args.x_lim[1]))
    if args.y_lim:
        config.y_lim = (float(args.y_lim[0]), float(args.y_lim[1]))
    if args.auto_sort is not None:
        config.auto_sort = args.auto_sort
    if args.show_after_save is not None:
        config.show_after_save = args.show_after_save
    if args.transparent_background is not None:
        config.transparent_background = args.transparent_background
    if args.mode_map is not None:
        config.mode_overrides = parse_mode_overrides(args.mode_map)
    if args.absolute_specific_capacity is not None:
        config.absolute_specific_capacity = args.absolute_specific_capacity

    sync_output_settings(config)
    return config


def prompt_existing_file(current: Path | None, label: str) -> Path:
    default_hint = f" [{current}]" if current else ""
    return prompt_until_valid(
        f"{label}{default_hint}: ",
        lambda text: current if not text.strip() and current else validate_input_file(text),
    )


def prompt_output_file(config: AppConfig, label: str) -> Path:
    current = str(config.output_path) if config.output_path else ""
    prompt = f"{label}{f' [{current}]' if current else ''}: "
    return prompt_until_valid(
        prompt,
        lambda text: config.output_path if not text.strip() and config.output_path else build_output_path(text, config.output_format),
    )


def configure_interactively(config: AppConfig, prompt_all: bool, use_demo: bool) -> AppConfig:
    print("\n=== 电池充放电曲线绘图工具 ===")
    print("直接回车即可使用默认值。\n")

    if not use_demo and (prompt_all or config.input_path is None):
        config.input_path = prompt_existing_file(config.input_path, "请输入数据文件路径")

    if prompt_all or config.sheet_name is None:
        sheet_default = config.sheet_name or ""
        sheet_text = prompt_text(
            f"请输入 sheet 名（直接回车则自动识别）{f' [{sheet_default}]' if sheet_default else ''}: ",
            default=sheet_default,
        )
        config.sheet_name = sheet_text or None

    if prompt_all or config.output_path is None:
        config.output_path = prompt_output_file(config, "请输入输出图片保存路径")
        if config.output_path.suffix:
            config.output_format = config.output_path.suffix.lstrip(".").lower()

    if prompt_all or config.cycles is None:
        cycles_default = ",".join(str(item) for item in config.cycles) if config.cycles else ""
        config.cycles = prompt_until_valid(
            f"请输入要绘制的循环编号，例如 1,2,3 或 1-5（直接回车表示全部循环）{f' [{cycles_default}]' if cycles_default else ''}: ",
            lambda text: parse_cycle_expression(text if text.strip() else cycles_default),
        )

    if prompt_all:
        config.show_legend = prompt_until_valid(
            f"是否显示图例（y/n） [{'y' if config.show_legend else 'n'}]: ",
            lambda text: parse_bool_text(text, config.show_legend),
        )
        config.title = prompt_text(
            f"请输入标题（直接回车使用默认标题） [{config.title}]: ",
            default=config.title,
        )

    advanced = prompt_until_valid(
        "是否进入高级参数设置（y/n，直接回车默认 n）: ",
        lambda text: parse_bool_text(text, False),
    )
    if not advanced:
        sync_output_settings(config)
        return config

    config.output_format = prompt_until_valid(
        f"请输入输出格式（png/svg/pdf） [{config.output_format}]: ",
        lambda text: parse_output_format(text, config.output_format),
    )
    sync_output_settings(config)

    config.dpi = prompt_until_valid(
        f"请输入输出 DPI [{config.dpi}]: ",
        lambda text: parse_int_text(text, config.dpi, minimum=50),
    )
    config.legend_loc = prompt_text(f"请输入图例位置 [{config.legend_loc}]: ", default=config.legend_loc)
    config.x_label = prompt_text(f"请输入 X 轴标题 [{config.x_label}]: ", default=config.x_label)
    config.y_label = prompt_text(f"请输入 Y 轴标题 [{config.y_label}]: ", default=config.y_label)
    config.title_fontsize = prompt_until_valid(
        f"请输入标题字号 [{config.title_fontsize}]: ",
        lambda text: parse_int_text(text, config.title_fontsize, minimum=1),
    )
    config.label_fontsize = prompt_until_valid(
        f"请输入坐标轴标题字号 [{config.label_fontsize}]: ",
        lambda text: parse_int_text(text, config.label_fontsize, minimum=1),
    )
    config.tick_fontsize = prompt_until_valid(
        f"请输入刻度字号 [{config.tick_fontsize}]: ",
        lambda text: parse_int_text(text, config.tick_fontsize, minimum=1),
    )
    config.line_width = prompt_until_valid(
        f"请输入线宽 [{config.line_width}]: ",
        lambda text: parse_float_text(text, config.line_width, minimum=0.1),
    )
    config.line_color = parse_optional_color(
        prompt_text(f"请输入统一线条颜色（为空使用当前值） [{config.line_color}]: ", default=config.line_color),
        config.line_color,
    )
    config.line_style_charge = prompt_text(
        f"请输入充电曲线线型 [{config.line_style_charge}]: ",
        default=config.line_style_charge,
    )
    config.line_style_discharge = prompt_text(
        f"请输入放电曲线线型 [{config.line_style_discharge}]: ",
        default=config.line_style_discharge,
    )
    config.grid = prompt_until_valid(
        f"是否显示网格（y/n） [{'y' if config.grid else 'n'}]: ",
        lambda text: parse_bool_text(text, config.grid),
    )
    config.grid_style = prompt_text(f"请输入网格线型 [{config.grid_style}]: ", default=config.grid_style)
    figure_width, figure_height = prompt_until_valid(
        f"请输入图尺寸（宽,高，单位英寸） [{config.figure_width},{config.figure_height}]: ",
        lambda text: parse_figure_size(text, config.figure_size()),
    )
    config.figure_width = figure_width
    config.figure_height = figure_height
    config.color_by_cycle = prompt_until_valid(
        f"是否按循环分别着色（y/n） [{'y' if config.color_by_cycle else 'n'}]: ",
        lambda text: parse_bool_text(text, config.color_by_cycle),
    )
    config.colormap = prompt_text(f"请输入 colormap 名称 [{config.colormap}]: ", default=config.colormap)
    config.theme = prompt_text(f"请输入主题（paper/default） [{config.theme}]: ", default=config.theme)
    config.font_family = parse_font_family(
        prompt_text(
            f"请输入字体列表（逗号分隔） [{','.join(config.font_family)}]: ",
            default=",".join(config.font_family),
        )
    )
    config.x_lim = prompt_until_valid(
        "请输入 X 轴范围（如 0,200；直接回车跳过）: ",
        parse_axis_limits,
    )
    config.y_lim = prompt_until_valid(
        "请输入 Y 轴范围（如 2.5,4.5；直接回车跳过）: ",
        parse_axis_limits,
    )
    config.auto_sort = prompt_until_valid(
        f"是否自动排序数据点（y/n） [{'y' if config.auto_sort else 'n'}]: ",
        lambda text: parse_bool_text(text, config.auto_sort),
    )
    config.absolute_specific_capacity = prompt_until_valid(
        f"是否对比容量取绝对值（y/n） [{'y' if config.absolute_specific_capacity else 'n'}]: ",
        lambda text: parse_bool_text(text, config.absolute_specific_capacity),
    )
    config.show_after_save = prompt_until_valid(
        f"保存后是否同时显示图像（y/n） [{'y' if config.show_after_save else 'n'}]: ",
        lambda text: parse_bool_text(text, config.show_after_save),
    )
    config.transparent_background = prompt_until_valid(
        f"是否使用透明背景（y/n） [{'y' if config.transparent_background else 'n'}]: ",
        lambda text: parse_bool_text(text, config.transparent_background),
    )
    config.mode_overrides = prompt_until_valid(
        "请输入工作模式映射补充（例如 恒流充电=charge,静置=rest；直接回车跳过）: ",
        parse_mode_overrides,
    )
    sync_output_settings(config)
    return config


def run(config: AppConfig, use_demo: bool) -> Path:
    if use_demo:
        dataset = generate_demo_dataset()
        data = dataset.data
        print("已加载内置示例数据。")
    else:
        if config.input_path is None:
            raise UserInputError("缺少输入文件路径。")
        dataset = load_battery_dataset(
            file_path=config.input_path,
            sheet_name=config.sheet_name,
            cycles=config.cycles,
            mode_overrides=config.mode_overrides,
            auto_sort=config.auto_sort,
            absolute_specific_capacity=config.absolute_specific_capacity,
        )
        data = dataset.data
        print(f"已识别 sheet: {dataset.sheet_name}")
        print(f"已识别表头行: 第 {dataset.header_row + 1} 行")

    saved_path = plot_voltage_specific_capacity(data, config)
    return saved_path


def main() -> int:
    parser = build_argument_parser()
    args = parser.parse_args()

    try:
        config = apply_cli_arguments(args)
        interactive_mode = args.interactive or len(sys.argv) == 1
        needs_required_prompt = (not args.demo) and (config.input_path is None or config.output_path is None)
        if interactive_mode or needs_required_prompt:
            config = configure_interactively(config, prompt_all=interactive_mode, use_demo=args.demo)

        if config.output_path is None:
            config.output_path = build_output_path("battery_curve.png", config.output_format)
            sync_output_settings(config)

        saved_path = run(config, use_demo=args.demo)
        print(f"图片已保存到 {saved_path}")
        return 0
    except (UserInputError, DataLoaderError, PlottingError, ValueError) as exc:
        print(f"错误: {exc}")
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
