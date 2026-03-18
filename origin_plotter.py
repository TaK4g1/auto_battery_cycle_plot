from __future__ import annotations

import re
from pathlib import Path

import matplotlib.pyplot as plt
import pandas as pd

from config import AppConfig
from plotter import MODE_LABELS, PlottingError

try:
    import originpro as op
except Exception:  # pragma: no cover - runtime dependency
    op = None


ORIGIN_LINE_STYLE_MAP = {
    "-": None,
    ":": 2,
    "--": 3,
    "-.": 4,
}


def _resolve_cycle_colors(cycles: list[int], config: AppConfig) -> dict[int, tuple[int, int, int] | str]:
    if not cycles:
        return {}
    if not config.color_by_cycle:
        return {cycle: config.line_color for cycle in cycles}

    cmap = plt.get_cmap(config.colormap)
    lower = 0.12
    upper = 0.88
    if len(cycles) == 1:
        rgba = cmap((lower + upper) / 2)
        return {cycles[0]: (int(rgba[0] * 255), int(rgba[1] * 255), int(rgba[2] * 255))}

    colors: dict[int, tuple[int, int, int]] = {}
    for index, cycle in enumerate(cycles):
        rgba = cmap(lower + (upper - lower) * index / max(len(cycles) - 1, 1))
        colors[cycle] = (int(rgba[0] * 255), int(rgba[1] * 255), int(rgba[2] * 255))
    return colors


def _safe_labtalk_text(text: str) -> str:
    return '"' + re.sub(r'["\r\n]+', " ", text) + '"'


def _effective_title(config: AppConfig) -> str:
    title = (config.title or "").strip()
    if title == "Battery Charge-Discharge Curves":
        return ""
    return title


def _build_origin_project_path(output_path: Path) -> Path:
    return output_path.with_suffix(".opju").resolve()


def _resolved_output_stem(output_path: Path) -> str:
    return re.sub(r"[^\w\-\.]+", "_", output_path.stem) or "plot"


def _origin_export_width(config: AppConfig) -> int:
    return max(int(config.figure_width * config.dpi), 1800)


def _style_label(label, size: float) -> None:
    if label is None:
        return
    try:
        label.set_float("fsize", size)
    except Exception:
        pass
    try:
        label.color = "Black"
    except Exception:
        pass


def _style_axis_titles(layer, config: AppConfig) -> None:
    layer.axis("x").title = config.x_label
    layer.axis("y").title = config.y_label
    axis_title_size = max(config.label_fontsize * 1.7, 18)
    _style_label(layer.label("xb"), axis_title_size)
    _style_label(layer.label("yl"), axis_title_size)


def _set_graph_title(layer, title: str, config: AppConfig) -> None:
    if not title.strip():
        return
    layer.lt_exec(f"label -s {_safe_labtalk_text(title)}")
    _style_label(layer.label("Title"), max(config.title_fontsize * 1.4, 20))


def _style_grid(layer, config: AppConfig) -> None:
    if config.grid:
        layer.lt_exec("layer.x.grid=1;layer.y.grid=1;")
    else:
        layer.lt_exec("layer.x.grid=0;layer.y.grid=0;")


def _apply_plot_style(plot, config: AppConfig, curve_type: str, color: tuple[int, int, int] | str) -> None:
    plot.color = color
    plot.transparency = 0
    plot.set_float("line.width", max(config.line_width, 0.1))

    line_style = config.line_style_charge if curve_type == "charge" else config.line_style_discharge
    style_code = ORIGIN_LINE_STYLE_MAP.get(line_style)
    if style_code is not None:
        plot.set_cmd(f"-d {style_code}")


def _set_manual_legend(
    layer,
    cycles: list[int],
    first_plot_index_by_cycle: dict[int, int],
    has_charge: bool,
    has_discharge: bool,
    config: AppConfig,
) -> None:
    if not config.show_legend:
        layer.lt_exec("legend -d")
        return

    layer.lt_exec("legend -r")
    legend = layer.label("Legend")
    if legend is None:
        return

    lines = [f"\\l({first_plot_index_by_cycle[cycle]}) Cycle {cycle}" for cycle in cycles if cycle in first_plot_index_by_cycle]
    if has_charge or has_discharge:
        lines.append("")
        if has_charge:
            lines.append("Solid: Charge")
        if has_discharge:
            lines.append("Dash: Discharge")

    legend.text = "\n".join(lines)
    _style_label(legend, max(config.tick_fontsize * 1.35, 14))

    x_from, x_to, _ = layer.axis("x").limits
    y_from, y_to, _ = layer.axis("y").limits
    x_span = x_to - x_from
    y_span = y_to - y_from
    if x_span > 0 and y_span > 0:
        try:
            legend.set_float("x1", x_from + 0.73 * x_span)
            legend.set_float("y1", y_from + 0.92 * y_span)
        except Exception:
            pass


def plot_voltage_specific_capacity_origin(data: pd.DataFrame, config: AppConfig) -> Path:
    if op is None:
        raise PlottingError("当前环境未安装 originpro，无法使用 Origin 绘图后端。")
    if data.empty:
        raise PlottingError("没有可绘制的数据。")
    if config.output_path is None:
        raise PlottingError("缺少输出路径。")

    cycles = sorted(data["cycle"].dropna().astype(int).unique().tolist())
    colors = _resolve_cycle_colors(cycles, config)

    workbook = None
    graph = None
    should_exit_origin = not config.show_after_save
    try:
        op.new(asksave=False)
        op.set_show(config.show_after_save)

        workbook = op.find_sheet("w")
        if workbook is None:
            workbook = op.new_sheet("w", hidden=False)
        workbook_book = workbook.get_book()
        workbook_book.show = True

        graph = op.new_graph(template="line", hidden=False)
        graph.show = True
        layer = graph[0]

        try:
            workbook_book.lname = f"PlotData_{_resolved_output_stem(config.output_path)}"
        except Exception:
            pass
        try:
            graph.lname = f"OriginPlot_{_resolved_output_stem(config.output_path)}"
        except Exception:
            pass

        next_col = 0
        plot_count = 0
        first_plot_index_by_cycle: dict[int, int] = {}
        has_charge = False
        has_discharge = False

        for cycle in cycles:
            cycle_df = data[data["cycle"] == cycle]
            cycle_segment_count = 0

            for curve_type in ["charge", "discharge"]:
                segment = cycle_df[cycle_df["curve_type"] == curve_type]
                if segment.empty:
                    continue

                x_values = segment["specific_capacity"].astype(float).tolist()
                y_values = segment["voltage"].astype(float).tolist()
                label = f"Cycle {cycle} - {MODE_LABELS[curve_type]}"

                workbook.from_list(next_col, x_values, lname=f"X_{cycle}_{curve_type}")
                workbook.from_list(next_col + 1, y_values, lname=label)

                plot = layer.add_plot(workbook, coly=next_col + 1, colx=next_col, type="l")
                _apply_plot_style(plot, config, curve_type, colors.get(cycle, config.line_color))

                plot_count += 1
                if cycle not in first_plot_index_by_cycle:
                    first_plot_index_by_cycle[cycle] = plot_count
                if curve_type == "charge":
                    has_charge = True
                elif curve_type == "discharge":
                    has_discharge = True

                next_col += 2
                cycle_segment_count += 1

            if cycle_segment_count == 1:
                present_curve = cycle_df["curve_type"].dropna().iloc[0]
                missing_curve = "discharge" if present_curve == "charge" else "charge"
                print(f"提示: Cycle {cycle} 仅检测到 {present_curve} 数据，未检测到 {missing_curve} 数据。")

        if plot_count == 0:
            raise PlottingError("Origin 后端没有生成任何曲线。")

        layer.rescale()
        if config.x_lim:
            layer.set_xlim(config.x_lim[0], config.x_lim[1])
        if config.y_lim:
            layer.set_ylim(config.y_lim[0], config.y_lim[1])

        _style_axis_titles(layer, config)
        _set_graph_title(layer, _effective_title(config), config)
        _style_grid(layer, config)
        _set_manual_legend(layer, cycles, first_plot_index_by_cycle, has_charge, has_discharge, config)

        resolved_output_path = config.output_path.resolve()
        resolved_output_path.parent.mkdir(parents=True, exist_ok=True)
        graph.save_fig(str(resolved_output_path), width=_origin_export_width(config))

        if config.save_origin_project:
            workbook_book.show = True
            graph.show = True
            try:
                workbook_book.activate()
            except Exception:
                pass
            try:
                graph.activate()
            except Exception:
                pass

            project_path = _build_origin_project_path(resolved_output_path)
            op.save(str(project_path))
            print(f"Origin工程已保存到 {project_path}")
            if config.show_after_save:
                print("你可以直接在当前已打开的 Origin 中继续编辑，或之后双击 .opju 文件重新打开编辑。")

        return resolved_output_path
    except PlottingError:
        raise
    except Exception as exc:
        raise PlottingError(f"Origin 绘图失败: {exc}") from exc
    finally:
        if should_exit_origin:
            try:
                op.exit()
            except Exception:
                pass
