from __future__ import annotations

import re
from pathlib import Path

import matplotlib.pyplot as plt
import pandas as pd

from config import AppConfig
from plot_catalog import build_cycle_summary, build_differential_dataset, resolve_plot_text
from plotter import MODE_LABELS, PlottingError, SUMMARY_COLORS

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


def _style_axis_titles(layer, x_label: str, y_label: str, config: AppConfig) -> None:
    layer.axis("x").title = x_label
    layer.axis("y").title = y_label
    axis_title_size = max(config.label_fontsize * 1.7, 18)
    _style_label(layer.label("xb"), axis_title_size)
    _style_label(layer.label("yl"), axis_title_size)


def _set_graph_title(layer, title: str, config: AppConfig) -> None:
    if not title.strip():
        return
    layer.lt_exec(f"label -s {_safe_labtalk_text(title)}")
    title_label = layer.label("Title")
    _style_label(title_label, max(config.title_fontsize * 1.4, 20))
    if title_label is not None:
        try:
            title_label.set_int("attach", 2)
            x_from, x_to, _ = layer.axis("x").limits
            y_from, y_to, _ = layer.axis("y").limits
            title_label.set_float("x1", x_from + 0.42 * (x_to - x_from))
            title_label.set_float("y1", y_from + 0.97 * (y_to - y_from))
        except Exception:
            pass


def _style_grid(layer, config: AppConfig) -> None:
    if config.grid:
        layer.lt_exec("layer.x.grid=1;layer.y.grid=1;")
    else:
        layer.lt_exec("layer.x.grid=0;layer.y.grid=0;")


def _sanitize_dataframe_for_origin(df: pd.DataFrame) -> pd.DataFrame:
    working = df.copy()
    used: dict[str, int] = {}
    columns: list[str] = []
    for index, column in enumerate(working.columns):
        name = str(column).strip() or f"blank_{index + 1}"
        if name in used:
            used[name] += 1
            name = f"{name}_{used[name]}"
        else:
            used[name] = 1
        columns.append(name)
    working.columns = columns
    return working


def _write_dataframe(wks, df: pd.DataFrame, long_name: str) -> None:
    clean_df = _sanitize_dataframe_for_origin(df).reset_index(drop=True)
    wks.from_df(clean_df)
    try:
        wks.get_book().lname = long_name
    except Exception:
        pass


def _apply_plot_style(plot, config: AppConfig, color: tuple[int, int, int] | str, line_style: str) -> None:
    plot.color = color
    plot.transparency = 0
    plot.set_float("line.width", max(config.line_width, 0.1))
    style_code = ORIGIN_LINE_STYLE_MAP.get(line_style)
    if style_code is not None:
        plot.set_cmd(f"-d {style_code}")


def _set_manual_cycle_legend(
    layer,
    cycles: list[int],
    first_plot_index_by_cycle: dict[int, int],
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
    lines.append("")
    lines.append("Solid: Charge")
    lines.append("Dash: Discharge")
    legend.text = "\n".join(lines)
    _style_label(legend, max(config.tick_fontsize * 1.3, 14))

    x_from, x_to, _ = layer.axis("x").limits
    y_from, y_to, _ = layer.axis("y").limits
    if x_to > x_from and y_to > y_from:
        legend.set_float("x1", x_from + 0.72 * (x_to - x_from))
        legend.set_float("y1", y_from + 0.92 * (y_to - y_from))


def _style_standard_legend(layer, config: AppConfig, x_ratio: float = 0.72) -> None:
    if not config.show_legend:
        layer.lt_exec("legend -d")
        return
    layer.lt_exec("legend -r")
    legend = layer.label("Legend")
    if legend is None:
        return
    _style_label(legend, max(config.tick_fontsize * 1.3, 14))
    x_from, x_to, _ = layer.axis("x").limits
    y_from, y_to, _ = layer.axis("y").limits
    if x_to > x_from and y_to > y_from:
        legend.set_float("x1", x_from + x_ratio * (x_to - x_from))
        legend.set_float("y1", y_from + 0.92 * (y_to - y_from))


def _move_label_in_data_coords(layer, label_name: str, x_ratio: float, y_ratio: float) -> None:
    label = layer.label(label_name)
    if label is None:
        return
    try:
        x_from, x_to, _ = layer.axis("x").limits
        y_from, y_to, _ = layer.axis("y").limits
        label.set_float("x1", x_from + x_ratio * (x_to - x_from))
        label.set_float("y1", y_from + y_ratio * (y_to - y_from))
    except Exception:
        pass


def _finalize_origin_output(graph, visible_books: list, config: AppConfig) -> Path:
    if config.output_path is None:
        raise PlottingError("缺少输出路径。")
    resolved_output_path = config.output_path.resolve()
    resolved_output_path.parent.mkdir(parents=True, exist_ok=True)
    graph.save_fig(str(resolved_output_path), width=_origin_export_width(config))

    if config.save_origin_project:
        for book in visible_books:
            try:
                book.show = True
                book.activate()
            except Exception:
                pass
        graph.show = True
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


def _new_origin_context(data: pd.DataFrame, config: AppConfig) -> tuple:
    if op is None:
        raise PlottingError("当前环境未安装 originpro，无法使用 Origin 绘图后端。")
    if data.empty:
        raise PlottingError("没有可绘制的数据。")
    if config.output_path is None:
        raise PlottingError("缺少输出路径。")

    op.new(asksave=False)
    op.set_show(config.show_after_save)

    raw_wks = op.find_sheet("w")
    if raw_wks is None:
        raw_wks = op.new_sheet("w", hidden=False)
    raw_book = raw_wks.get_book()
    raw_book.show = True

    plot_wks = op.new_sheet("w", hidden=False)
    plot_book = plot_wks.get_book()
    plot_book.show = True

    graph = op.new_graph(template="line", hidden=False)
    graph.show = True
    layer = graph[0]

    stem = _resolved_output_stem(config.output_path)
    try:
        raw_book.lname = f"RawData_{stem}"
        plot_book.lname = f"PlotData_{stem}"
        graph.lname = f"OriginPlot_{stem}"
    except Exception:
        pass

    _write_dataframe(raw_wks, data, f"RawData_{stem}")
    return raw_book, plot_wks, plot_book, graph, layer


def plot_voltage_specific_capacity_origin(data: pd.DataFrame, config: AppConfig, *, total_plot_types: int = 1) -> Path:
    cycles = sorted(data["cycle"].dropna().astype(int).unique().tolist())
    colors = _resolve_cycle_colors(cycles, config)
    text = resolve_plot_text(
        "voltage_capacity",
        custom_title=config.title,
        custom_x_label=config.x_label,
        custom_y_label=config.y_label,
        total_plot_types=total_plot_types,
        default_title="Battery Charge-Discharge Curves",
        default_x_label="Specific Capacity (mAh/g)",
        default_y_label="Voltage (V)",
    )

    raw_book = plot_wks = plot_book = graph = layer = None
    should_exit_origin = not config.show_after_save
    try:
        raw_book, plot_wks, plot_book, graph, layer = _new_origin_context(data, config)

        next_col = 0
        plot_count = 0
        first_plot_index_by_cycle: dict[int, int] = {}

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
                plot_wks.from_list(next_col, x_values, lname=f"X_{cycle}_{curve_type}")
                plot_wks.from_list(next_col + 1, y_values, lname=label)
                plot = layer.add_plot(plot_wks, coly=next_col + 1, colx=next_col, type="l")
                line_style = config.line_style_charge if curve_type == "charge" else config.line_style_discharge
                _apply_plot_style(plot, config, colors.get(cycle, config.line_color), line_style)
                plot_count += 1
                if cycle not in first_plot_index_by_cycle:
                    first_plot_index_by_cycle[cycle] = plot_count
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
            layer.set_xlim(*config.x_lim)
        if config.y_lim:
            layer.set_ylim(*config.y_lim)
        _style_axis_titles(layer, text.x_label, text.y_label, config)
        _set_graph_title(layer, text.title, config)
        _style_grid(layer, config)
        _set_manual_cycle_legend(layer, cycles, first_plot_index_by_cycle, config)
        return _finalize_origin_output(graph, [raw_book, plot_book], config)
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


def _plot_differential_origin(data: pd.DataFrame, config: AppConfig, plot_type: str, *, total_plot_types: int = 1) -> Path:
    differential = build_differential_dataset(data, plot_type)
    if differential.empty:
        raise PlottingError(f"没有足够的数据生成 {plot_type} 图。")

    cycles = sorted(differential["cycle"].dropna().astype(int).unique().tolist())
    colors = _resolve_cycle_colors(cycles, config)
    text = resolve_plot_text(
        plot_type,
        custom_title=config.title,
        custom_x_label=config.x_label,
        custom_y_label=config.y_label,
        total_plot_types=total_plot_types,
        default_title="Battery Charge-Discharge Curves",
        default_x_label="Specific Capacity (mAh/g)",
        default_y_label="Voltage (V)",
    )

    raw_book = plot_wks = plot_book = graph = layer = None
    should_exit_origin = not config.show_after_save
    try:
        raw_book, plot_wks, plot_book, graph, layer = _new_origin_context(data, config)
        next_col = 0
        plot_count = 0
        first_plot_index_by_cycle: dict[int, int] = {}
        for cycle in cycles:
            cycle_df = differential[differential["cycle"] == cycle]
            for curve_type in ["charge", "discharge"]:
                segment = cycle_df[cycle_df["curve_type"] == curve_type]
                if segment.empty:
                    continue
                label = f"Cycle {cycle} - {MODE_LABELS[curve_type]}"
                plot_wks.from_list(next_col, segment["x"].astype(float).tolist(), lname=f"X_{cycle}_{curve_type}")
                plot_wks.from_list(next_col + 1, segment["y"].astype(float).tolist(), lname=label)
                plot = layer.add_plot(plot_wks, coly=next_col + 1, colx=next_col, type="l")
                line_style = config.line_style_charge if curve_type == "charge" else config.line_style_discharge
                _apply_plot_style(plot, config, colors.get(cycle, config.line_color), line_style)
                plot_count += 1
                if cycle not in first_plot_index_by_cycle:
                    first_plot_index_by_cycle[cycle] = plot_count
                next_col += 2

        if plot_count == 0:
            raise PlottingError(f"Origin 后端没有生成 {plot_type} 曲线。")

        layer.rescale()
        if config.x_lim:
            layer.set_xlim(*config.x_lim)
        if config.y_lim:
            layer.set_ylim(*config.y_lim)
        _style_axis_titles(layer, text.x_label, text.y_label, config)
        _set_graph_title(layer, text.title, config)
        _style_grid(layer, config)
        _set_manual_cycle_legend(layer, cycles, first_plot_index_by_cycle, config)
        return _finalize_origin_output(graph, [raw_book, plot_book], config)
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


def plot_dqdv_origin(data: pd.DataFrame, config: AppConfig, *, total_plot_types: int = 1) -> Path:
    return _plot_differential_origin(data, config, "dqdv", total_plot_types=total_plot_types)


def plot_dvdq_origin(data: pd.DataFrame, config: AppConfig, *, total_plot_types: int = 1) -> Path:
    return _plot_differential_origin(data, config, "dvdq", total_plot_types=total_plot_types)


def plot_long_cycling_origin(data: pd.DataFrame, config: AppConfig, *, total_plot_types: int = 1) -> Path:
    summary = build_cycle_summary(data)
    if summary.empty:
        raise PlottingError("没有足够的数据生成循环性能图。")

    text = resolve_plot_text(
        "long_cycling",
        custom_title=config.title,
        custom_x_label=config.x_label,
        custom_y_label=config.y_label,
        total_plot_types=total_plot_types,
        default_title="Battery Charge-Discharge Curves",
        default_x_label="Specific Capacity (mAh/g)",
        default_y_label="Voltage (V)",
    )
    plot_df = pd.DataFrame({"cycle": summary["cycle"].astype(int)})
    left_series_specs: list[tuple[str, str, str]] = []
    right_series_specs: list[tuple[str, str, str]] = []
    if summary["charge_specific_capacity"].notna().any():
        plot_df["Charge Capacity"] = summary["charge_specific_capacity"]
        left_series_specs.append(("Charge Capacity", SUMMARY_COLORS["charge_capacity"], "-"))
    if summary["discharge_specific_capacity"].notna().any():
        plot_df["Discharge Capacity"] = summary["discharge_specific_capacity"]
        left_series_specs.append(("Discharge Capacity", SUMMARY_COLORS["discharge_capacity"], "-"))
    if summary["coulombic_efficiency"].notna().any():
        plot_df["Coulombic Efficiency"] = summary["coulombic_efficiency"]
        right_series_specs.append(("Coulombic Efficiency", SUMMARY_COLORS["coulombic_efficiency"], "--"))
    if len(left_series_specs) == 0:
        raise PlottingError("没有足够的数据生成循环性能图。")

    raw_book = plot_wks = plot_book = graph = layer = right_layer = None
    should_exit_origin = not config.show_after_save
    try:
        raw_book, plot_wks, plot_book, graph, layer = _new_origin_context(data, config)
        _write_dataframe(plot_wks, plot_df, f"Cycling_{_resolved_output_stem(config.output_path)}")

        for index, (column_name, color, line_style) in enumerate(left_series_specs, start=1):
            plot = layer.add_plot(plot_wks, coly=index, colx=0, type="l")
            _apply_plot_style(plot, config, color, line_style)

        right_layer = None
        if right_series_specs:
            right_layer = graph.add_layer(2)
            for column_name, color, line_style in right_series_specs:
                col_index = list(plot_df.columns).index(column_name)
                plot = right_layer.add_plot(plot_wks, coly=col_index, colx=0, type="l")
                _apply_plot_style(plot, config, color, line_style)

        layer.rescale()
        if right_layer is not None:
            right_layer.rescale()
            try:
                right_layer.axis("x").title = ""
            except Exception:
                pass
            try:
                right_layer.axis("y").title = text.y2_label or "Coulombic Efficiency (%)"
            except Exception:
                pass
            _style_label(right_layer.label("yr"), max(config.label_fontsize * 1.7, 18))
            ce_values = plot_df["Coulombic Efficiency"].dropna() if "Coulombic Efficiency" in plot_df.columns else pd.Series(dtype=float)
            if not ce_values.empty:
                lower = max(0, min(95, float(ce_values.min()) - 2))
                upper = max(102, float(ce_values.max()) + 2)
                if upper - lower < 8:
                    lower = max(0, upper - 8)
                try:
                    right_layer.set_ylim(lower, upper)
                except Exception:
                    pass

        _style_axis_titles(layer, text.x_label, text.y_label, config)
        _set_graph_title(layer, text.title, config)
        _move_label_in_data_coords(layer, "Title", 0.14, 0.88)
        _style_grid(layer, config)
        try:
            layer.axis("x").sstep = 1
        except Exception:
            pass
        _style_standard_legend(layer, config, x_ratio=0.78)
        _move_label_in_data_coords(layer, "Legend", 0.72, 0.96)
        if right_layer is not None:
            try:
                right_layer.lt_exec("legend -r")
                right_legend = right_layer.label("Legend")
                if right_legend is not None:
                    right_legend.text = "\\l(1) Coulombic Efficiency"
                    _style_label(right_legend, max(config.tick_fontsize * 1.25, 13))
                    x_from, x_to, _ = right_layer.axis("x").limits
                    y_from, y_to, _ = right_layer.axis("y").limits
                    right_legend.set_float("x1", x_from + 0.72 * (x_to - x_from))
                    right_legend.set_float("y1", y_from + 0.72 * (y_to - y_from))
            except Exception:
                pass
        return _finalize_origin_output(graph, [raw_book, plot_book], config)
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


def plot_rate_capability_origin(data: pd.DataFrame, config: AppConfig, *, total_plot_types: int = 1) -> Path:
    summary = build_cycle_summary(data)
    if summary.empty:
        raise PlottingError("没有足够的数据生成倍率性能图。")

    text = resolve_plot_text(
        "rate_capability",
        custom_title=config.title,
        custom_x_label=config.x_label,
        custom_y_label=config.y_label,
        total_plot_types=total_plot_types,
        default_title="Battery Charge-Discharge Curves",
        default_x_label="Specific Capacity (mAh/g)",
        default_y_label="Voltage (V)",
    )

    plot_df = pd.DataFrame({"cycle": summary["cycle"].astype(int)})
    series_specs: list[tuple[str, str, str]] = []

    if summary["discharge_specific_capacity"].notna().any():
        plot_df["Specific Capacity"] = summary["discharge_specific_capacity"]
    elif summary["charge_specific_capacity"].notna().any():
        plot_df["Specific Capacity"] = summary["charge_specific_capacity"]
    else:
        raise PlottingError("没有足够的容量数据生成倍率性能图。")
    series_specs.append(("Specific Capacity", SUMMARY_COLORS["rate_capacity"], "-"))

    y_label = text.y_label
    if summary["current_density"].notna().any():
        plot_df["Current Density"] = summary["current_density"]
        series_specs.append(("Current Density", SUMMARY_COLORS["current_density"], "--"))
        y_label = "Specific Capacity / Current Density"
    elif summary["discharge_current"].notna().any():
        plot_df["Current"] = summary["discharge_current"].abs()
        series_specs.append(("Current", SUMMARY_COLORS["current_density"], "--"))
        y_label = "Specific Capacity / Current"
    elif summary["charge_current"].notna().any():
        plot_df["Current"] = summary["charge_current"].abs()
        series_specs.append(("Current", SUMMARY_COLORS["current_density"], "--"))
        y_label = "Specific Capacity / Current"

    raw_book = plot_wks = plot_book = graph = layer = None
    should_exit_origin = not config.show_after_save
    try:
        raw_book, plot_wks, plot_book, graph, layer = _new_origin_context(data, config)
        _write_dataframe(plot_wks, plot_df, f"Rate_{_resolved_output_stem(config.output_path)}")

        for index, (column_name, color, line_style) in enumerate(series_specs, start=1):
            plot = layer.add_plot(plot_wks, coly=index, colx=0, type="l")
            _apply_plot_style(plot, config, color, line_style)

        layer.rescale()
        _style_axis_titles(layer, text.x_label, y_label, config)
        _set_graph_title(layer, text.title, config)
        _style_grid(layer, config)
        try:
            layer.axis("x").sstep = 1
        except Exception:
            pass
        _style_standard_legend(layer, config)
        return _finalize_origin_output(graph, [raw_book, plot_book], config)
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


def plot_dataset_by_type_origin(data: pd.DataFrame, config: AppConfig, plot_type: str, *, total_plot_types: int = 1) -> Path:
    if plot_type == "voltage_capacity":
        return plot_voltage_specific_capacity_origin(data, config, total_plot_types=total_plot_types)
    if plot_type == "long_cycling":
        return plot_long_cycling_origin(data, config, total_plot_types=total_plot_types)
    if plot_type == "rate_capability":
        return plot_rate_capability_origin(data, config, total_plot_types=total_plot_types)
    if plot_type == "dqdv":
        return plot_dqdv_origin(data, config, total_plot_types=total_plot_types)
    if plot_type == "dvdq":
        return plot_dvdq_origin(data, config, total_plot_types=total_plot_types)
    raise PlottingError(f"不支持的绘图类型: {plot_type}")
