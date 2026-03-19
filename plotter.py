from __future__ import annotations

from pathlib import Path

import matplotlib as mpl
import matplotlib.pyplot as plt
import pandas as pd
from matplotlib.lines import Line2D

from config import AppConfig
from plot_catalog import build_cycle_summary, build_differential_dataset, resolve_plot_text


class PlottingError(RuntimeError):
    """Raised when no valid chart can be rendered."""


MODE_LABELS = {
    "charge": "Charge",
    "discharge": "Discharge",
}


SUMMARY_COLORS = {
    "charge_capacity": "#c44e52",
    "discharge_capacity": "#4c72b0",
    "coulombic_efficiency": "#2f4858",
    "rate_capacity": "#2a9d8f",
    "current_density": "#6c757d",
}


def apply_plot_theme(config: AppConfig) -> None:
    font_family = config.font_family or []
    mpl.rcParams["font.sans-serif"] = font_family
    mpl.rcParams["font.family"] = "sans-serif"
    mpl.rcParams["axes.unicode_minus"] = False

    if config.theme == "paper":
        mpl.rcParams.update(
            {
                "axes.linewidth": 1.2,
                "axes.labelweight": "normal",
                "xtick.direction": "in",
                "ytick.direction": "in",
                "grid.alpha": 0.35,
                "legend.frameon": False,
                "savefig.bbox": "tight",
            }
        )
    elif config.theme == "default":
        mpl.rcParams.update(
            {
                "axes.linewidth": 1.0,
                "xtick.direction": "out",
                "ytick.direction": "out",
                "grid.alpha": 0.25,
                "legend.frameon": True,
                "savefig.bbox": "tight",
            }
        )


def _resolve_cycle_colors(cycles: list[int], config: AppConfig) -> dict[int, tuple[float, float, float, float] | str]:
    if not cycles:
        return {}
    if not config.color_by_cycle:
        return {cycle: config.line_color for cycle in cycles}

    cmap = plt.get_cmap(config.colormap)
    lower = 0.12
    upper = 0.88
    if len(cycles) == 1:
        return {cycles[0]: cmap((lower + upper) / 2)}
    return {
        cycle: cmap(lower + (upper - lower) * index / max(len(cycles) - 1, 1))
        for index, cycle in enumerate(cycles)
    }


def _style_axes(ax, config: AppConfig) -> None:
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.spines["left"].set_linewidth(1.4 if config.theme == "paper" else 1.1)
    ax.spines["bottom"].set_linewidth(1.4 if config.theme == "paper" else 1.1)
    ax.tick_params(
        axis="both",
        which="major",
        direction="in",
        length=5.5,
        width=1.1,
        labelsize=config.tick_fontsize,
    )
    ax.tick_params(axis="both", which="minor", direction="in", length=3, width=0.9)
    ax.minorticks_on()
    ax.set_facecolor("white")
    ax.margins(x=0.02, y=0.04)


def _apply_common_labels(ax, config: AppConfig, title: str, x_label: str, y_label: str) -> None:
    ax.set_title(title, fontsize=config.title_fontsize)
    ax.set_xlabel(x_label, fontsize=config.label_fontsize)
    ax.set_ylabel(y_label, fontsize=config.label_fontsize)
    _style_axes(ax, config)


def _save_figure(fig, config: AppConfig) -> Path:
    if config.output_path is None:
        raise PlottingError("缺少输出路径。")
    save_kwargs = {
        "dpi": config.dpi,
        "transparent": config.transparent_background,
        "format": config.normalized_output_format(),
    }
    config.output_path.parent.mkdir(parents=True, exist_ok=True)
    fig.savefig(config.output_path, **save_kwargs)
    if config.show_after_save:
        plt.show()
    else:
        plt.close(fig)
    return config.output_path


def _build_compact_cycle_legend(ax, cycles: list[int], colors: dict[int, tuple[float, float, float, float] | str], config: AppConfig) -> None:
    if not config.show_legend or not cycles:
        return

    cycle_handles = [
        Line2D([0], [0], color=colors.get(cycle, config.line_color), linewidth=config.line_width, linestyle="-")
        for cycle in cycles
    ]
    cycle_labels = [f"Cycle {cycle}" for cycle in cycles]
    style_handles = [
        Line2D([0], [0], color="#444444", linewidth=config.line_width, linestyle=config.line_style_charge),
        Line2D([0], [0], color="#444444", linewidth=config.line_width, linestyle=config.line_style_discharge),
    ]
    style_labels = ["Charge", "Discharge"]

    handles = cycle_handles + style_handles
    labels = cycle_labels + style_labels
    ax.legend(
        handles,
        labels,
        loc=config.legend_loc,
        fontsize=max(config.tick_fontsize - 1, 9),
        ncol=1 if len(cycles) <= 6 else 2,
        frameon=False,
        handlelength=2.6,
        borderpad=0.2,
        labelspacing=0.45,
    )


def _build_standard_legend(ax, ax2, config: AppConfig) -> None:
    if not config.show_legend:
        return
    handles, labels = ax.get_legend_handles_labels()
    if ax2 is not None:
        handles2, labels2 = ax2.get_legend_handles_labels()
        handles += handles2
        labels += labels2
    ax.legend(
        handles,
        labels,
        loc=config.legend_loc,
        fontsize=max(config.tick_fontsize - 1, 9),
        frameon=False,
        handlelength=2.4,
        borderpad=0.2,
        labelspacing=0.45,
    )


def plot_voltage_specific_capacity(data: pd.DataFrame, config: AppConfig, *, total_plot_types: int = 1) -> Path:
    if data.empty:
        raise PlottingError("没有可绘制的数据。")
    if config.output_path is None:
        raise PlottingError("缺少输出路径。")

    apply_plot_theme(config)
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

    fig, ax = plt.subplots(figsize=config.figure_size())
    fig.patch.set_facecolor("white")

    for cycle in cycles:
        cycle_df = data[data["cycle"] == cycle]
        for curve_type in ["charge", "discharge"]:
            segment = cycle_df[cycle_df["curve_type"] == curve_type]
            if segment.empty:
                continue
            line_style = config.line_style_charge if curve_type == "charge" else config.line_style_discharge
            ax.plot(
                segment["specific_capacity"],
                segment["voltage"],
                color=colors.get(cycle, config.line_color),
                linewidth=config.line_width,
                linestyle=line_style,
            )

    _apply_common_labels(ax, config, text.title, text.x_label, text.y_label)
    if config.grid:
        ax.grid(True, linestyle=config.grid_style, linewidth=0.8)
    else:
        ax.grid(False)
    if config.x_lim:
        ax.set_xlim(*config.x_lim)
    if config.y_lim:
        ax.set_ylim(*config.y_lim)

    _build_compact_cycle_legend(ax, cycles, colors, config)
    fig.tight_layout(pad=0.9)
    return _save_figure(fig, config)


def _plot_differential(data: pd.DataFrame, config: AppConfig, plot_type: str, *, total_plot_types: int = 1) -> Path:
    differential = build_differential_dataset(data, plot_type)
    if differential.empty:
        raise PlottingError(f"没有足够的数据生成 {plot_type} 图。")

    apply_plot_theme(config)
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

    fig, ax = plt.subplots(figsize=config.figure_size())
    fig.patch.set_facecolor("white")

    for cycle in cycles:
        cycle_df = differential[differential["cycle"] == cycle]
        for curve_type in ["charge", "discharge"]:
            segment = cycle_df[cycle_df["curve_type"] == curve_type]
            if segment.empty:
                continue
            line_style = config.line_style_charge if curve_type == "charge" else config.line_style_discharge
            ax.plot(
                segment["x"],
                segment["y"],
                color=colors.get(cycle, config.line_color),
                linewidth=config.line_width,
                linestyle=line_style,
            )

    _apply_common_labels(ax, config, text.title, text.x_label, text.y_label)
    if config.grid:
        ax.grid(True, linestyle=config.grid_style, linewidth=0.8)
    else:
        ax.grid(False)
    if config.x_lim:
        ax.set_xlim(*config.x_lim)
    if config.y_lim:
        ax.set_ylim(*config.y_lim)

    _build_compact_cycle_legend(ax, cycles, colors, config)
    fig.tight_layout(pad=0.9)
    return _save_figure(fig, config)


def plot_dqdv(data: pd.DataFrame, config: AppConfig, *, total_plot_types: int = 1) -> Path:
    return _plot_differential(data, config, "dqdv", total_plot_types=total_plot_types)


def plot_dvdq(data: pd.DataFrame, config: AppConfig, *, total_plot_types: int = 1) -> Path:
    return _plot_differential(data, config, "dvdq", total_plot_types=total_plot_types)


def plot_long_cycling(data: pd.DataFrame, config: AppConfig, *, total_plot_types: int = 1) -> Path:
    summary = build_cycle_summary(data)
    if summary.empty:
        raise PlottingError("没有足够的数据生成循环性能图。")

    apply_plot_theme(config)
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

    fig, ax = plt.subplots(figsize=config.figure_size())
    fig.patch.set_facecolor("white")

    cycles = summary["cycle"].astype(int).to_numpy()
    if summary["charge_specific_capacity"].notna().any():
        ax.plot(
            cycles,
            summary["charge_specific_capacity"],
            marker="o",
            markersize=5.5,
            linewidth=config.line_width,
            color=SUMMARY_COLORS["charge_capacity"],
            label="Charge Capacity",
        )
    if summary["discharge_specific_capacity"].notna().any():
        ax.plot(
            cycles,
            summary["discharge_specific_capacity"],
            marker="o",
            markersize=5.5,
            linewidth=config.line_width,
            color=SUMMARY_COLORS["discharge_capacity"],
            label="Discharge Capacity",
        )

    ax2 = None
    if "coulombic_efficiency" in summary.columns and summary["coulombic_efficiency"].notna().any():
        ax2 = ax.twinx()
        ax2.plot(
            cycles,
            summary["coulombic_efficiency"],
            marker="s",
            markersize=5,
            linewidth=max(config.line_width - 0.2, 1.4),
            linestyle="--",
            color=SUMMARY_COLORS["coulombic_efficiency"],
            label="Coulombic Efficiency",
        )
        ax2.set_ylabel(text.y2_label or "Coulombic Efficiency (%)", fontsize=config.label_fontsize)
        ax2.tick_params(axis="y", labelsize=config.tick_fontsize, direction="in", length=5, width=1.0)
        ax2.spines["top"].set_visible(False)
        ax2.spines["right"].set_linewidth(1.2)
        ce_values = summary["coulombic_efficiency"].dropna()
        if not ce_values.empty:
            lower = max(0, min(95, float(ce_values.min()) - 2))
            upper = max(102, float(ce_values.max()) + 2)
            if upper - lower < 8:
                lower = max(0, upper - 8)
            ax2.set_ylim(lower, upper)

    _apply_common_labels(ax, config, text.title, text.x_label, text.y_label)
    ax.set_xticks(cycles)
    if config.grid:
        ax.grid(True, linestyle=config.grid_style, linewidth=0.8)
    else:
        ax.grid(False)
    if config.x_lim:
        ax.set_xlim(*config.x_lim)
    if config.y_lim:
        ax.set_ylim(*config.y_lim)

    _build_standard_legend(ax, ax2, config)
    fig.tight_layout(pad=0.9)
    return _save_figure(fig, config)


def plot_rate_capability(data: pd.DataFrame, config: AppConfig, *, total_plot_types: int = 1) -> Path:
    summary = build_cycle_summary(data)
    if summary.empty:
        raise PlottingError("没有足够的数据生成倍率性能图。")

    apply_plot_theme(config)
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

    fig, ax = plt.subplots(figsize=config.figure_size())
    fig.patch.set_facecolor("white")

    cycles = summary["cycle"].astype(int).to_numpy()
    discharge_capacity = summary["discharge_specific_capacity"]
    if discharge_capacity.notna().sum() == 0:
        discharge_capacity = summary["charge_specific_capacity"]
    if discharge_capacity.notna().sum() == 0:
        raise PlottingError("没有足够的容量数据生成倍率性能图。")

    ax.plot(
        cycles,
        discharge_capacity,
        marker="o",
        markersize=5.5,
        linewidth=config.line_width,
        color=SUMMARY_COLORS["rate_capacity"],
        label="Specific Capacity",
    )

    ax2 = None
    rate_label = text.y2_label or "Current Density (mA/g)"
    if "current_density" in summary.columns and summary["current_density"].notna().any():
        rate_values = summary["current_density"]
    elif "discharge_current" in summary.columns and summary["discharge_current"].notna().any():
        rate_values = summary["discharge_current"].abs()
        rate_label = "Current (mA)"
    elif "charge_current" in summary.columns and summary["charge_current"].notna().any():
        rate_values = summary["charge_current"].abs()
        rate_label = "Current (mA)"
    else:
        rate_values = None

    if rate_values is not None:
        ax2 = ax.twinx()
        ax2.step(
            cycles,
            rate_values,
            where="mid",
            linewidth=max(config.line_width - 0.2, 1.4),
            linestyle="--",
            color=SUMMARY_COLORS["current_density"],
            label=rate_label,
        )
        ax2.scatter(cycles, rate_values, s=28, color=SUMMARY_COLORS["current_density"], zorder=3)
        ax2.set_ylabel(rate_label, fontsize=config.label_fontsize)
        ax2.tick_params(axis="y", labelsize=config.tick_fontsize, direction="in", length=5, width=1.0)
        ax2.spines["top"].set_visible(False)
        ax2.spines["right"].set_linewidth(1.2)

    _apply_common_labels(ax, config, text.title, text.x_label, text.y_label)
    ax.set_xticks(cycles)
    if config.grid:
        ax.grid(True, linestyle=config.grid_style, linewidth=0.8)
    else:
        ax.grid(False)
    if config.x_lim:
        ax.set_xlim(*config.x_lim)
    if config.y_lim:
        ax.set_ylim(*config.y_lim)

    _build_standard_legend(ax, ax2, config)
    fig.tight_layout(pad=0.9)
    return _save_figure(fig, config)


def plot_dataset_by_type(data: pd.DataFrame, config: AppConfig, plot_type: str, *, total_plot_types: int = 1) -> Path:
    if plot_type == "voltage_capacity":
        return plot_voltage_specific_capacity(data, config, total_plot_types=total_plot_types)
    if plot_type == "long_cycling":
        return plot_long_cycling(data, config, total_plot_types=total_plot_types)
    if plot_type == "rate_capability":
        return plot_rate_capability(data, config, total_plot_types=total_plot_types)
    if plot_type == "dqdv":
        return plot_dqdv(data, config, total_plot_types=total_plot_types)
    if plot_type == "dvdq":
        return plot_dvdq(data, config, total_plot_types=total_plot_types)
    raise PlottingError(f"不支持的绘图类型: {plot_type}")
