from __future__ import annotations

from pathlib import Path

import matplotlib as mpl
import matplotlib.pyplot as plt
import pandas as pd

from config import AppConfig


class PlottingError(RuntimeError):
    """Raised when no valid chart can be rendered."""


MODE_LABELS = {
    "charge": "Charge",
    "discharge": "Discharge",
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
    if len(cycles) == 1:
        return {cycles[0]: cmap(0.15)}
    return {cycle: cmap(index / max(len(cycles) - 1, 1)) for index, cycle in enumerate(cycles)}


def plot_voltage_specific_capacity(data: pd.DataFrame, config: AppConfig) -> Path:
    if data.empty:
        raise PlottingError("没有可绘制的数据。")
    if config.output_path is None:
        raise PlottingError("缺少输出路径。")

    apply_plot_theme(config)
    cycles = sorted(data["cycle"].dropna().astype(int).unique().tolist())
    colors = _resolve_cycle_colors(cycles, config)

    fig, ax = plt.subplots(figsize=config.figure_size())

    for cycle in cycles:
        cycle_df = data[data["cycle"] == cycle]
        for curve_type in ["charge", "discharge"]:
            segment = cycle_df[cycle_df["curve_type"] == curve_type]
            if segment.empty:
                continue
            label = f"Cycle {cycle} - {MODE_LABELS[curve_type]}"
            line_style = config.line_style_charge if curve_type == "charge" else config.line_style_discharge
            ax.plot(
                segment["specific_capacity"],
                segment["voltage"],
                color=colors.get(cycle, config.line_color),
                linewidth=config.line_width,
                linestyle=line_style,
                label=label,
            )

    ax.set_title(config.title, fontsize=config.title_fontsize)
    ax.set_xlabel(config.x_label, fontsize=config.label_fontsize)
    ax.set_ylabel(config.y_label, fontsize=config.label_fontsize)
    ax.tick_params(axis="both", labelsize=config.tick_fontsize)

    if config.grid:
        ax.grid(True, linestyle=config.grid_style, linewidth=0.8)
    else:
        ax.grid(False)

    if config.x_lim:
        ax.set_xlim(*config.x_lim)
    if config.y_lim:
        ax.set_ylim(*config.y_lim)

    if config.show_legend:
        ax.legend(loc=config.legend_loc, fontsize=max(config.tick_fontsize - 1, 8), ncol=1)

    fig.tight_layout()
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
