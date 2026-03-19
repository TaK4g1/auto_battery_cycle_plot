from __future__ import annotations

from dataclasses import dataclass
from typing import Iterable

import numpy as np
import pandas as pd


DEFAULT_PLOT_TYPE = "voltage_capacity"


PLOT_TYPE_REGISTRY: dict[str, dict[str, str]] = {
    "voltage_capacity": {
        "label": "GCD: Voltage-Capacity",
        "suffix": "gcd",
        "title": "Battery Charge-Discharge Curves",
        "x_label": "Specific Capacity (mAh/g)",
        "y_label": "Voltage (V)",
    },
    "long_cycling": {
        "label": "Long Cycling",
        "suffix": "long_cycling",
        "title": "Long-Term Cycling Performance",
        "x_label": "Cycle Number",
        "y_label": "Specific Capacity (mAh/g)",
        "y2_label": "Coulombic Efficiency (%)",
    },
    "rate_capability": {
        "label": "Rate Capability",
        "suffix": "rate",
        "title": "Rate Capability",
        "x_label": "Cycle Number",
        "y_label": "Discharge Specific Capacity (mAh/g)",
        "y2_label": "Current Density (mA/g)",
    },
    "dqdv": {
        "label": "dQ/dV Curves",
        "suffix": "dqdv",
        "title": "Differential Capacity Curves",
        "x_label": "Voltage (V)",
        "y_label": "dQ/dV (mAh/V)",
    },
    "dvdq": {
        "label": "dV/dQ Curves",
        "suffix": "dvdq",
        "title": "Differential Voltage Curves",
        "x_label": "Specific Capacity (mAh/g)",
        "y_label": "dV/dQ (V/mAh)",
    },
}


PLOT_TYPE_ALIASES = {
    "gcd": "voltage_capacity",
    "voltage-capacity": "voltage_capacity",
    "voltage_capacity": "voltage_capacity",
    "cycling": "long_cycling",
    "cycle": "long_cycling",
    "cycling_performance": "long_cycling",
    "long_cycling": "long_cycling",
    "rate": "rate_capability",
    "rate_capability": "rate_capability",
    "dqdv": "dqdv",
    "dq/dv": "dqdv",
    "dvdq": "dvdq",
    "dv/dq": "dvdq",
}


@dataclass
class PlotText:
    title: str
    x_label: str
    y_label: str
    y2_label: str | None = None


def canonical_plot_type(name: str) -> str:
    key = (name or "").strip().lower().replace(" ", "_")
    if not key:
        return DEFAULT_PLOT_TYPE
    if key == "all":
        return key
    resolved = PLOT_TYPE_ALIASES.get(key, key)
    if resolved not in PLOT_TYPE_REGISTRY:
        supported = ", ".join(PLOT_TYPE_REGISTRY)
        raise ValueError(f"不支持的绘图类型: {name}。可选: {supported} / all")
    return resolved


def parse_plot_types(value: str | Iterable[str] | None) -> list[str]:
    if value is None:
        return [DEFAULT_PLOT_TYPE]

    if isinstance(value, str):
        raw = value.strip()
        if not raw:
            return [DEFAULT_PLOT_TYPE]
        tokens = [item.strip() for item in raw.replace("，", ",").replace("；", ",").split(",") if item.strip()]
    else:
        tokens = [str(item).strip() for item in value if str(item).strip()]
        if not tokens:
            return [DEFAULT_PLOT_TYPE]

    if any(token.lower() == "all" for token in tokens):
        return list(PLOT_TYPE_REGISTRY.keys())

    seen: list[str] = []
    for token in tokens:
        resolved = canonical_plot_type(token)
        if resolved not in seen:
            seen.append(resolved)
    return seen or [DEFAULT_PLOT_TYPE]


def plot_type_suffix(plot_type: str) -> str:
    return PLOT_TYPE_REGISTRY[plot_type]["suffix"]


def plot_type_label(plot_type: str) -> str:
    return PLOT_TYPE_REGISTRY[plot_type]["label"]


def available_plot_type_labels() -> list[tuple[str, str]]:
    return [(key, value["label"]) for key, value in PLOT_TYPE_REGISTRY.items()]


def resolve_plot_text(
    plot_type: str,
    *,
    custom_title: str,
    custom_x_label: str,
    custom_y_label: str,
    total_plot_types: int,
    default_title: str,
    default_x_label: str,
    default_y_label: str,
) -> PlotText:
    spec = PLOT_TYPE_REGISTRY[plot_type]
    fallback_title = spec["title"]
    fallback_x = spec["x_label"]
    fallback_y = spec["y_label"]
    fallback_y2 = spec.get("y2_label")

    if total_plot_types > 1:
        return PlotText(title=fallback_title, x_label=fallback_x, y_label=fallback_y, y2_label=fallback_y2)

    title = custom_title if (custom_title or "").strip() and custom_title != default_title else fallback_title
    x_label = custom_x_label if (custom_x_label or "").strip() and custom_x_label != default_x_label else fallback_x
    y_label = custom_y_label if (custom_y_label or "").strip() and custom_y_label != default_y_label else fallback_y
    return PlotText(title=title, x_label=x_label, y_label=y_label, y2_label=fallback_y2)


def compute_active_mass_g(data: pd.DataFrame) -> float | None:
    if "capacity" not in data.columns:
        return None
    if "specific_capacity" not in data.columns:
        return None

    capacity = pd.to_numeric(data["capacity"], errors="coerce")
    specific_capacity = pd.to_numeric(data["specific_capacity"], errors="coerce")
    ratio = capacity / specific_capacity.replace(0, np.nan)
    ratio = ratio.replace([np.inf, -np.inf], np.nan)
    ratio = ratio[(ratio > 0) & ratio.notna()]
    if ratio.empty:
        return None
    return float(ratio.median())


def build_cycle_summary(data: pd.DataFrame) -> pd.DataFrame:
    working = data.copy()
    aggregations: dict[str, tuple[str, str]] = {
        "specific_capacity_max": ("specific_capacity", "max"),
    }
    if "capacity" in working.columns:
        aggregations["capacity_max"] = ("capacity", "max")
    if "current" in working.columns:
        aggregations["current_median"] = ("current", "median")
    if "比能量/mWh/g" in working.columns:
        aggregations["specific_energy_max"] = ("比能量/mWh/g", "max")
    elif "能量/mWh" in working.columns:
        aggregations["energy_max"] = ("能量/mWh", "max")

    grouped = (
        working.groupby(["cycle", "curve_type"], as_index=False)
        .agg(**aggregations)
        .sort_values(["cycle", "curve_type"], kind="mergesort")
        .reset_index(drop=True)
    )

    cycles = sorted(grouped["cycle"].dropna().astype(int).unique().tolist())
    rows: list[dict[str, float | int | None]] = []
    active_mass_g = compute_active_mass_g(working)

    for cycle in cycles:
        subset = grouped[grouped["cycle"] == cycle]
        charge = subset[subset["curve_type"] == "charge"]
        discharge = subset[subset["curve_type"] == "discharge"]

        def _value(frame: pd.DataFrame, column: str) -> float | None:
            if column not in frame.columns or frame.empty:
                return None
            value = frame.iloc[0][column]
            return None if pd.isna(value) else float(value)

        charge_spec = _value(charge, "specific_capacity_max")
        discharge_spec = _value(discharge, "specific_capacity_max")
        charge_cap = _value(charge, "capacity_max")
        discharge_cap = _value(discharge, "capacity_max")
        charge_current = _value(charge, "current_median")
        discharge_current = _value(discharge, "current_median")

        ce = None
        if charge_spec and discharge_spec is not None and charge_spec > 0:
            ce = discharge_spec / charge_spec * 100

        current_density = None
        if active_mass_g and active_mass_g > 0:
            current_source = discharge_current if discharge_current is not None else charge_current
            if current_source is not None:
                current_density = abs(current_source) / active_mass_g

        rows.append(
            {
                "cycle": cycle,
                "charge_specific_capacity": charge_spec,
                "discharge_specific_capacity": discharge_spec,
                "charge_capacity": charge_cap,
                "discharge_capacity": discharge_cap,
                "charge_current": charge_current,
                "discharge_current": discharge_current,
                "coulombic_efficiency": ce,
                "current_density": current_density,
            }
        )

    summary = pd.DataFrame(rows)
    if not summary.empty and "discharge_specific_capacity" in summary.columns:
        first_valid = summary["discharge_specific_capacity"].dropna()
        if not first_valid.empty and first_valid.iloc[0] > 0:
            summary["capacity_retention"] = summary["discharge_specific_capacity"] / first_valid.iloc[0] * 100
    return summary


def _smooth_series(series: pd.Series, window: int = 7) -> pd.Series:
    if len(series) < 5:
        return series
    actual_window = min(window, len(series) if len(series) % 2 == 1 else len(series) - 1)
    actual_window = max(actual_window, 3)
    return series.rolling(actual_window, center=True, min_periods=1).mean()


def build_differential_dataset(data: pd.DataFrame, plot_type: str) -> pd.DataFrame:
    source_column = "dQdV/mAh/V" if plot_type == "dqdv" else "dVdQ/V/mAh"
    x_column = "voltage" if plot_type == "dqdv" else "specific_capacity"
    result_rows: list[pd.DataFrame] = []

    for (cycle, curve_type), segment in data.groupby(["cycle", "curve_type"], sort=True):
        sort_columns = [column for column in ["specific_capacity", "record", "_row_order"] if column in segment.columns]
        ordered = segment.sort_values(sort_columns, kind="mergesort").copy()
        if len(ordered.index) < 5:
            continue

        x_series = pd.to_numeric(ordered[x_column], errors="coerce")
        if source_column in ordered.columns:
            y_series = pd.to_numeric(ordered[source_column], errors="coerce")
        else:
            q = pd.to_numeric(ordered["specific_capacity"], errors="coerce")
            v = pd.to_numeric(ordered["voltage"], errors="coerce")
            if plot_type == "dqdv":
                gradient = np.gradient(q.to_numpy(dtype=float), v.to_numpy(dtype=float))
            else:
                gradient = np.gradient(v.to_numpy(dtype=float), q.to_numpy(dtype=float))
            y_series = pd.Series(gradient, index=ordered.index, dtype=float)

        y_series = _smooth_series(y_series)
        valid = x_series.notna() & y_series.notna() & np.isfinite(x_series) & np.isfinite(y_series)
        filtered = pd.DataFrame(
            {
                "cycle": int(cycle),
                "curve_type": curve_type,
                "x": x_series[valid].astype(float).to_numpy(),
                "y": y_series[valid].astype(float).to_numpy(),
            }
        )
        if not filtered.empty:
            result_rows.append(filtered)

    if not result_rows:
        return pd.DataFrame(columns=["cycle", "curve_type", "x", "y"])

    differential = pd.concat(result_rows, ignore_index=True)
    if not differential.empty:
        y = differential["y"].replace([np.inf, -np.inf], np.nan).dropna()
        if not y.empty:
            threshold = max(abs(y.quantile(0.005)), abs(y.quantile(0.995))) * 1.6
            if threshold > 0:
                differential = differential[differential["y"].abs() <= threshold].copy()
    return differential.reset_index(drop=True)
