from __future__ import annotations

import re
from pathlib import Path
from typing import Callable, Iterable

from config import DEFAULT_FONT_CANDIDATES, DEFAULT_OUTPUT_FORMAT, SUPPORTED_EXCEL_SUFFIXES


class UserInputError(ValueError):
    """Raised when user-provided text cannot be parsed into a valid option."""


def strip_wrapping_quotes(text: str) -> str:
    value = text.strip()
    if len(value) >= 2 and value[0] == value[-1] and value[0] in {'"', "'"}:
        return value[1:-1].strip()
    return value


def normalize_path_input(text: str) -> Path:
    value = strip_wrapping_quotes(text)
    if not value:
        raise UserInputError("路径不能为空。")
    return Path(value).expanduser()


def validate_input_file(path_text: str) -> Path:
    path = normalize_path_input(path_text)
    if not path.exists() or not path.is_file():
        raise UserInputError(f"文件不存在: {path}")
    if path.suffix.lower() == ".ccs":
        return path
    if path.suffix.lower() not in SUPPORTED_EXCEL_SUFFIXES:
        raise UserInputError("当前仅支持 .xlsx / .xlsm 文件，或未直接支持的 .ccs 原始文件。")
    return path


def build_output_path(path_text: str, output_format: str = DEFAULT_OUTPUT_FORMAT) -> Path:
    path = normalize_path_input(path_text)
    try:
        suffix = path.suffix.lower()
        if not suffix:
            path = path.with_suffix(f".{output_format.lstrip('.').lower()}")
    except ValueError as exc:
        raise UserInputError("输出路径必须包含有效的文件名。") from exc
    path.parent.mkdir(parents=True, exist_ok=True)
    return path


def parse_cycle_expression(value: str | Iterable[str] | None) -> list[int] | None:
    if value is None:
        return None

    if isinstance(value, str):
        raw = value.strip()
    else:
        tokens = [str(item).strip() for item in value if str(item).strip()]
        raw = ",".join(tokens)

    if not raw:
        return None

    raw = raw.replace("~", "-").replace("，", ",").replace("；", ",")
    selected: set[int] = set()

    for token in raw.split(","):
        part = token.strip()
        if not part:
            continue
        if "-" in part:
            start_text, end_text = [item.strip() for item in part.split("-", 1)]
            if not start_text or not end_text:
                raise UserInputError(f"无法解析循环范围: {part}")
            start = int(float(start_text))
            end = int(float(end_text))
            lower, upper = sorted((start, end))
            selected.update(range(lower, upper + 1))
        else:
            selected.add(int(float(part)))

    if not selected:
        return None
    return sorted(selected)


def parse_mode_overrides(text: str | Iterable[str] | None) -> dict[str, str]:
    if text is None:
        return {}

    if isinstance(text, str):
        raw = text.strip()
        if not raw:
            return {}
        items = [item.strip() for item in raw.replace("；", ",").split(",") if item.strip()]
    else:
        items = [str(item).strip() for item in text if str(item).strip()]

    overrides: dict[str, str] = {}
    for item in items:
        if "=" not in item:
            raise UserInputError(f"工作模式映射格式错误: {item}，应为 别名=charge")
        alias, target = [part.strip() for part in item.split("=", 1)]
        if not alias or not target:
            raise UserInputError(f"工作模式映射格式错误: {item}")
        canonical = target.lower()
        if canonical not in {"charge", "discharge", "rest"}:
            raise UserInputError(f"工作模式映射目标必须是 charge/discharge/rest: {item}")
        overrides[alias] = canonical
    return overrides


def parse_bool_text(value: str, default: bool) -> bool:
    text = value.strip().lower()
    if not text:
        return default
    if text in {"y", "yes", "true", "1", "是"}:
        return True
    if text in {"n", "no", "false", "0", "否"}:
        return False
    raise UserInputError("请输入 y/n。")


def parse_int_text(value: str, default: int, minimum: int | None = None) -> int:
    text = value.strip()
    if not text:
        return default
    result = int(float(text))
    if minimum is not None and result < minimum:
        raise UserInputError(f"数值不能小于 {minimum}。")
    return result


def parse_float_text(value: str, default: float, minimum: float | None = None) -> float:
    text = value.strip()
    if not text:
        return default
    result = float(text)
    if minimum is not None and result < minimum:
        raise UserInputError(f"数值不能小于 {minimum}。")
    return result


def parse_axis_limits(value: str | None) -> tuple[float, float] | None:
    if value is None:
        return None
    text = value.strip()
    if not text:
        return None
    parts = [item for item in re.split(r"[,\s]+", text.replace("，", ",")) if item]
    if len(parts) != 2:
        raise UserInputError("坐标轴范围请输入两个数字，例如 0,4.5")
    lower = float(parts[0])
    upper = float(parts[1])
    if lower >= upper:
        raise UserInputError("坐标轴范围必须满足下限 < 上限。")
    return (lower, upper)


def parse_figure_size(value: str | None, default: tuple[float, float]) -> tuple[float, float]:
    if value is None:
        return default
    text = value.strip()
    if not text:
        return default
    parts = [item for item in re.split(r"[,xX\s]+", text.replace("，", ",")) if item]
    if len(parts) != 2:
        raise UserInputError("图尺寸请输入两个数字，例如 8,6")
    width = float(parts[0])
    height = float(parts[1])
    if width <= 0 or height <= 0:
        raise UserInputError("图尺寸必须为正数。")
    return (width, height)


def parse_font_family(value: str | None) -> list[str]:
    if value is None:
        return list(DEFAULT_FONT_CANDIDATES)
    text = value.strip()
    if not text:
        return list(DEFAULT_FONT_CANDIDATES)
    fonts = [item.strip() for item in text.replace("；", ",").split(",") if item.strip()]
    return fonts or list(DEFAULT_FONT_CANDIDATES)


def parse_output_format(value: str | None, default: str = DEFAULT_OUTPUT_FORMAT) -> str:
    text = (value or "").strip().lower().lstrip(".")
    if not text:
        return default
    if text not in {"png", "svg", "pdf"}:
        raise UserInputError("输出格式仅支持 png / svg / pdf。")
    return text


def parse_optional_color(value: str | None, default: str) -> str:
    text = (value or "").strip()
    return text or default


def prompt_until_valid(
    prompt: str,
    parser: Callable[[str], object],
    error_prefix: str = "输入无效",
) -> object:
    while True:
        user_text = input(prompt)
        try:
            return parser(user_text)
        except UserInputError as exc:
            print(f"{error_prefix}: {exc}")
        except ValueError as exc:
            print(f"{error_prefix}: {exc}")


def prompt_text(prompt: str, default: str | None = None) -> str:
    text = input(prompt)
    if not text.strip() and default is not None:
        return default
    return text.strip()
