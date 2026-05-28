from __future__ import annotations

import json
import os
import sys
from pathlib import Path
from typing import Any

from probe_core import (
    DEFAULT_CONCURRENCY,
    DEFAULT_RETRY_COUNT,
    DEFAULT_TIMEOUT_SECONDS,
    ProbeConfig,
    SUCCESS_MODE_2XX_3XX,
    SUCCESS_MODE_2XX_3XX_OR_NO_STATUS,
    SUCCESS_MODE_ANY_HTTP,
    SUCCESS_MODE_CUSTOM,
    load_workbook_from_input,
    resolve_url_column_index,
)


class ConfigError(ValueError):
    """Raised when the probe configuration file is invalid."""


SUCCESS_MODE_ALIASES: dict[str, str] = {
    "2xx_3xx": SUCCESS_MODE_2XX_3XX,
    "2xx_3xx_or_no_status": SUCCESS_MODE_2XX_3XX_OR_NO_STATUS,
    "any_http": SUCCESS_MODE_ANY_HTTP,
    "custom": SUCCESS_MODE_CUSTOM,
}


def _require_mapping(value: object, field_name: str) -> dict[str, Any]:
    if not isinstance(value, dict):
        raise ConfigError(f"{field_name} 必须是对象")
    return value


def _resolve_password(input_section: dict[str, Any]) -> str:
    password_env = input_section.get("password_env")
    if password_env is not None:
        env_name = str(password_env).strip()
        if env_name:
            value = os.environ.get(env_name, "")
            if not value:
                raise ConfigError(f"环境变量 {env_name!r} 未设置或为空")
            return value
    password = input_section.get("password", "")
    return str(password) if password is not None else ""


def map_success_mode(raw_mode: object) -> str:
    if raw_mode is None:
        return SUCCESS_MODE_ANY_HTTP
    mode = str(raw_mode).strip()
    if not mode:
        raise ConfigError("probe.success_mode 不能为空")
    if mode in SUCCESS_MODE_ALIASES:
        return SUCCESS_MODE_ALIASES[mode]
    allowed = ", ".join(sorted(SUCCESS_MODE_ALIASES))
    raise ConfigError(f"不支持的 success_mode: {mode!r}，可选: {allowed}")


def resolve_output_path(file_path: str, output_section: dict[str, Any]) -> str:
    raw_path = output_section.get("path")
    if raw_path is None or not str(raw_path).strip():
        return os.path.splitext(file_path)[0] + "_monitoring_result.xlsx"
    return str(raw_path).strip()


def _validate_numeric_range(name: str, value: object, minimum: int, maximum: int) -> int:
    if isinstance(value, bool) or not isinstance(value, int):
        raise ConfigError(f"{name} 必须是整数")
    if value < minimum or value > maximum:
        raise ConfigError(f"{name} 必须在 {minimum}~{maximum} 之间")
    return value


def load_probe_config_from_dict(data: dict[str, Any], *, validate_paths: bool = True) -> ProbeConfig:
    input_section = _require_mapping(data.get("input"), "input")
    probe_section = _require_mapping(data.get("probe"), "probe")
    output_section = _require_mapping(data.get("output", {}), "output")

    file_path = str(input_section.get("file_path", "")).strip()
    if not file_path:
        raise ConfigError("input.file_path 不能为空")

    use_office_fallback = bool(input_section.get("use_office_fallback", False))
    if use_office_fallback and sys.platform != "win32":
        raise ConfigError("input.use_office_fallback 仅在 Windows 上可用")

    sheet_name = str(probe_section.get("sheet_name", "")).strip()
    if not sheet_name:
        raise ConfigError("probe.sheet_name 不能为空")

    if "has_header" not in probe_section:
        raise ConfigError("probe.has_header 必须指定")
    has_header = bool(probe_section["has_header"])

    if "url_column" not in probe_section:
        raise ConfigError("probe.url_column 必须指定")
    url_column = probe_section["url_column"]

    concurrency = _validate_numeric_range(
        "probe.concurrency",
        probe_section.get("concurrency", DEFAULT_CONCURRENCY),
        1,
        100,
    )
    timeout_seconds = _validate_numeric_range(
        "probe.timeout_seconds",
        probe_section.get("timeout_seconds", DEFAULT_TIMEOUT_SECONDS),
        1,
        300,
    )
    retry_count = _validate_numeric_range(
        "probe.retry_count",
        probe_section.get("retry_count", DEFAULT_RETRY_COUNT),
        0,
        5,
    )

    success_mode = map_success_mode(probe_section.get("success_mode", "any_http"))
    custom_codes_input = str(probe_section.get("custom_status_codes", "")).strip()
    if success_mode == SUCCESS_MODE_CUSTOM and not custom_codes_input:
        raise ConfigError("success_mode 为 custom 时必须设置 probe.custom_status_codes")

    probe_mode = str(probe_section.get("probe_mode", "lightweight")).strip()
    if probe_mode not in ("lightweight", "browser"):
        raise ConfigError(f"probe.probe_mode 必须是 lightweight 或 browser，当前: {probe_mode!r}")

    extract_title = bool(probe_section.get("extract_title", True))
    track_redirects = bool(probe_section.get("track_redirects", True))

    password = _resolve_password(input_section)
    output_path = resolve_output_path(file_path, output_section)

    if validate_paths and not os.path.isfile(file_path):
        raise ConfigError(f"输入文件不存在: {file_path}")

    wb, _ = load_workbook_from_input(file_path, password, use_office_fallback)
    if sheet_name not in wb.sheetnames:
        raise ConfigError(f"Sheet 不存在: {sheet_name!r}，可用: {', '.join(wb.sheetnames)}")

    sheet = wb[sheet_name]
    try:
        url_col_idx = resolve_url_column_index(sheet, url_column, has_header)
    except ValueError as exc:
        raise ConfigError(str(exc)) from exc

    return ProbeConfig(
        file_path=file_path,
        use_office_fallback=use_office_fallback,
        sheet_name=sheet_name,
        has_header=has_header,
        url_col_idx=url_col_idx,
        concurrency=concurrency,
        timeout_seconds=timeout_seconds,
        retry_count=retry_count,
        success_mode=success_mode,
        custom_codes_input=custom_codes_input,
        output_path=output_path,
        password=password,
        probe_mode=probe_mode,
        extract_title=extract_title,
        track_redirects=track_redirects,
    )


def load_probe_config(path: str | Path, *, validate_paths: bool = True) -> ProbeConfig:
    config_path = Path(path)
    if not config_path.is_file():
        raise ConfigError(f"配置文件不存在: {config_path}")

    try:
        data = json.loads(config_path.read_text(encoding="utf-8"))
    except json.JSONDecodeError as exc:
        raise ConfigError(f"配置文件 JSON 格式错误: {exc}") from exc

    if not isinstance(data, dict):
        raise ConfigError("配置文件根节点必须是对象")

    return load_probe_config_from_dict(data, validate_paths=validate_paths)


def dry_run_probe_config(config: ProbeConfig) -> dict[str, int | str]:
    wb, _ = load_workbook_from_input(config.file_path, config.password, config.use_office_fallback)
    sheet = wb[config.sheet_name]
    start_row = 2 if config.has_header else 1
    row_count = max(0, sheet.max_row - start_row + 1)
    return {
        "file_path": config.file_path,
        "sheet_name": config.sheet_name,
        "url_col_idx": config.url_col_idx,
        "row_count": row_count,
        "output_path": config.output_path,
    }
