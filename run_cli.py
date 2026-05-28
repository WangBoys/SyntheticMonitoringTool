#!/usr/bin/env python3
from __future__ import annotations

import argparse
import signal
import sys
import threading

from config_loader import ConfigError, dry_run_probe_config, load_probe_config
from probe_core import ProbeResult, detect_public_ip_and_isp, execute_probe_job
from runtime_env import setup_runtime_environment


def print_public_network_info() -> None:
    print("检测公网 IP/运营商/地区...")
    try:
        ip, isp, location = detect_public_ip_and_isp()
        print(f"公网 IP: {ip}")
        print(f"运营商: {isp}")
        print(f"国家/省份/城市: {location}")
    except Exception as exc:
        print(f"公网 IP: 检测失败 ({exc})", file=sys.stderr)
        print("运营商: 检测失败（网络受限）", file=sys.stderr)
        print("国家/省份/城市: 检测失败", file=sys.stderr)


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="SyntheticMonitoringTool CLI：通过 JSON 配置文件执行 Excel 网址拨测。",
    )
    parser.add_argument(
        "--config",
        default="config.json",
        help="拨测配置文件路径（JSON，默认: 当前目录下的 config.json）",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="仅校验配置与 Excel 结构，不执行拨测",
    )
    return parser


def main(argv: list[str] | None = None) -> int:
    args = _build_parser().parse_args(argv)
    setup_runtime_environment()
    print_public_network_info()

    try:
        config = load_probe_config(args.config, validate_paths=True)
    except ConfigError as exc:
        print(f"配置错误: {exc}", file=sys.stderr)
        return 1
    except Exception as exc:
        print(f"配置加载失败: {exc}", file=sys.stderr)
        return 1

    if args.dry_run:
        try:
            summary = dry_run_probe_config(config)
        except Exception as exc:
            print(f"结构校验失败: {exc}", file=sys.stderr)
            return 1
        print("配置校验通过。")
        print(f"  输入文件: {summary['file_path']}")
        print(f"  Sheet: {summary['sheet_name']}")
        print(f"  网址列: {summary['url_col_idx']}")
        print(f"  待拨测行数: {summary['row_count']}")
        print(f"  输出文件: {summary['output_path']}")
        return 0

    stop_event = threading.Event()
    interrupted = {"value": False}

    def _handle_stop(signum: int, _frame: object) -> None:
        interrupted["value"] = True
        stop_event.set()

    previous_sigint = signal.signal(signal.SIGINT, _handle_stop)
    previous_sigterm = None
    if hasattr(signal, "SIGTERM"):
        previous_sigterm = signal.signal(signal.SIGTERM, _handle_stop)

    def progress_callback(completed: int, total: int, failed: int, _active_workers: int, result: ProbeResult) -> None:
        status = "OK" if result.accessible else "FAIL"
        code = result.status_code if result.status_code is not None else "-"
        print(f"[{completed}/{total}] 失败:{failed} {status} {code} {result.url}")

    try:
        result = execute_probe_job(
            config,
            progress_callback=progress_callback,
            is_stopped=stop_event.is_set,
        )
    except KeyboardInterrupt:
        return 130
    except Exception as exc:
        print(f"拨测失败: {exc}", file=sys.stderr)
        return 1
    finally:
        signal.signal(signal.SIGINT, previous_sigint)
        if previous_sigterm is not None:
            signal.signal(signal.SIGTERM, previous_sigterm)

    if interrupted["value"] or result.outcome == "aborted":
        print("拨测已中断。", file=sys.stderr)
        return 130

    ratio = (result.success_count / result.total_count) if result.total_count else 0
    print(
        f"拨测完成: 成功 {result.success_count}/{result.total_count} "
        f"({ratio:.2%})，结果已写入 {result.output_path}"
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
