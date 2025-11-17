#!/usr/bin/env python3
"""Orchestrator for the find_my_boss workflow (step B)."""

from __future__ import annotations

import argparse
import os
import json
import logging
import subprocess
import sys
import threading
from pathlib import Path
from typing import Any, Dict, Optional

try:  # pragma: no cover - imported only when type checking
    from main import ChoujiRobo
except Exception:  # pragma: no cover - during CLI execution main.py is not imported
    ChoujiRobo = Any  # type: ignore[assignment]

SCOPES = ["User.Read.All", "Directory.Read.All"]
REQUEST_TIMEOUT_SECONDS = 3
MANAGER_MAX_DEPTH = 15

HERE = Path(__file__).resolve().parent
ROBO_SCRIPTS_ROOT = HERE.parent
if str(ROBO_SCRIPTS_ROOT) not in sys.path:
    sys.path.append(str(ROBO_SCRIPTS_ROOT))
POWER_SHELL = "pwsh"
LOGGER = logging.getLogger("chouji_robo.find_my_boss")

from module_loader import load_helper


def _configure_cli_logging() -> None:
    """Ensure logging is configured when the module is executed standalone."""

    root = logging.getLogger()
    if root.handlers:
        return

    handler = logging.StreamHandler(sys.stdout)
    formatter = logging.Formatter(
        fmt="%(asctime)s [%(levelname)s] %(name)s :: %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )
    handler.setFormatter(formatter)
    root.addHandler(handler)
    root.setLevel(logging.DEBUG)


def _run_powershell(script: Path, *extra_args: str) -> Dict[str, Any]:
    """Execute a PowerShell helper and return its JSON payload."""

    cmd = [
        POWER_SHELL,
        "-NoLogo",
        "-NoProfile",
        "-ExecutionPolicy",
        "Bypass",
        "-File",
        str(script),
    ]
    cmd.extend(extra_args)

    LOGGER.info("[STEP] PowerShell 実行開始: %s %s", script.name, " ".join(extra_args))

    process = subprocess.Popen(
        cmd,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
        bufsize=1,
    )

    stdout_lines: list[str] = []
    stderr_lines: list[str] = []

    def _consume(stream, collector, prefix: str) -> None:
        if stream is None:
            return
        json_prefix_chars = {"{", "[", "}"}
        passthrough_prefixes = ("[STEP", "[INFO", "[DEBUG", "[WARNING", "[ERROR", "[VERBOSE")
        for raw_line in stream:
            line = raw_line.rstrip("\n")
            collector.append(line)
            stripped = line.strip()
            if not stripped:
                continue
            first = stripped[:1]
            if stripped.startswith(passthrough_prefixes):
                if prefix == "STDOUT":
                    LOGGER.info(stripped)
                else:
                    LOGGER.error(stripped)
                continue
            if first in json_prefix_chars or stripped.startswith('"') or stripped.endswith(':') or stripped.endswith(','):
                continue
            if prefix == "STDOUT":
                LOGGER.info(stripped)
            else:
                LOGGER.error(stripped)

    stdout_thread = threading.Thread(target=_consume, args=(process.stdout, stdout_lines, "STDOUT"), daemon=True)
    stderr_thread = threading.Thread(target=_consume, args=(process.stderr, stderr_lines, "STDERR"), daemon=True)
    stdout_thread.start()
    stderr_thread.start()

    return_code = process.wait()
    stdout_thread.join()
    stderr_thread.join()

    json_payload = "{}"
    if stdout_lines:
        # Search from bottom for first line starting JSON payload
        start_idx = None
        for idx in range(len(stdout_lines) - 1, -1, -1):
            stripped = stdout_lines[idx].strip()
            if stripped.startswith("{") or stripped.startswith("["):
                start_idx = idx
                break
        if start_idx is not None:
            payload_lines = stdout_lines[start_idx:]
            json_payload = []
            brace_stack = 0
            for line in payload_lines:
                json_payload.append(line)
                stripped = line.strip()
                brace_stack += sum(1 for ch in stripped if ch in "{[")
                brace_stack -= sum(1 for ch in stripped if ch in "}]")
                if brace_stack <= 0 and stripped.endswith(('}', ']')):
                    break
            json_payload = "\n".join(json_payload).strip()
        else:
            json_payload = stdout_lines[-1].strip() if stdout_lines else "{}"

    if return_code != 0:
        stdout_joined = "\n".join(stdout_lines)
        stderr_joined = "\n".join(stderr_lines)
        raise RuntimeError(
            f"Script {script.name} failed with exit code {return_code}.\n"
            f"STDOUT:\n{stdout_joined}\nSTDERR:\n{stderr_joined}"
        )

    try:
        data = json.loads(json_payload)
    except json.JSONDecodeError as exc:
        raise RuntimeError(
            f"Failed to parse JSON output from {script.name}: {json_payload}"
        ) from exc

    LOGGER.info("[STEP] PowerShell 実行完了: %s", script.name)
    return data


def _execute_workflow(
    scopes: Optional[list[str]] = None,
    timeout_seconds: Optional[int] = None,
    max_depth: Optional[int] = None,
    prefer_device_auth: Optional[bool] = None,
    skip_module_install: Optional[bool] = None,
    include_user_extended: Optional[bool] = None,
    include_manager_extended: Optional[bool] = None,
) -> Dict[str, Any]:
    """Run the full find_my_boss workflow and return collected data."""

    active_scopes = scopes or SCOPES
    active_timeout = timeout_seconds or REQUEST_TIMEOUT_SECONDS
    active_depth = max_depth or MANAGER_MAX_DEPTH

    scope_arg = ",".join(active_scopes)
    timeout_arg = str(active_timeout)
    depth_arg = str(active_depth)

    login_script = HERE / "Ba.login_msGraph.ps1"
    user_script = HERE / "Bb.get_user_data.ps1"
    boss_script = HERE / "Bc.get_boss_data.ps1"

    LOGGER.info("[STEP] B.find_my_boss workflowを開始します。")
    login_args = [
        "-Scopes",
        scope_arg,
        "-RequestTimeoutSeconds",
        timeout_arg,
    ]

    if skip_module_install is None:
        skip_module_install = True
        skip_install_env = os.getenv("CHOUJI_SKIP_GRAPH_INSTALL", "")
        if skip_install_env:
            skip_module_install = skip_install_env.lower() in {"1", "true", "yes", "on"}
    if skip_module_install:
        LOGGER.info("[INFO] Graph module installation check will be skipped (SkipModuleInstall).")
        login_args.append("-SkipModuleInstall")


    if include_user_extended is None:
        include_user_extended = os.getenv("CHOUJI_INCLUDE_USER_EXTENDED", "").lower() in {"1", "true", "yes", "on"}
    if include_manager_extended is None:
        include_manager_extended = os.getenv("CHOUJI_INCLUDE_MANAGER_EXTENDED", "").lower() in {"1", "true", "yes", "on"}

    device_auth = True
    if prefer_device_auth is not None:
        device_auth = prefer_device_auth
    else:
        device_env = os.getenv("CHOUJI_USE_DEVICE_AUTH", "")
        if device_env:
            device_auth = device_env.lower() in {"1", "true", "yes", "on"}
        browser_env = os.getenv("CHOUJI_USE_BROWSER_AUTH", "")
        if browser_env.lower() in {"1", "true", "yes", "on"}:
            device_auth = False

    if device_auth:
        LOGGER.info("[INFO] Device-code authentication will be used for Microsoft Graph.")
        login_args.append("-UseDeviceAuth")
    else:
        LOGGER.info("[INFO] Browser-based authentication will be used for Microsoft Graph.")

    login_data = _run_powershell(
        login_script,
        *login_args,
    )

    mail_honnin = (login_data.get("mail_honnin") or "").strip()
    if not mail_honnin:
        raise RuntimeError("RPAシートのJ5からメールアドレスを取得できませんでした。")

    LOGGER.info("[STEP] 対象ユーザー: %s", mail_honnin)

    user_args = [
        "-UserEmail",
        mail_honnin,
        "-Scopes",
        scope_arg,
        "-RequestTimeoutSeconds",
        timeout_arg,
    ]
    if skip_module_install:
        user_args.append("-SkipModuleInstall")
    if include_user_extended:
        user_args.append("-IncludeExtendedData")

    user_data = _run_powershell(
        user_script,
        *user_args,
    )

    boss_args = [
        "-UserEmail",
        mail_honnin,
        "-Scopes",
        scope_arg,
        "-RequestTimeoutSeconds",
        timeout_arg,
        "-MaxDepth",
        depth_arg,
    ]
    if skip_module_install:
        boss_args.append("-SkipModuleInstall")
    if include_manager_extended:
        boss_args.append("-IncludeExtendedData")

    bosses_data = _run_powershell(
        boss_script,
        *boss_args,
    )

    results: Dict[str, Any] = {
        "mail_honnin": mail_honnin,
        "login": login_data,
        "user": user_data,
        "managers": bosses_data,
    }

    _emit_summary(results)
    LOGGER.info("[STEP] B.find_my_boss ワークフローを終了します。")
    return results


def _emit_summary(results: Dict[str, Any]) -> None:
    """Log a short summary to match the test script output style."""

    mail_honnin: str = results.get("mail_honnin", "")
    user_data: Dict[str, Any] = results.get("user") or {}
    user_detail: Dict[str, Any] = user_data.get("userDetail") or {}
    user_extended: Dict[str, Any] = user_data.get("extended") or {}

    LOGGER.info("User: %s", mail_honnin)
    if user_detail:
        LOGGER.info(
            "  DisplayName: %s",
            user_data.get("nameFullWidth") or user_detail.get("displayName") or "",
        )
        LOGGER.info("  CompanyName: %s", user_detail.get("companyName") or "")
        LOGGER.info("  Department: %s", user_detail.get("department") or "")
        LOGGER.info("  JobTitle: %s", user_detail.get("jobTitle") or "")

    managers_wrapper: Dict[str, Any] = results.get("managers") or {}
    if isinstance(managers_wrapper, dict):
        managers = managers_wrapper.get("managers")
    else:
        managers = managers_wrapper
    if managers is None:
        managers = []

    LOGGER.info("Manager chain (count=%s):", len(managers))
    for entry in managers:
        LOGGER.info(
            "  Level %s -> %s / %s / %s / %s",
            entry.get("Index"),
            entry.get("DisplayName") or entry.get("Mail") or "(no name)",
            entry.get("Mail") or "",
            entry.get("Department") or "",
            entry.get("JobTitle") or "",
        )

    if user_extended:
        LOGGER.info(
            "Extended info: licenses=%s, groups=%s, appRoles=%s",
            len(user_extended.get("LicenseDetails") or []),
            len(user_extended.get("MemberOf") or []),
            len(user_extended.get("AppRoleAssignments") or []),
        )





def _run_python_helper(module_name: str, phase_name: str, robot: Optional["ChoujiRobo"]) -> Optional[Dict[str, Any]]:
    """Load and execute a sibling Bd/Be helper while updating the UI phase."""

    try:
        helper = load_helper(module_name)
    except Exception as exc:
        LOGGER.exception("Failed to load helper %s: %s", module_name, exc)
        return None

    run_func = getattr(helper, "run", None)
    if not callable(run_func):
        LOGGER.error("Helper %s does not define run().", module_name)
        return None

    previous_phase = getattr(robot, "current_phase", None)
    if robot is not None:
        robot.current_phase = phase_name

    LOGGER.info("%s を開始します。", phase_name)
    try:
        return run_func()
    except Exception as exc:
        LOGGER.exception("%s 実行中にエラーが発生しました: %s", phase_name, exc)
        return None
    finally:
        if robot is not None:
            robot.current_phase = previous_phase or "B.find_my_boss"


def run(robot: Optional["ChoujiRobo"] = None) -> Dict[str, Any]:
    """Entry point invoked from the main robot workflow."""

    if robot is not None:
        robot.current_phase = "B.find_my_boss"
        LOGGER.info("B.find_my_boss を開始します。")
    results = _execute_workflow(skip_module_install=True, include_user_extended=True, include_manager_extended=True)

    job_lookup = _run_python_helper("Bd.find_job_title", "Bd.find_job_title", robot)
    if job_lookup is not None:
        results["job_title_lookup"] = job_lookup

    kachou_logic = _run_python_helper("Be.Kachou_hantei", "Be.Kachou_hantei", robot)
    if kachou_logic is not None:
        results["kachou_logic"] = kachou_logic
    if robot is not None:
        manager_payload = results.get("managers") or {}
        setattr(robot.state, "manager_chain", manager_payload.get("managers"))
        setattr(robot.state, "manager_user_profile", results.get("user"))
        LOGGER.info("B.find_my_boss を正常終了しました。")
    return results


def main(argv: Optional[list[str]] = None) -> int:  # pragma: no cover - CLI support
    """CLI wrapper to execute the workflow outside of the robot."""

    _configure_cli_logging()

    parser = argparse.ArgumentParser(description="Run the find_my_boss workflow.")
    parser.add_argument(
        "--scopes",
        help="Comma separated Microsoft Graph scopes.",
        default=",".join(SCOPES),
    )
    parser.add_argument(
        "--timeout",
        type=int,
        default=REQUEST_TIMEOUT_SECONDS,
        help="Graph request timeout in seconds.",
    )
    parser.add_argument(
        "--max-depth",
        type=int,
        default=MANAGER_MAX_DEPTH,
        help="Maximum number of manager levels to fetch.",
    )
    parser.add_argument(
        "--force-module-install",
        action="store_true",
        help="Force Microsoft Graph module installation check before each run.",
    )
    parser.add_argument(
        "--use-device-auth",
        action="store_true",
        help="Force device code flow for Microsoft Graph authentication.",
    )
    parser.add_argument(
        "--use-browser-auth",
        action="store_true",
        help="Force browser-based authentication for Microsoft Graph.",
    )
    parser.add_argument(
        "--include-user-extended",
        action="store_true",
        help="Include extended user data (slower).",
    )
    parser.add_argument(
        "--include-manager-extended",
        action="store_true",
        help="Include extended manager data (significantly slower).",
    )
    args = parser.parse_args(argv)

    scopes = [scope.strip() for scope in args.scopes.split(",") if scope.strip()]
    timeout = args.timeout
    depth = args.max_depth
    skip_module_install = not args.force_module_install

    prefer_device = None
    if args.use_device_auth:
        prefer_device = True
    if args.use_browser_auth:
        prefer_device = False

    include_user_ext = args.include_user_extended
    include_manager_ext = args.include_manager_extended

    _execute_workflow(
        scopes=scopes,
        timeout_seconds=timeout,
        max_depth=depth,
        prefer_device_auth=prefer_device,
        skip_module_install=skip_module_install,
        include_user_extended=include_user_ext,
        include_manager_extended=include_manager_ext,
    )
    return 0


if __name__ == "__main__":  # pragma: no cover - manual execution only
    sys.exit(main())

