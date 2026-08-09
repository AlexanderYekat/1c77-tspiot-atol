#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Сборка kktserverindyProject без ручного запуска IDE.

Использование:
  python build_delphi.py
  python build_delphi.py --config Release
  python build_delphi.py --clean
  python build_delphi.py --open-ide

Требуется установленный Embarcadero RAD Studio / Delphi.
На части редакций (в т.ч. без CLI-лицензии) компилятор отвечает:
  "This version of the product does not support command line compiling."
Тогда сборка возможна только из IDE (Project → Build / Shift+F9).
"""

from __future__ import annotations

import argparse
import os
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
PROJECT = ROOT / "kktserverindyProject.dproj"
DPR = ROOT / "kktserverindyProject.dpr"

DEFAULT_BDS = Path(r"c:\program files (x86)\embarcadero\studio\23.0")
CLI_BLOCKED_MARKERS = (
    "does not support command line compiling",
    "не поддерживает компиляцию из командной строки",
)


def find_bds_root() -> Path:
    env = os.environ.get("BDS", "").strip()
    if env:
        p = Path(env)
        if (p / "bin" / "rsvars.bat").is_file():
            return p
    if (DEFAULT_BDS / "bin" / "rsvars.bat").is_file():
        return DEFAULT_BDS
    # Перебор типичных версий Studio
    base = Path(r"c:\program files (x86)\embarcadero\studio")
    if base.is_dir():
        for child in sorted(base.iterdir(), reverse=True):
            if (child / "bin" / "rsvars.bat").is_file():
                return child
    raise FileNotFoundError(
        "Не найден Embarcadero Studio (rsvars.bat). "
        "Укажите путь через переменную окружения BDS."
    )


def exe_out(config: str, platform: str) -> Path:
    return ROOT / platform / config / "kktserverindyProject.exe"


def run_msbuild(bds: Path, config: str, platform: str, clean: bool) -> tuple[int, str]:
    rsvars = bds / "bin" / "rsvars.bat"
    target = "Rebuild" if clean else "Build"
    bat = ROOT / "_build_delphi_run.bat"
    bat_text = (
        "@echo off\r\n"
        "setlocal\r\n"
        f'call "{rsvars}"\r\n'
        "if errorlevel 1 exit /b 1\r\n"
        # Относительный путь: в каталоге с кириллицей abs-path в .bat ломает кодировку cmd
        f"msbuild kktserverindyProject.dproj /t:{target} "
        f"/p:Config={config} /p:Platform={platform} /v:minimal\r\n"
        "exit /b %ERRORLEVEL%\r\n"
    )
    bat.write_text(bat_text, encoding="ascii")
    print(f"BDS: {bds}")
    print(f"Project: {PROJECT}")
    print(f"Target: {target}  Config={config}  Platform={platform}")
    print("---")
    try:
        proc = subprocess.run(
            ["cmd", "/c", str(bat)],
            cwd=str(ROOT),
            capture_output=True,
            text=True,
            encoding="cp866",
            errors="replace",
        )
        out = (proc.stdout or "") + (proc.stderr or "")
        if out.strip():
            try:
                print(out.rstrip())
            except UnicodeEncodeError:
                sys.stdout.buffer.write(
                    (out.rstrip() + "\n").encode(sys.stdout.encoding or "utf-8", errors="replace")
                )
        return proc.returncode, out
    finally:
        try:
            bat.unlink(missing_ok=True)
        except OSError:
            pass


def looks_cli_blocked(output: str) -> bool:
    low = output.lower()
    return any(m.lower() in low for m in CLI_BLOCKED_MARKERS)


def open_ide(bds: Path) -> int:
    bds_exe = bds / "bin" / "bds.exe"
    if not bds_exe.is_file():
        print(f"ERROR: не найден {bds_exe}", file=sys.stderr)
        return 1
    print(f"Открываю IDE: {bds_exe}")
    print(f"Проект: {PROJECT}")
    subprocess.Popen([str(bds_exe), str(PROJECT)], cwd=str(ROOT))
    return 0


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Сборка Delphi KKT-сервера (MSBuild)")
    parser.add_argument("--config", default="Debug", choices=["Debug", "Release"])
    parser.add_argument("--platform", default="Win32", choices=["Win32", "Win64"])
    parser.add_argument("--clean", action="store_true", help="Rebuild вместо Build")
    parser.add_argument(
        "--open-ide",
        action="store_true",
        help="Только открыть проект в Delphi IDE",
    )
    parser.add_argument(
        "--bds",
        default="",
        help="Корень Studio (по умолчанию авто / %%BDS%%)",
    )
    args = parser.parse_args(argv)

    if not PROJECT.is_file():
        print(f"ERROR: нет файла проекта {PROJECT}", file=sys.stderr)
        return 1

    try:
        bds = Path(args.bds) if args.bds else find_bds_root()
    except FileNotFoundError as e:
        print(f"ERROR: {e}", file=sys.stderr)
        return 1

    if args.open_ide:
        return open_ide(bds)

    code, output = run_msbuild(bds, args.config, args.platform, args.clean)
    out_exe = exe_out(args.config, args.platform)

    if looks_cli_blocked(output):
        print()
        print("=== CLI-сборка недоступна для этой редакции Delphi ===")
        print('Компилятор: "This version of the product does not support command line compiling."')
        print()
        print("Что делать:")
        print("  1) Собрать в IDE: открыть kktserverindyProject.dproj -> Project -> Build")
        print("     или:  python build_delphi.py --open-ide")
        print("  2) Либо использовать редакцию/лицензию с поддержкой командной строки")
        print("     (Professional/Enterprise с CLI, Build Server и т.п.)")
        return 2

    if code != 0:
        print(f"\nERROR: msbuild exit code {code}", file=sys.stderr)
        return code

    if out_exe.is_file():
        print(f"\nOK: {out_exe} ({out_exe.stat().st_size} bytes)")
        return 0

    print(f"\nWARN: msbuild OK, но exe не найден: {out_exe}", file=sys.stderr)
    return 1


if __name__ == "__main__":
    sys.exit(main())
