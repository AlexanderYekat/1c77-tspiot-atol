#!/usr/bin/env python3
"""
Перенос кода модулей: 1cv77/*.txt  →  1cv77-extforms/*.ert

Сопоставление по имени файла без расширения
(например razresh_regim_check.txt ↔ razresh_regim_check.ert).

Примеры:
    python push_txt_to_ert.py              # показать пары и спросить подтверждение
    python push_txt_to_ert.py --dry-run    # только показать, что будет сделано
    python push_txt_to_ert.py --yes        # перенести без вопроса
    python push_txt_to_ert.py razresh_regim_check  # только указанные имена
"""

from __future__ import annotations

import argparse
import sys
import tempfile
from pathlib import Path

from ert_module import ErtError, replace_text

ROOT = Path(__file__).resolve().parent
TXT_DIR = ROOT / "1cv77"
ERT_DIR = ROOT / "1cv77-extforms"


def collect_pairs(names: list[str] | None = None) -> tuple[list[tuple[str, Path, Path]], list[str], list[str]]:
    """Вернуть (пары, только_txt, только_ert)."""
    txt_by_stem = {p.stem: p for p in TXT_DIR.glob("*.txt")}
    ert_by_stem = {p.stem: p for p in ERT_DIR.glob("*.ert")}

    if names:
        wanted = set(names)
        unknown = wanted - set(txt_by_stem) - set(ert_by_stem)
        if unknown:
            raise ErtError(
                "Не найдены имена среди txt/ert: " + ", ".join(sorted(unknown))
            )
        stems = sorted(wanted & set(txt_by_stem) & set(ert_by_stem))
        only_txt = sorted(wanted & set(txt_by_stem) - set(ert_by_stem))
        only_ert = sorted(wanted & set(ert_by_stem) - set(txt_by_stem))
    else:
        stems = sorted(set(txt_by_stem) & set(ert_by_stem))
        only_txt = sorted(set(txt_by_stem) - set(ert_by_stem))
        only_ert = sorted(set(ert_by_stem) - set(txt_by_stem))

    pairs = [(stem, txt_by_stem[stem], ert_by_stem[stem]) for stem in stems]
    return pairs, only_txt, only_ert


def print_plan(pairs: list[tuple[str, Path, Path]], only_txt: list[str], only_ert: list[str]) -> None:
    print(f"Направление: {TXT_DIR.name}/*.txt  ->  {ERT_DIR.name}/*.ert")
    print()
    if pairs:
        print(f"Сопоставлено ({len(pairs)}):")
        for stem, txt_path, ert_path in pairs:
            print(f"  {txt_path.name}  ->  {ert_path.name}")
    else:
        print("Сопоставленных пар нет.")

    if only_txt:
        print()
        print(f"Только в {TXT_DIR.name} (пропуск):")
        for stem in only_txt:
            print(f"  {stem}.txt")

    if only_ert:
        print()
        print(f"Только в {ERT_DIR.name} (пропуск):")
        for stem in only_ert:
            print(f"  {stem}.ert")


def push_one(txt_path: Path, ert_path: Path) -> None:
    with tempfile.NamedTemporaryFile(
        suffix=".ert", delete=False, dir=ert_path.parent
    ) as tmp:
        tmp_path = Path(tmp.name)
    try:
        replace_text(ert_path, txt_path, tmp_path)
        tmp_path.replace(ert_path)
    finally:
        if tmp_path.exists():
            tmp_path.unlink(missing_ok=True)


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Записать текст из 1cv77/*.txt в модули 1cv77-extforms/*.ert"
    )
    parser.add_argument(
        "names",
        nargs="*",
        help="имена модулей без расширения; если не указаны — все сопоставленные",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="только показать сопоставление, ничего не менять",
    )
    parser.add_argument(
        "--yes",
        "-y",
        action="store_true",
        help="не спрашивать подтверждение",
    )
    args = parser.parse_args(argv)

    try:
        if not TXT_DIR.is_dir():
            raise ErtError(f"Нет каталога {TXT_DIR}")
        if not ERT_DIR.is_dir():
            raise ErtError(f"Нет каталога {ERT_DIR}")

        pairs, only_txt, only_ert = collect_pairs(args.names or None)
        print_plan(pairs, only_txt, only_ert)

        if not pairs:
            return 0
        if args.dry_run:
            print()
            print("Dry-run: файлы не изменены.")
            return 0

        if not args.yes:
            print()
            answer = input("Записать текст в ERT? [y/N]: ").strip().lower()
            if answer not in ("y", "yes", "д", "да"):
                print("Отменено.")
                return 0

        print()
        for stem, txt_path, ert_path in pairs:
            push_one(txt_path, ert_path)
            print(f"OK: {txt_path.name} -> {ert_path.name}")
        return 0
    except (ErtError, OSError, KeyboardInterrupt) as exc:
        if isinstance(exc, KeyboardInterrupt):
            print("\nОтменено.", file=sys.stderr)
            return 130
        print(f"Ошибка: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
