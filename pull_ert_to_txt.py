#!/usr/bin/env python3
"""
Перенос кода модулей: 1cv77-extforms/*.ert  →  1cv77/*.txt

Сопоставление по имени файла без расширения
(например razresh_regim_check.ert ↔ razresh_regim_check.txt).

Примеры:
    python pull_ert_to_txt.py              # показать пары и спросить подтверждение
    python pull_ert_to_txt.py --dry-run    # только показать, что будет сделано
    python pull_ert_to_txt.py --yes        # перенести без вопроса
    python pull_ert_to_txt.py razresh_regim_check  # только указанные имена
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from ert_module import ErtError, extract_text

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

    pairs = [(stem, ert_by_stem[stem], txt_by_stem[stem]) for stem in stems]
    return pairs, only_txt, only_ert


def print_plan(pairs: list[tuple[str, Path, Path]], only_txt: list[str], only_ert: list[str]) -> None:
    print(f"Направление: {ERT_DIR.name}/*.ert  ->  {TXT_DIR.name}/*.txt")
    print()
    if pairs:
        print(f"Сопоставлено ({len(pairs)}):")
        for stem, ert_path, txt_path in pairs:
            print(f"  {ert_path.name}  ->  {txt_path.name}")
    else:
        print("Сопоставленных пар нет.")

    if only_ert:
        print()
        print(f"Только в {ERT_DIR.name} (пропуск):")
        for stem in only_ert:
            print(f"  {stem}.ert")

    if only_txt:
        print()
        print(f"Только в {TXT_DIR.name} (пропуск):")
        for stem in only_txt:
            print(f"  {stem}.txt")


def pull_one(ert_path: Path, txt_path: Path) -> None:
    text = extract_text(ert_path)
    txt_path.write_text(text, encoding="utf-8-sig", newline="")


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Извлечь текст модулей из 1cv77-extforms/*.ert в 1cv77/*.txt"
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
            answer = input("Перезаписать TXT из ERT? [y/N]: ").strip().lower()
            if answer not in ("y", "yes", "д", "да"):
                print("Отменено.")
                return 0

        print()
        for stem, ert_path, txt_path in pairs:
            pull_one(ert_path, txt_path)
            print(f"OK: {ert_path.name} -> {txt_path.name}")
        return 0
    except (ErtError, OSError, KeyboardInterrupt) as exc:
        if isinstance(exc, KeyboardInterrupt):
            print("\nОтменено.", file=sys.stderr)
            return 130
        print(f"Ошибка: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
