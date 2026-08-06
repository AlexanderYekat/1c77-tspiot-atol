#!/usr/bin/env python3
"""
Автосинхронизация модулей 1cv77/*.txt <-> 1cv77-extforms/*.ert по версии.

Для каждой сопоставленной пары (одинаковое имя без расширения):
  - если версия в TXT новее  -> push (txt -> ert)
  - если версия в ERT новее  -> pull (ert -> txt)
  - если равны / нет версий -> пропуск

Версия ищется в таком порядке:
  1) глЭтотМодульВерсияМодуля = "..."
  2) комментарий "ver: YYYY.MM.DD-N"
  3) комментарий "версия YYYY.MM.DD-N"

Примеры:
    python sync_modules.py              # план + вопрос
    python sync_modules.py --dry-run    # только план
    python sync_modules.py -y           # выполнить без вопроса
    python sync_modules.py frATOL54_comm
"""

from __future__ import annotations

import argparse
import re
import sys
from dataclasses import dataclass
from pathlib import Path

from ert_module import ErtError, extract_text
from pull_ert_to_txt import pull_one
from push_txt_to_ert import push_one

ROOT = Path(__file__).resolve().parent
TXT_DIR = ROOT / "1cv77"
ERT_DIR = ROOT / "1cv77-extforms"

# Формат: 2026.08.06-03
VERSION_RE = re.compile(r"^(\d{4})\.(\d{2})\.(\d{2})-(\d+)$")

ASSIGN_RE = re.compile(
    r'глЭтотМодульВерсияМодуля\s*=\s*"([^"]+)"',
    re.IGNORECASE,
)
VER_COMMENT_RE = re.compile(
    r"\bver:\s*(\d{4}\.\d{2}\.\d{2}-\d+)\b",
    re.IGNORECASE,
)
VERSION_WORD_RE = re.compile(
    r"версия\s+(\d{4}\.\d{2}\.\d{2}-\d+)\b",
    re.IGNORECASE,
)


@dataclass(frozen=True)
class ModuleVersion:
    raw: str | None
    key: tuple[int, ...]  # () = неизвестна / отсутствует

    @property
    def known(self) -> bool:
        return bool(self.key)

    def label(self) -> str:
        return self.raw if self.raw else "-"


def parse_version_key(raw: str) -> tuple[int, ...] | None:
    match = VERSION_RE.match(raw.strip())
    if not match:
        return None
    return tuple(int(x) for x in match.groups())


def extract_version(text: str) -> ModuleVersion:
    """Достать версию модуля из исходника."""
    assign = ASSIGN_RE.search(text)
    if assign:
        raw = assign.group(1).strip()
        key = parse_version_key(raw)
        if key is not None:
            return ModuleVersion(raw, key)

    head = "\n".join(text.splitlines()[:40])
    for pattern in (VER_COMMENT_RE, VERSION_WORD_RE):
        match = pattern.search(head)
        if match:
            raw = match.group(1).strip()
            key = parse_version_key(raw)
            if key is not None:
                return ModuleVersion(raw, key)

    if assign:
        return ModuleVersion(assign.group(1).strip(), ())
    return ModuleVersion(None, ())


def read_txt_version(path: Path) -> tuple[str, ModuleVersion]:
    raw = path.read_bytes()
    for encoding in ("utf-8-sig", "cp1251"):
        try:
            text = raw.decode(encoding)
            break
        except UnicodeDecodeError:
            text = None
    else:
        raise ErtError(f"Не удалось прочитать {path} (нужен UTF-8 или CP1251)")
    return text, extract_version(text)


def read_ert_version(path: Path) -> tuple[str, ModuleVersion]:
    text = extract_text(path)
    return text, extract_version(text)


@dataclass
class SyncDecision:
    stem: str
    txt_path: Path
    ert_path: Path
    txt_ver: ModuleVersion
    ert_ver: ModuleVersion
    action: str  # push | pull | skip
    reason: str


def decide(stem: str, txt_path: Path, ert_path: Path) -> SyncDecision:
    _, txt_ver = read_txt_version(txt_path)
    _, ert_ver = read_ert_version(ert_path)

    if not txt_ver.known and not ert_ver.known:
        return SyncDecision(
            stem, txt_path, ert_path, txt_ver, ert_ver, "skip", "версии не найдены"
        )
    if txt_ver.known and not ert_ver.known:
        return SyncDecision(
            stem,
            txt_path,
            ert_path,
            txt_ver,
            ert_ver,
            "push",
            "в ERT нет версии, в TXT есть",
        )
    if ert_ver.known and not txt_ver.known:
        return SyncDecision(
            stem,
            txt_path,
            ert_path,
            txt_ver,
            ert_ver,
            "pull",
            "в TXT нет версии, в ERT есть",
        )
    if txt_ver.key > ert_ver.key:
        return SyncDecision(
            stem, txt_path, ert_path, txt_ver, ert_ver, "push", "TXT новее"
        )
    if ert_ver.key > txt_ver.key:
        return SyncDecision(
            stem, txt_path, ert_path, txt_ver, ert_ver, "pull", "ERT новее"
        )
    return SyncDecision(
        stem, txt_path, ert_path, txt_ver, ert_ver, "skip", "версии равны"
    )


def collect_pairs(names: list[str] | None = None) -> tuple[list[tuple[str, Path, Path]], list[str], list[str]]:
    txt_by_stem = {p.stem: p for p in TXT_DIR.glob("*.txt")}
    ert_by_stem = {p.stem: p for p in ERT_DIR.glob("*.ert")}

    if names:
        wanted = set(names)
        unknown = wanted - set(txt_by_stem) - set(ert_by_stem)
        if unknown:
            raise ErtError("Не найдены имена среди txt/ert: " + ", ".join(sorted(unknown)))
        stems = sorted(wanted & set(txt_by_stem) & set(ert_by_stem))
        only_txt = sorted(wanted & set(txt_by_stem) - set(ert_by_stem))
        only_ert = sorted(wanted & set(ert_by_stem) - set(txt_by_stem))
    else:
        stems = sorted(set(txt_by_stem) & set(ert_by_stem))
        only_txt = sorted(set(txt_by_stem) - set(ert_by_stem))
        only_ert = sorted(set(ert_by_stem) - set(txt_by_stem))

    pairs = [(stem, txt_by_stem[stem], ert_by_stem[stem]) for stem in stems]
    return pairs, only_txt, only_ert


def print_plan(
    decisions: list[SyncDecision],
    only_txt: list[str],
    only_ert: list[str],
) -> None:
    print("Автосинк по версии: 1cv77 <-> 1cv77-extforms")
    print()
    if decisions:
        print(f"{'модуль':<32} {'TXT':<16} {'ERT':<16} действие")
        print("-" * 80)
        for d in decisions:
            arrow = {
                "push": "TXT -> ERT",
                "pull": "ERT -> TXT",
                "skip": "пропуск",
            }[d.action]
            print(
                f"{d.stem:<32} {d.txt_ver.label():<16} {d.ert_ver.label():<16} "
                f"{arrow}  ({d.reason})"
            )
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


def apply_decision(d: SyncDecision) -> None:
    if d.action == "push":
        push_one(d.txt_path, d.ert_path)
        print(f"OK push: {d.txt_path.name} -> {d.ert_path.name}")
    elif d.action == "pull":
        pull_one(d.ert_path, d.txt_path)
        print(f"OK pull: {d.ert_path.name} -> {d.txt_path.name}")


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Синхронизация модулей 1cv77 <-> 1cv77-extforms по версии"
    )
    parser.add_argument(
        "names",
        nargs="*",
        help="имена модулей без расширения; если не указаны — все сопоставленные",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="только показать план, ничего не менять",
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
        decisions = [decide(stem, txt, ert) for stem, txt, ert in pairs]
        print_plan(decisions, only_txt, only_ert)

        todo = [d for d in decisions if d.action in ("push", "pull")]
        if not todo:
            print()
            print("Нечего синхронизировать.")
            return 0
        if args.dry_run:
            print()
            print(f"Dry-run: было бы изменений: {len(todo)}")
            return 0

        if not args.yes:
            print()
            answer = input(f"Выполнить {len(todo)} перенос(ов)? [y/N]: ").strip().lower()
            if answer not in ("y", "yes", "д", "да"):
                print("Отменено.")
                return 0

        print()
        for d in todo:
            apply_decision(d)
        return 0
    except (ErtError, OSError, KeyboardInterrupt) as exc:
        if isinstance(exc, KeyboardInterrupt):
            print("\nОтменено.", file=sys.stderr)
            return 130
        print(f"Ошибка: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
