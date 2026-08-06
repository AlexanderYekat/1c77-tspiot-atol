#!/usr/bin/env python3
"""
Extract and replace source text of a 1C:Enterprise 7.7 external report (*.ert).

No third-party packages are required.

Examples:
    python ert_module.py extract hello.ert module.txt
    python ert_module.py replace hello.ert module.txt hello_new.ert
    python ert_module.py show hello.ert

Пакетный перенос между каталогами (отдельные скрипты, чтобы не перепутать направление):
    python push_txt_to_ert.py --dry-run   # 1cv77/*.txt -> 1cv77-extforms/*.ert
    python pull_ert_to_txt.py --dry-run   # 1cv77-extforms/*.ert -> 1cv77/*.txt
    python sync_modules.py --dry-run      # авто по версии (куда новее)

The module text is stored in the OLE/CFB stream "MD Programm text" as
raw DEFLATE data encoded in Windows-1251.
"""

from __future__ import annotations

import argparse
import math
import struct
import sys
import zlib
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable

SIGNATURE = bytes.fromhex("D0CF11E0A1B11AE1")
FREESECT = 0xFFFFFFFF
ENDOFCHAIN = 0xFFFFFFFE
FATSECT = 0xFFFFFFFD
DIFSECT = 0xFFFFFFFC
NOSTREAM = 0xFFFFFFFF
MODULE_STREAM = "MD Programm text"


class ErtError(Exception):
    """Readable error for invalid/unsupported ERT files."""


@dataclass
class DirectoryEntry:
    entry_id: int
    name: str
    object_type: int
    start_sector: int
    size: int
    raw: bytearray


class CompoundFile:
    """Small CFB/OLE reader and writer sufficient for 1C 7.7 ERT files."""

    def __init__(self, data: bytes):
        if len(data) < 512 or data[:8] != SIGNATURE:
            raise ErtError("Файл не является OLE/CFB-контейнером 1С 7.7")

        self.data = data
        self.header = bytearray(data[:512])
        self.major_version = self._u16(0x1A)
        self.byte_order = self._u16(0x1C)
        self.sector_shift = self._u16(0x1E)
        self.mini_sector_shift = self._u16(0x20)

        if self.major_version != 3 or self.byte_order != 0xFFFE:
            raise ErtError(
                f"Поддерживается CFB v3 (512-байтовые сектора), найден CFB v{self.major_version}"
            )

        self.sector_size = 1 << self.sector_shift
        self.mini_sector_size = 1 << self.mini_sector_shift
        self.mini_stream_cutoff = self._u32(0x38)
        if self.sector_size != 512 or self.mini_sector_size != 64:
            raise ErtError("Неожиданный размер сектора OLE-контейнера")

        self.fat = self._read_fat()
        self.directory_entries = self._read_directory()
        if not self.directory_entries or self.directory_entries[0].object_type != 5:
            raise ErtError("В OLE-контейнере отсутствует Root Entry")

        self.mini_fat = self._read_mini_fat()
        root = self.directory_entries[0]
        self.mini_stream = self._read_regular_stream(root.start_sector, root.size)

    def _u16(self, offset: int) -> int:
        return struct.unpack_from("<H", self.data, offset)[0]

    def _u32(self, offset: int) -> int:
        return struct.unpack_from("<I", self.data, offset)[0]

    def _sector(self, sector_id: int) -> bytes:
        start = (sector_id + 1) * self.sector_size
        end = start + self.sector_size
        if start < self.sector_size or end > len(self.data):
            raise ErtError(f"Некорректная ссылка на сектор {sector_id}")
        return self.data[start:end]

    @staticmethod
    def _chain(start: int, table: list[int], label: str) -> list[int]:
        if start in (FREESECT, ENDOFCHAIN):
            return []
        result: list[int] = []
        seen: set[int] = set()
        current = start
        while current not in (FREESECT, ENDOFCHAIN):
            if current in seen:
                raise ErtError(f"Цикл в цепочке {label}")
            if current >= len(table):
                raise ErtError(f"Некорректная цепочка {label}: сектор {current}")
            seen.add(current)
            result.append(current)
            current = table[current]
        return result

    def _read_fat(self) -> list[int]:
        number_of_fat_sectors = self._u32(0x2C)
        first_difat_sector = self._u32(0x44)
        number_of_difat_sectors = self._u32(0x48)

        difat = list(struct.unpack_from("<109I", self.data, 0x4C))
        fat_sector_ids = [x for x in difat if x != FREESECT]

        current = first_difat_sector
        entries_per_difat = self.sector_size // 4 - 1
        for _ in range(number_of_difat_sectors):
            sector = self._sector(current)
            values = struct.unpack("<%dI" % (entries_per_difat + 1), sector)
            fat_sector_ids.extend(x for x in values[:-1] if x != FREESECT)
            current = values[-1]

        fat_sector_ids = fat_sector_ids[:number_of_fat_sectors]
        if len(fat_sector_ids) != number_of_fat_sectors:
            raise ErtError("Повреждена таблица DIFAT/FAT")

        fat: list[int] = []
        entries_per_fat = self.sector_size // 4
        for sector_id in fat_sector_ids:
            fat.extend(struct.unpack("<%dI" % entries_per_fat, self._sector(sector_id)))
        return fat

    def _read_regular_stream(self, start_sector: int, size: int | None = None) -> bytes:
        chunks = [self._sector(x) for x in self._chain(start_sector, self.fat, "FAT")]
        result = b"".join(chunks)
        return result if size is None else result[:size]

    def _read_directory(self) -> list[DirectoryEntry]:
        first_directory_sector = self._u32(0x30)
        raw_directory = self._read_regular_stream(first_directory_sector)
        result: list[DirectoryEntry] = []

        for offset in range(0, len(raw_directory), 128):
            raw = bytearray(raw_directory[offset : offset + 128])
            if len(raw) < 128:
                break
            name_length = struct.unpack_from("<H", raw, 64)[0]
            if 2 <= name_length <= 64:
                name = bytes(raw[: name_length - 2]).decode("utf-16le", errors="replace")
            else:
                name = ""
            result.append(
                DirectoryEntry(
                    entry_id=offset // 128,
                    name=name,
                    object_type=raw[66],
                    start_sector=struct.unpack_from("<I", raw, 116)[0],
                    size=struct.unpack_from("<Q", raw, 120)[0],
                    raw=raw,
                )
            )
        return result

    def _read_mini_fat(self) -> list[int]:
        first_mini_fat_sector = self._u32(0x3C)
        number_of_mini_fat_sectors = self._u32(0x40)
        if number_of_mini_fat_sectors == 0:
            return []

        sector_ids = self._chain(first_mini_fat_sector, self.fat, "MiniFAT")
        sector_ids = sector_ids[:number_of_mini_fat_sectors]
        result: list[int] = []
        entries_per_sector = self.sector_size // 4
        for sector_id in sector_ids:
            result.extend(struct.unpack("<%dI" % entries_per_sector, self._sector(sector_id)))
        return result

    def find_stream(self, name: str) -> DirectoryEntry:
        matches = [x for x in self.directory_entries if x.object_type == 2 and x.name == name]
        if not matches:
            raise ErtError(f'Поток "{name}" не найден')
        if len(matches) > 1:
            raise ErtError(f'Найдено несколько потоков "{name}"')
        return matches[0]

    def read_stream(self, entry: DirectoryEntry) -> bytes:
        if entry.object_type != 2:
            raise ErtError(f'"{entry.name}" не является потоком')
        if entry.size == 0:
            return b""
        if entry.size < self.mini_stream_cutoff:
            mini_ids = self._chain(entry.start_sector, self.mini_fat, "mini stream")
            chunks = [
                self.mini_stream[x * self.mini_sector_size : (x + 1) * self.mini_sector_size]
                for x in mini_ids
            ]
            return b"".join(chunks)[: entry.size]
        return self._read_regular_stream(entry.start_sector, entry.size)

    def all_stream_data(self) -> dict[int, bytes]:
        return {
            entry.entry_id: self.read_stream(entry)
            for entry in self.directory_entries
            if entry.object_type == 2
        }

    def rebuild(self, replacements: dict[int, bytes]) -> bytes:
        """Rebuild the whole CFB container, allowing stream sizes to change."""
        stream_data = self.all_stream_data()
        stream_data.update(replacements)

        sector_size = self.sector_size
        mini_sector_size = self.mini_sector_size
        entries_per_fat_sector = sector_size // 4
        entries_per_difat_sector = entries_per_fat_sector - 1

        small_stream_ids = [
            e.entry_id
            for e in self.directory_entries
            if e.object_type == 2 and len(stream_data[e.entry_id]) < self.mini_stream_cutoff
        ]
        large_stream_ids = [
            e.entry_id
            for e in self.directory_entries
            if e.object_type == 2 and len(stream_data[e.entry_id]) >= self.mini_stream_cutoff
        ]

        mini_fat_values: list[int] = []
        mini_stream_parts: list[bytes] = []
        stream_start: dict[int, int] = {}

        for entry_id in small_stream_ids:
            payload = stream_data[entry_id]
            if not payload:
                stream_start[entry_id] = ENDOFCHAIN
                continue
            count = math.ceil(len(payload) / mini_sector_size)
            first = len(mini_fat_values)
            stream_start[entry_id] = first
            for index in range(count):
                mini_fat_values.append(first + index + 1 if index + 1 < count else ENDOFCHAIN)
                part = payload[index * mini_sector_size : (index + 1) * mini_sector_size]
                mini_stream_parts.append(part.ljust(mini_sector_size, b"\x00"))

        mini_stream_payload = b"".join(mini_stream_parts)

        directory_size = max(1, math.ceil(len(self.directory_entries) * 128 / sector_size)) * sector_size
        mini_fat_sector_count = (
            math.ceil(len(mini_fat_values) * 4 / sector_size) if mini_fat_values else 0
        )

        large_sector_counts = {
            entry_id: math.ceil(len(stream_data[entry_id]) / sector_size)
            for entry_id in large_stream_ids
        }
        root_sector_count = (
            math.ceil(len(mini_stream_payload) / sector_size) if mini_stream_payload else 0
        )
        directory_sector_count = directory_size // sector_size

        base_sector_count = (
            sum(large_sector_counts.values())
            + root_sector_count
            + directory_sector_count
            + mini_fat_sector_count
        )

        fat_sector_count = 0
        difat_sector_count = 0
        while True:
            total = base_sector_count + fat_sector_count + difat_sector_count
            new_fat = math.ceil(total / entries_per_fat_sector)
            new_difat = (
                math.ceil(max(new_fat - 109, 0) / entries_per_difat_sector)
                if new_fat > 109
                else 0
            )
            if (new_fat, new_difat) == (fat_sector_count, difat_sector_count):
                break
            fat_sector_count, difat_sector_count = new_fat, new_difat

        sectors: list[bytearray] = []
        chains: list[list[int]] = []

        def allocate(payload: bytes, count: int) -> list[int]:
            if count == 0:
                return []
            ids = list(range(len(sectors), len(sectors) + count))
            for index in range(count):
                part = payload[index * sector_size : (index + 1) * sector_size]
                sectors.append(bytearray(part.ljust(sector_size, b"\x00")))
            chains.append(ids)
            return ids

        for entry_id in large_stream_ids:
            ids = allocate(stream_data[entry_id], large_sector_counts[entry_id])
            stream_start[entry_id] = ids[0] if ids else ENDOFCHAIN

        root_chain = allocate(mini_stream_payload, root_sector_count)

        # Directory entries are patched after all starts are known.
        directory_placeholder_index = len(sectors)
        directory_chain = allocate(b"\x00" * directory_size, directory_sector_count)

        mini_fat_bytes = b""
        if mini_fat_sector_count:
            capacity = mini_fat_sector_count * entries_per_fat_sector
            padded = mini_fat_values + [FREESECT] * (capacity - len(mini_fat_values))
            mini_fat_bytes = struct.pack("<%dI" % capacity, *padded)
        mini_fat_chain = allocate(mini_fat_bytes, mini_fat_sector_count)

        difat_sector_ids = list(
            range(len(sectors), len(sectors) + difat_sector_count)
        )
        for _ in difat_sector_ids:
            sectors.append(bytearray(sector_size))

        fat_sector_ids = list(range(len(sectors), len(sectors) + fat_sector_count))
        for _ in fat_sector_ids:
            sectors.append(bytearray(sector_size))

        total_sector_count = len(sectors)
        fat_capacity = fat_sector_count * entries_per_fat_sector
        fat_values = [FREESECT] * fat_capacity

        def mark_chain(ids: Iterable[int]) -> None:
            ids = list(ids)
            for index, sector_id in enumerate(ids):
                fat_values[sector_id] = ids[index + 1] if index + 1 < len(ids) else ENDOFCHAIN

        for chain_ids in chains:
            mark_chain(chain_ids)
        for sector_id in difat_sector_ids:
            fat_values[sector_id] = DIFSECT
        for sector_id in fat_sector_ids:
            fat_values[sector_id] = FATSECT

        if total_sector_count > fat_capacity:
            raise ErtError("Внутренняя ошибка расчёта FAT")

        # Patch directory entries while preserving names, tree links, CLSIDs and timestamps.
        directory_bytes = bytearray(directory_size)
        for entry in self.directory_entries:
            raw = bytearray(entry.raw)
            if entry.entry_id == 0:
                root_start = root_chain[0] if root_chain else ENDOFCHAIN
                struct.pack_into("<I", raw, 116, root_start)
                struct.pack_into("<Q", raw, 120, len(mini_stream_payload))
            elif entry.object_type == 2:
                struct.pack_into("<I", raw, 116, stream_start[entry.entry_id])
                struct.pack_into("<Q", raw, 120, len(stream_data[entry.entry_id]))
            start = entry.entry_id * 128
            directory_bytes[start : start + 128] = raw

        for index, sector_id in enumerate(directory_chain):
            part = directory_bytes[index * sector_size : (index + 1) * sector_size]
            sectors[sector_id][:] = part

        # DIFAT sectors contain additional FAT sector IDs and a pointer to the next DIFAT sector.
        remaining_fat_ids = fat_sector_ids[109:]
        for index, sector_id in enumerate(difat_sector_ids):
            chunk = remaining_fat_ids[
                index * entries_per_difat_sector : (index + 1) * entries_per_difat_sector
            ]
            values = chunk + [FREESECT] * (entries_per_difat_sector - len(chunk))
            next_sector = (
                difat_sector_ids[index + 1]
                if index + 1 < len(difat_sector_ids)
                else ENDOFCHAIN
            )
            values.append(next_sector)
            sectors[sector_id][:] = struct.pack(
                "<%dI" % entries_per_fat_sector, *values
            )

        for index, sector_id in enumerate(fat_sector_ids):
            chunk = fat_values[
                index * entries_per_fat_sector : (index + 1) * entries_per_fat_sector
            ]
            sectors[sector_id][:] = struct.pack(
                "<%dI" % entries_per_fat_sector, *chunk
            )

        header = bytearray(self.header)
        struct.pack_into("<I", header, 0x28, 0)  # directory sectors for CFB v3
        struct.pack_into("<I", header, 0x2C, fat_sector_count)
        struct.pack_into("<I", header, 0x30, directory_chain[0])
        struct.pack_into("<I", header, 0x34, 0)  # transaction signature
        struct.pack_into(
            "<I", header, 0x3C, mini_fat_chain[0] if mini_fat_chain else ENDOFCHAIN
        )
        struct.pack_into("<I", header, 0x40, mini_fat_sector_count)
        struct.pack_into(
            "<I", header, 0x44, difat_sector_ids[0] if difat_sector_ids else ENDOFCHAIN
        )
        struct.pack_into("<I", header, 0x48, difat_sector_count)

        header_difat = fat_sector_ids[:109] + [FREESECT] * max(0, 109 - len(fat_sector_ids))
        struct.pack_into("<109I", header, 0x4C, *header_difat[:109])

        return bytes(header) + b"".join(bytes(x) for x in sectors)


def decompress_module(payload: bytes) -> bytes:
    try:
        return zlib.decompress(payload, wbits=-15)
    except zlib.error as exc:
        raise ErtError(f"Не удалось распаковать поток {MODULE_STREAM}: {exc}") from exc


def compress_module(source: bytes) -> bytes:
    compressor = zlib.compressobj(level=9, wbits=-15)
    return compressor.compress(source) + compressor.flush()


def load_ert(path: Path) -> tuple[CompoundFile, DirectoryEntry]:
    try:
        data = path.read_bytes()
    except OSError as exc:
        raise ErtError(f"Не удалось прочитать {path}: {exc}") from exc
    compound = CompoundFile(data)
    return compound, compound.find_stream(MODULE_STREAM)


def extract_text(ert_path: Path) -> str:
    compound, entry = load_ert(ert_path)
    module_bytes = decompress_module(compound.read_stream(entry))
    try:
        return module_bytes.decode("cp1251")
    except UnicodeDecodeError as exc:
        raise ErtError("Текст модуля не декодируется как Windows-1251") from exc


def read_source_file(path: Path) -> str:
    raw = path.read_bytes()
    # UTF-8/UTF-8 BOM is the default exchange format. CP1251 is accepted for convenience.
    for encoding in ("utf-8-sig", "cp1251"):
        try:
            return raw.decode(encoding)
        except UnicodeDecodeError:
            pass
    raise ErtError(f"Файл {path} должен быть в UTF-8 или Windows-1251")


def normalize_newlines(text: str) -> str:
    return text.replace("\r\n", "\n").replace("\r", "\n").replace("\n", "\r\n")


def replace_text(ert_path: Path, source_path: Path, output_path: Path) -> None:
    compound, entry = load_ert(ert_path)
    text = normalize_newlines(read_source_file(source_path))
    try:
        source_bytes = text.encode("cp1251")
    except UnicodeEncodeError as exc:
        bad = text[exc.start : exc.end]
        line = text.count("\n", 0, exc.start) + 1
        col = exc.start - (text.rfind("\n", 0, exc.start) + 1) + 1
        raise ErtError(
            f"Символ {bad!r} (U+{ord(bad[0]):04X}) в строке {line}:{col} "
            f"нельзя записать в модуль 1С 7.7 (нужна Windows-1251)"
        ) from exc

    compressed = compress_module(source_bytes)
    rebuilt = compound.rebuild({entry.entry_id: compressed})

    try:
        output_path.write_bytes(rebuilt)
    except OSError as exc:
        raise ErtError(f"Не удалось записать {output_path}: {exc}") from exc

    # Immediate structural and content verification.
    check = CompoundFile(rebuilt)
    check_entry = check.find_stream(MODULE_STREAM)
    check_source = decompress_module(check.read_stream(check_entry))
    if check_source != source_bytes:
        raise ErtError("Проверка созданного ERT не пройдена: текст после записи отличается")


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Извлечение и замена текста модуля в *.ert 1С:Предприятия 7.7"
    )
    subparsers = parser.add_subparsers(dest="command", required=True)

    show = subparsers.add_parser("show", help="показать текст модуля в консоли")
    show.add_argument("ert", type=Path)

    extract = subparsers.add_parser("extract", help="извлечь текст модуля в UTF-8")
    extract.add_argument("ert", type=Path)
    extract.add_argument("text", type=Path)

    replace = subparsers.add_parser("replace", help="заменить текст модуля и создать новый ERT")
    replace.add_argument("ert", type=Path, help="исходный ERT")
    replace.add_argument("text", type=Path, help="текст в UTF-8 или CP1251")
    replace.add_argument("output", type=Path, help="новый ERT (не исходный файл)")

    return parser


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)
    try:
        if args.command == "show":
            sys.stdout.write(extract_text(args.ert))
        elif args.command == "extract":
            text = extract_text(args.ert)
            args.text.write_text(text, encoding="utf-8-sig", newline="")
            print(f"Извлечено: {args.text}")
        elif args.command == "replace":
            if args.ert.resolve() == args.output.resolve():
                raise ErtError("Выходной файл должен отличаться от исходного ERT")
            replace_text(args.ert, args.text, args.output)
            print(f"Создано и проверено: {args.output}")
        return 0
    except (ErtError, OSError) as exc:
        print(f"Ошибка: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
