"""Поиск файлов шрифтов по имени семейства.

Первый кирпич модуля замера (decisions D3): сейчас используется визуальным
редактором (@font-face с настоящим шрифтом шаблона), позже — точным замером
текста через fontTools.
"""

from __future__ import annotations

from pathlib import Path

_FONT_DIRS = [
    Path("/System/Library/Fonts"),
    Path("/System/Library/Fonts/Supplemental"),
    Path("/Library/Fonts"),
    Path.home() / "Library/Fonts",
]

# .ttf/.otf браузеры берут в @font-face надёжно; .ttc — хуже, поэтому в конце
_SUFFIX_PRIORITY = {".ttf": 0, ".otf": 1, ".ttc": 2}


def find_font_file(typeface: str) -> Path | None:
    """Файл шрифта по имени семейства (эвристика по имени файла).

    Предпочитаем regular-начертание: «TenorSans-Regular.ttf» лучше, чем
    «TenorSans-Bold.ttf», а точное совпадение имени — лучше любого суффикса.
    """
    needle = typeface.lower().replace(" ", "")
    candidates: list[tuple[int, int, Path]] = []
    for font_dir in _FONT_DIRS:
        if not font_dir.is_dir():
            continue
        for f in font_dir.iterdir():
            if f.suffix.lower() not in _SUFFIX_PRIORITY:
                continue
            stem = f.stem.lower().replace(" ", "")
            if needle not in stem:
                continue
            # штраф за не-regular начертания
            weight_penalty = 0 if stem in (needle, needle + "-regular",
                                           needle + "regular") else 1
            if any(w in stem for w in ("bold", "italic", "light", "thin",
                                       "medium", "black", "condensed")):
                weight_penalty = 2
            candidates.append(
                (weight_penalty, _SUFFIX_PRIORITY[f.suffix.lower()], f))
    if not candidates:
        return None
    return sorted(candidates)[0][2]
