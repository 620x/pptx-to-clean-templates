"""Проверки-антипаттерны поверх инвентаря шаблона.

Каждая проверка отвечает на один вопрос: «сможет ли генератор, который собирает
слайды ТОЛЬКО из layouts и заполняет ТОЛЬКО плейсхолдеры, воспроизвести стиль
этого шаблона?» — и объясняет пользователю последствия человеческим языком.

Severity:
- blocker: строгая генерация по этому шаблону не даст ожидаемого результата;
- warning: генерация возможна, но с оговорками/сниженным качеством;
- info:    полезно знать, на генерацию не влияет.

Числовой балл 0-100 сознательно отложен до прогона на 3-5 реальных шаблонах
(см. docs/decisions.md, D7) — пока веса были бы выдуманными.
"""

from __future__ import annotations

TITLE_TYPES = {"title", "ctrTitle"}


def run_checks(inv: dict) -> list[dict]:
    findings: list[dict] = []
    for check in (
        _structure_problems,
        _design_lives_on_slides,
        _templatability,
        _no_usable_layouts,
        _duplicate_placeholder_idx,
        _scaffolding_hardcodes_fonts,
        _slide_fonts_differ_from_theme,
        _no_picture_placeholders,
        _fonts_missing_for_measurement,
        _pictures_baked_into_layouts,
        _unused_layouts,
        _embedded_fonts_note,
        _orientation_note,
    ):
        findings.extend(check(inv))
    order = {"blocker": 0, "warning": 1, "info": 2}
    return sorted(findings, key=lambda f: order[f["severity"]])


def _all_layouts(inv: dict) -> list[dict]:
    return [layout for master in inv["masters"] for layout in master["layouts"]]


def _finding(severity: str, code: str, title: str, explanation: str, details=None) -> dict:
    f = {"severity": severity, "code": code, "title": title, "explanation": explanation}
    if details is not None:
        f["details"] = details
    return f


# --- блокеры -----------------------------------------------------------------

def _structure_problems(inv: dict) -> list[dict]:
    """Деградации, зафиксированные при построении инвентаря (битые связи и т.п.)."""
    return [
        _finding(
            "warning", "structure-degraded",
            "Файл с особенностями структуры",
            problem,
        )
        for problem in inv.get("problems", [])
    ]


def _decor_full_bleed(slide: dict) -> bool:
    """Full-bleed фото, которое НЕ является использованием плейсхолдера.

    Фото в picture placeholder — использование шаблона по назначению.
    Картинки в группах не считаем: их масштаб скорректирован по ext/chExt,
    но повороты/сдвиги делают оценку ненадёжной — для блокера лучше недобдеть.
    """
    return any(
        p["full_bleed"] and not p["is_placeholder"] and not p["in_group"]
        for p in slide["pictures"]
    )


def _design_lives_on_slides(inv: dict) -> list[dict]:
    """Главный антипаттерн: слайды построены плавающими боксами/фонами, не плейсхолдерами."""
    slides = inv["slides"]
    if not slides:
        return []
    no_ph = [s for s in slides if not any(p["has_content"] for p in s["placeholders_used"])]
    dirty = [
        s for s in no_ph
        if len(s["floating_text_boxes"]) >= 2
        or _decor_full_bleed(s)
        or s.get("background_fill") == "picture"  # «Формат фона > рисунок» на слайде
    ]
    if len(dirty) < max(1, len(slides) // 2):
        return []

    # Если в файле УЖЕ есть брендированные пригодные layouts (плейсхолдеры +
    # крупная фотография/фон-картинка) — например, после лечения шаблона —
    # грязные примеры становятся историей: генерация через layouts воспроизведёт
    # стиль, а примеры при генерации удаляются. Тогда это warning, не blocker.
    branded_usable = []
    for layout in _all_layouts(inv):
        types = {p["type"] for p in layout["placeholders"] if p["is_content"]}
        usable = bool(types & TITLE_TYPES) and bool(types - TITLE_TYPES)
        # «Дизайн» = фон-картинка или фото площадью >=20% слайда. Порог по
        # площади, а не по одной стороне: полоска-водяной знак 60%x4% — это
        # 2.4% площади, а не дизайн (находка ревью — иначе один такой layout
        # снимал блокер с грязного шаблона).
        has_design = (
            layout.get("background_fill") == "picture"
            or any(
                (p["pct_w"] or 0) * (p["pct_h"] or 0) >= 2000
                for p in layout["pictures"]
            )
        )
        if usable and has_design:
            branded_usable.append(layout["id"])
    if branded_usable:
        return [_finding(
            "warning", "design-on-slides-historic",
            "Слайды-примеры собраны вручную, но брендированные layouts есть",
            "Примеры игнорируют плейсхолдеры (исторический ручной стиль), однако "
            "в файле есть layouts с дизайном и плейсхолдерами — генерация через "
            "них воспроизводит стиль, а слайды-примеры при генерации удаляются.",
            {"branded_usable_layouts": branded_usable},
        )]

    return [_finding(
        "blocker", "design-on-slides",
        "Дизайн живёт на слайдах, а не в layouts",
        "Слайды-примеры собраны из плавающих текст-боксов и фоновых картинок поверх "
        "слайда, плейсхолдеры layouts не используются. Генератор, собирающий слайды "
        "из layouts, воспроизвёл бы «пустую» версию оформления — поэтому мы честно "
        "останавливаемся (конкуренты на таком файле молча выдают безликую презу). "
        "Это лечится разовой процедурой онбординга: дизайн переносится со слайдов "
        "в layouts, после чего генерация воспроизводит фирменный стиль. Насколько "
        "файл поддаётся автоматическому лечению — см. находку «Шаблонизируемость».",
        {
            "slides_without_placeholders": [s["n"] for s in no_ph],
            "slides_with_full_bleed_photos": [s["n"] for s in slides if _decor_full_bleed(s)],
            "slides_with_picture_background": [
                s["n"] for s in slides if s.get("background_fill") == "picture"
            ],
            "floating_boxes_per_slide": {s["n"]: len(s["floating_text_boxes"]) for s in slides},
        },
    )]


def _no_usable_layouts(inv: dict) -> list[dict]:
    usable = []
    for layout in _all_layouts(inv):
        types = {p["type"] for p in layout["placeholders"] if p["is_content"]}
        has_title = bool(types & TITLE_TYPES)
        has_content = bool(types - TITLE_TYPES)
        if has_title and has_content:
            usable.append(layout["id"])
    if usable:
        return []
    return [_finding(
        "blocker", "no-usable-layouts",
        "Нет ни одного пригодного layout",
        "Пригодный layout = хотя бы один титульный плейсхолдер плюс хотя бы один "
        "контентный. Без таких layouts генератору некуда класть текст.",
    )]


def _slide_signature(s: dict) -> tuple:
    """Структурная сигнатура слайда: слайды одного «архетипа» (копипаст-вёрстка)
    дают одинаковые сигнатуры. Логотипы и мелкие иконки игнорируем."""
    pics = tuple(sorted(
        ((p["pct_w"] or 0) // 25, (p["pct_h"] or 0) // 25)
        for p in s["pictures"]
        if (p["pct_w"] or 0) * (p["pct_h"] or 0) >= 25
    ))
    return (
        # число текст-боксов бакетами по 2: 5 и 6 боксов — один архетип
        min(len(s["floating_text_boxes"]), 7) // 2,
        sum(1 for p in s["placeholders_used"] if p["has_content"]),
        pics,
        s.get("background_fill") == "picture",
    )


def _templatability(inv: dict) -> list[dict]:
    """Оценка «лечится ли файл автоматически»: доля слайдов, повторяющих
    структуру друг друга. Ручная вёрстка почти всегда копипастная — и именно
    повторяемость позволяет перенести дизайн в layouts полуавтоматически.
    Показываем только для «грязных» файлов: чистым шаблонам лечение не нужно."""
    slides = inv["slides"]
    if len(slides) < 4:
        return []
    if any(p["has_content"] for s in slides for p in s["placeholders_used"]):
        return []

    clusters: dict[tuple, int] = {}
    for s in slides:
        sig = _slide_signature(s)
        clusters[sig] = clusters.get(sig, 0) + 1
    repeated = sum(n for n in clusters.values() if n >= 2)
    families = sum(1 for n in clusters.values() if n >= 2)
    share = repeated / len(slides)

    if share >= 0.6:
        grade, meaning = "высокая", (
            "дизайн переносится в layouts преимущественно автоматически — "
            "онбординг быстрый и делается один раз для бренда."
        )
    elif share >= 0.3:
        grade, meaning = "средняя", (
            "часть слайдов уникальна — онбординг полуавтоматический, "
            "с ручной проверкой спорных страниц."
        )
    else:
        grade, meaning = "низкая", (
            "структуры слайдов почти не повторяются — лечение потребует "
            "заметной ручной работы дизайнера."
        )
    return [_finding(
        "info", "templatability",
        f"Шаблонизируемость: {grade} "
        f"({repeated} из {len(slides)} слайдов в {families} повторяющихся семействах)",
        f"Повторяемость структуры слайдов — мера того, насколько файл поддаётся "
        f"автоматическому переносу дизайна в layouts. Здесь {meaning} "
        f"Оценка по одному файлу — нижняя граница: у брендов с серией презентаций "
        f"одни и те же страницы повторяются между файлами, и лечение делается "
        f"один раз по всей серии.",
        {"repeated": repeated, "total": len(slides), "families": families,
         "share": round(share, 2)},
    )]


# --- предупреждения ----------------------------------------------------------

def _duplicate_placeholder_idx(inv: dict) -> list[dict]:
    findings = []
    for layout in _all_layouts(inv):
        dupes = sorted({p["idx"] for p in layout["placeholders"] if p["duplicate_idx"]})
        if dupes:
            findings.append(_finding(
                "warning", "duplicate-ph-idx",
                f"Дубли idx плейсхолдеров в layout {layout['id']} ({layout['name']})",
                "Несколько плейсхолдеров с одинаковым idx — наследование и адресация "
                "неоднозначны; генератор не должен использовать эти idx как цели.",
                {"layout": layout["id"], "idx": dupes},
            ))
    return findings


def _scaffolding_hardcodes_fonts(inv: dict) -> list[dict]:
    ratio = inv["fonts"]["usage"]["scaffolding_literal_ratio"]
    if ratio <= 0.5:
        return []
    return [_finding(
        "warning", "hardcoded-fonts-in-scaffolding",
        f"Мастер/layouts хардкодят шрифты ({ratio:.0%} литеральных ссылок)",
        "Вместо ссылок на тему (+mj-lt/+mn-lt) в каркасе прописаны конкретные "
        "шрифты. Тема «нечестная»: смена шрифтов темы ничего не изменит. Для "
        "генерации терпимо (шрифт фиксирован), но это признак шаблона, собранного "
        "визуально, а не структурно.",
        {"histogram": inv["fonts"]["usage"]["master_layouts"]},
    )]


def _slide_fonts_differ_from_theme(inv: dict) -> list[dict]:
    # Сверяем шрифты каждого слайда с темой ЕГО мастера (мультимастерные шаблоны
    # реальны; slide_masters[0] здесь — классическая ошибка, см. decisions D7).
    themes = inv["fonts"]["themes_by_master"]
    alien_by_slide: dict[int, list[str]] = {}
    for s in inv["slides"]:
        master_id = s["layout_id"].split("/")[0]
        theme = themes.get(master_id, {"major_latin": "", "minor_latin": ""})
        theme_faces = {theme["major_latin"], theme["minor_latin"]}
        alien = sorted(
            f for f in s["run_fonts"] if f not in theme_faces and not f.startswith("+")
        )
        if alien:
            alien_by_slide[s["n"]] = alien
    if not alien_by_slide:
        return []
    all_alien = sorted({f for faces in alien_by_slide.values() for f in faces})
    return [_finding(
        "warning", "slide-fonts-differ-from-theme",
        f"Реальные шрифты слайдов не совпадают с темой: {', '.join(all_alien)}",
        "Текст на слайдах-примерах набран шрифтами, которых нет в теме их мастера. "
        "Слайды, созданные генератором через layouts, унаследуют шрифты темы/мастера — "
        "и будут выглядеть иначе, чем примеры. Реальный стиль живёт в ручном "
        "форматировании.",
        {"themes_by_master": themes, "alien_fonts_by_slide": alien_by_slide},
    )]


def _no_picture_placeholders(inv: dict) -> list[dict]:
    for layout in _all_layouts(inv):
        if any(p["type"] == "pic" for p in layout["placeholders"]):
            return []
    return [_finding(
        "warning", "no-picture-placeholders",
        "Ни в одном layout нет picture placeholder",
        "Пользовательские картинки некуда вставлять штатным способом. Запасной путь — "
        "вставка по геометрии контентного плейсхолдера (add_picture), см. decisions D4.",
    )]


def _fonts_missing_for_measurement(inv: dict) -> list[dict]:
    missing = sorted(
        face for face, found in inv["fonts"]["system_availability_heuristic"].items()
        if not found
    )
    if not missing:
        return []
    return [_finding(
        "warning", "fonts-missing-for-measurement",
        f"Шрифты не найдены в системе: {', '.join(missing)}",
        "Замер текста (защита от переполнения) будет работать по метрик-клону или "
        "запасному шрифту с увеличенным запасом — лимиты символов будут жёстче. "
        "Эвристика по именам файлов; точный резолвинг появится в модуле замера.",
        {"missing": missing},
    )]


def _pictures_baked_into_layouts(inv: dict) -> list[dict]:
    findings = []
    for layout in _all_layouts(inv):
        has_photo_bg = (
            any(p["full_bleed"] for p in layout["pictures"])
            or layout.get("background_fill") == "picture"
        )
        if has_photo_bg:
            findings.append(_finding(
                "info", "layout-background-photo",
                f"Фоновая картинка в layout {layout['id']} ({layout['name']})",
                "Фото запечено в layout: слайды на его основе унаследуют его "
                "автоматически — для генерации это хорошо. Заменить фото на свой "
                "можно будет только подменой image part (v1.1).",
            ))
    return findings


# --- инфо --------------------------------------------------------------------

def _unused_layouts(inv: dict) -> list[dict]:
    unused = [
        f"{l['id']} ({l['name']})" for l in _all_layouts(inv) if not l["used_by_slides"]
    ]
    if not unused:
        return []
    return [_finding(
        "info", "unused-layouts",
        f"Layouts, не использованные слайдами-примерами: {len(unused)}",
        "Не ошибка: это кандидаты для генерации. Но если примеры игнорируют layouts "
        "целиком — их оформление, скорее всего, никто не поддерживал.",
        {"unused": unused},
    )]


def _embedded_fonts_note(inv: dict) -> list[dict]:
    embedded = inv["fonts"]["embedded"]
    if not embedded:
        return []
    faces = ", ".join(f"{e['typeface']} ({'/'.join(e['styles'])})" for e in embedded)
    return [_finding(
        "info", "embedded-fonts",
        f"В файл встроены шрифты: {faces}",
        "Встроенные шрифты (EOT, обычно сабсеты) сохранятся в выходном файле как есть. "
        "Для замера текста они не используются: сабсет без нужных глифов молча "
        "искажает ширины (см. decisions D3).",
    )]


def _orientation_note(inv: dict) -> list[dict]:
    size = inv["slide_size"]
    if size["orientation"] != "portrait":
        return []
    return [_finding(
        "info", "portrait-format",
        f"Вертикальный формат слайда ({size['inches']}\")",
        "Портретная ориентация. На генерацию не влияет (все расчёты идут от "
        "фактического p:sldSz), но рендер-проверку и превью стоит смотреть глазами.",
    )]
