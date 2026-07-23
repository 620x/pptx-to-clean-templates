"""Человекочитаемый отчёт аудита в консоль (rich).

Отчёт — витрина продукта: он должен объяснять «что это значит для генерации»,
а не вываливать XML-термины. Machine-readable JSON пишет cli.py отдельно.
"""

from __future__ import annotations

from rich.console import Console
from rich.panel import Panel
from rich.table import Table

_SEVERITY_STYLE = {
    "blocker": ("БЛОКЕР", "bold white on red"),
    "warning": ("ВНИМАНИЕ", "black on yellow"),
    "info": ("ИНФО", "cyan"),
}


def render_report(inv: dict, findings: list[dict], console: Console | None = None) -> None:
    console = console or Console()

    size = inv["slide_size"]
    n_layouts = sum(len(m["layouts"]) for m in inv["masters"])
    console.print(Panel(
        f"[bold]{inv['source_file']}[/bold]\n"
        f"Слайд: {size['inches']}\" ({size['orientation']})  ·  "
        f"мастеров: {len(inv['masters'])}  ·  layouts: {n_layouts}  ·  "
        f"слайдов-примеров: {len(inv['slides'])}",
        title="Аудит шаблона", border_style="blue",
    ))

    # --- layouts ---
    table = Table(title="Layouts", show_lines=False)
    table.add_column("id", style="dim")
    table.add_column("имя")
    table.add_column("плейсхолдеры (тип:idx)")
    table.add_column("исп. слайдами", justify="center")
    for master in inv["masters"]:
        for layout in master["layouts"]:
            phs = "  ".join(
                f"{p['type']}:{p['idx']}" + ("[red]![/red]" if p["duplicate_idx"] else "")
                for p in layout["placeholders"]
            )
            used = ", ".join(map(str, layout["used_by_slides"])) or "—"
            table.add_row(layout["id"], layout["name"], phs or "[dim]нет[/dim]", used)
    console.print(table)

    # --- слайды-примеры ---
    if inv["slides"]:
        table = Table(title="Слайды-примеры")
        table.add_column("№", justify="right")
        table.add_column("layout")
        table.add_column("ph с контентом", justify="center")
        table.add_column("плавающих боксов", justify="center")
        table.add_column("картинки")
        table.add_column("шрифты runs")
        for s in inv["slides"]:
            used = sum(1 for p in s["placeholders_used"] if p["has_content"])
            cells = []
            if s.get("background_fill") == "picture":
                cells.append("[red]фон-заливка[/red]")
            for p in s["pictures"]:
                size = f"{p['pct_w']}x{p['pct_h']}%" if p["pct_w"] is not None else "?x?%"
                mark = " [red](фон)[/red]" if p["full_bleed"] else ""
                mark += " [dim](ph)[/dim]" if p["is_placeholder"] else ""
                cells.append(size + mark)
            table.add_row(
                str(s["n"]), s["layout_id"], str(used),
                str(len(s["floating_text_boxes"])), ", ".join(cells) or "—",
                ", ".join(s["run_fonts"]) or "—",
            )
        console.print(table)

    # --- находки ---
    console.print()
    for f in findings:
        label, style = _SEVERITY_STYLE[f["severity"]]
        console.print(f"[{style}] {label} [/{style}] [bold]{f['title']}[/bold]")
        console.print(f"  {f['explanation']}", style="dim", highlight=False)
        console.print()

    blockers = sum(1 for f in findings if f["severity"] == "blocker")
    warnings = sum(1 for f in findings if f["severity"] == "warning")
    if blockers:
        console.print(Panel(
            f"[bold red]Шаблон собран вручную — сначала нужен разовый онбординг[/bold red] "
            f"(блокеров: {blockers}, предупреждений: {warnings}).\n"
            "Прямо сейчас генерация воспроизвела бы «пустой» стиль, поэтому мы честно "
            "останавливаемся. Лечение: дизайн переносится со слайдов в layouts — "
            "один раз для бренда, дальше все презентации собираются в фирменном стиле. "
            "Диагноз выше — это и есть план лечения.",
            border_style="red",
        ))
    elif warnings:
        console.print(Panel(
            f"[bold yellow]Шаблон пригоден с оговорками[/bold yellow] — предупреждений: {warnings}.",
            border_style="yellow",
        ))
    else:
        console.print(Panel("[bold green]Шаблон чистый — генерация без оговорок.[/bold green]",
                            border_style="green"))
