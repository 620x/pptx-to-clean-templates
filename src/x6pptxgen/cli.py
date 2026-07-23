"""Command-line interface.

    x6 audit <file.pptx|.potx> [--json report.json]
        Inventory a template's masters/layouts/placeholders and run
        anti-pattern checks; print a human-readable report.

    x6 heal  <file.pptx|.potx> [-o healed.pptx] [--probe preview.pptx]
        Turn hand-built repeated slides into real reusable layouts and
        write a new .pptx. With --probe, also emit a one-slide-per-layout
        preview deck.

Note: audit findings and healed placeholder prompts are currently in
Russian (the engine's origin). Translations are welcome — see the README.
"""

from __future__ import annotations

import json
from pathlib import Path

import typer
from rich.console import Console

from x6pptxgen.audit.checks import run_checks
from x6pptxgen.audit.report import render_report
from x6pptxgen.template.inventory import build_inventory

app = typer.Typer(
    add_completion=False,
    help="Analyze and heal PowerPoint templates into real reusable layouts.",
)


@app.callback()
def _main() -> None:
    """Empty callback keeps a multi-command shape (don't collapse to `x6 <file>`)."""


@app.command()
def audit(
    file: Path = typer.Argument(..., exists=True, readable=True, help="Template .pptx/.potx"),
    json_out: Path | None = typer.Option(
        None, "--json", help="Also write the machine-readable inventory here",
    ),
) -> None:
    """Audit a template: layout/placeholder inventory + anti-pattern checks."""
    inv = build_inventory(file)
    findings = run_checks(inv)
    render_report(inv, findings, Console())

    if json_out is not None:
        json_out.parent.mkdir(parents=True, exist_ok=True)
        payload = {**inv, "findings": findings}
        json_out.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
        Console().print(f"Inventory written: [bold]{json_out}[/bold]")


@app.command()
def heal(
    file: Path = typer.Argument(..., exists=True, readable=True, help="Template .pptx/.potx"),
    out: Path = typer.Option(Path("healed.pptx"), "--out", "-o", help="Output .pptx"),
    probe: Path | None = typer.Option(
        None, "--probe", help="Also write a one-slide-per-layout preview deck here",
    ),
) -> None:
    """Heal hand-built slides into real reusable layouts."""
    from x6pptxgen.heal.engine import build_probe, heal_template

    console = Console()
    report = heal_template(file, out)
    console.print(
        f"Healed: [bold]{report['out']}[/bold] — "
        f"{report['families']} layout family(ies) from {report['slides_total']} slides."
    )
    for layout in report["layouts"]:
        console.print(
            f"  • {layout['name']}: {len(layout['placeholders'])} placeholder(s), "
            f"{layout['photo_zones']} photo zone(s) [from slides {layout['from_slides']}]"
        )
    for w in report["warnings"]:
        console.print(f"  [yellow]![/yellow] {w}")

    if probe is not None:
        n = build_probe(out, probe, name_prefix=report["menu_prefix"])
        console.print(f"Probe deck: [bold]{probe}[/bold] — {n} slide(s).")


if __name__ == "__main__":
    app()
