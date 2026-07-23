# pptx-to-clean-templates

**Analyze any PowerPoint file and turn hand-built slides into real, reusable layouts — automatically.**

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)
![Python](https://img.shields.io/badge/python-3.12%2B-blue.svg)

**English** · [Русский](README.ru.md)

AI slide generators can apply a *brand kit* — your colors, fonts and logo — but they can't
reproduce **your** corporate template's actual layouts. And most real corporate decks are worse
than that: they're built by hand, as floating text boxes and full-bleed background images, with
**no real layouts at all**. Point any "generate from your template" tool at one and you get a
bland, empty-looking result.

`pptx-to-clean-templates` fixes the input. It:

1. **Audits** any `.pptx`/`.potx` and tells you, in plain language, whether it's usable, what was
   built by hand, and how automatable a fix is.
2. **Heals** it — detects repeated slide *types* (cover, day card, hotel page, …) and rewrites
   them into genuine PowerPoint layouts (`p:sldLayout` parts on the master), carrying over the
   fonts, colors, backgrounds and paragraph-level typography so new slides inherit the brand
   **by construction**.

The output is a normal `.pptx` you can open in PowerPoint and build on — or feed to any generator
that fills placeholders.

## Why it's different

Most tools stop at the **brand kit** (colors + fonts + logo). This one reconstructs the **layout
structure** itself, from a file that may have had none — the hard, unglamorous OOXML surgery that
"apply my brand" tools skip.

## Install

```bash
pip install git+https://github.com/620x/pptx-to-clean-templates
```

Or for development:

```bash
git clone https://github.com/620x/pptx-to-clean-templates
cd pptx-to-clean-templates
uv sync            # or: pip install -e ".[dev]"
```

Requires Python 3.12+.

## CLI

```bash
# Diagnose a template
x6 audit deck.pptx
x6 audit deck.pptx --json report.json

# Heal hand-built slides into real layouts
x6 heal deck.pptx -o healed.pptx --probe preview.pptx
```

`--probe` writes a one-slide-per-layout preview so you can see what the reconstructed layouts look
like before generating anything.

## Library

```python
from x6pptxgen.template.inventory import build_inventory
from x6pptxgen.audit.checks import run_checks
from x6pptxgen.heal.engine import heal_template, build_probe

inv = build_inventory("deck.pptx")          # structured template inventory (JSON-serializable)
findings = run_checks(inv)                    # anti-pattern findings, human-readable

report = heal_template("deck.pptx", "healed.pptx")
build_probe("healed.pptx", "preview.pptx")    # one slide per new layout
```

## How healing works

1. **Cluster** example slides by a structural signature — copy-paste layout produces repeating
   families, and those families become your layouts.
2. **Pick a donor** slide per family (the fullest example) and **assign roles** to its text zones
   by size, position and length — title, kicker, subtitle, body, footnote — never by reading the
   content.
3. **Rebuild** each family into a real layout: static decor is copied (with image/hyperlink
   relations re-bound), content zones become placeholders with typography lifted into `lstStyle`
   so it inherits, the slide background is transplanted, and tables become column placeholders.
   When the file runs out of unused layouts, new layout parts are created at the OPC level.

## Honest limitations

- Role assignment is **heuristic** — tuned on real decks, but not perfect on every file. The probe
  deck exists so you can eyeball the result.
- **Text-fit / overflow measurement is not solved here.** `measure/` currently only locates font
  files; precise capacity measurement (via fontTools) is future work.
- The system-font availability heuristic looks in **macOS font directories**; on other OSes it
  under-reports (healing itself is unaffected).
- Audit findings and healed placeholder prompts are currently written in **Russian** (the engine's
  origin). **i18n is a great first contribution** — the strings live in `audit/checks.py`,
  `audit/report.py` and `heal/engine.py`.

## Scope

This is the engine extracted from a larger project. The AI content-generation, brief pipeline and
web editor are intentionally **not** included — this repo is the template analysis + healing core,
usable on its own.

## License

MIT © 2026 Ilia Mogirev
