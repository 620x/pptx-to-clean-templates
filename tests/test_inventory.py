from x6pptxgen.audit.checks import run_checks
from x6pptxgen.template.inventory import build_inventory


def test_default_template_inventory(default_pptx):
    inv = build_inventory(default_pptx)

    assert inv["inventory_version"]
    assert len(inv["masters"]) == 1
    layouts = inv["masters"][0]["layouts"]
    assert len(layouts) == 11  # Office default template

    # Every placeholder has a type, an idx, and a resolved geometry source.
    for layout in layouts:
        for ph in layout["placeholders"]:
            assert ph["type"]
            assert isinstance(ph["idx"], int)
            assert ph["geometry_source"] in ("layout", "master", "missing")
            assert not ph["duplicate_idx"]

    # The default template has no example slides.
    assert inv["slides"] == []


def test_default_template_checks_have_no_blockers(default_pptx):
    inv = build_inventory(default_pptx)
    findings = run_checks(inv)
    assert not [f for f in findings if f["severity"] == "blocker"]
