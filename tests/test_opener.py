from pptx import Presentation

from x6pptxgen.template.opener import load_template


def test_opens_plain_pptx(default_pptx):
    prs = load_template(default_pptx)
    assert len(prs.slide_masters) == 1


def test_opens_potx_via_content_type_patch(default_potx):
    # Прямое открытие через python-pptx обязано падать (это и есть причина патча) —
    # если однажды перестанет, патч можно будет убрать.
    try:
        Presentation(str(default_potx))
        raise AssertionError("python-pptx неожиданно открыл .potx — проверь версию")
    except ValueError:
        pass

    prs = load_template(default_potx)
    assert len(prs.slide_masters) == 1


def test_source_file_is_not_modified(default_potx):
    before = default_potx.read_bytes()
    load_template(default_potx)
    assert default_potx.read_bytes() == before
