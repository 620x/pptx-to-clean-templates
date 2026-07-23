"""Shared test fixtures.

All unit tests run on synthetic decks (python-pptx's default template),
so the suite ships no .pptx files.
"""

from __future__ import annotations

import io
from pathlib import Path

import pytest
from pptx import Presentation


@pytest.fixture()
def default_pptx(tmp_path: Path) -> Path:
    """python-pptx's default template saved to a file (synthetic fixture)."""
    path = tmp_path / "default.pptx"
    Presentation().save(path)
    return path


@pytest.fixture()
def default_potx(tmp_path: Path) -> Path:
    """Same default template, but with a .potx content-type (for the opener test)."""
    buf = io.BytesIO()
    Presentation().save(buf)
    data = buf.getvalue()

    import zipfile

    out = tmp_path / "default.potx"
    with zipfile.ZipFile(io.BytesIO(data)) as src, zipfile.ZipFile(out, "w") as dst:
        for item in src.infolist():
            payload = src.read(item.filename)
            if item.filename == "[Content_Types].xml":
                payload = payload.replace(
                    b"presentationml.presentation.main+xml",
                    b"presentationml.template.main+xml",
                )
            dst.writestr(item, payload)
    return out
