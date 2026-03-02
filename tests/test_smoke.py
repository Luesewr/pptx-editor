"""Smoke tests for pptx_editor package."""

import pptx_editor


def test_version_exists() -> None:
    assert hasattr(pptx_editor, "__version__")


def test_version_format() -> None:
    parts = pptx_editor.__version__.split(".")
    assert len(parts) == 3
    assert all(part.isdigit() for part in parts)
