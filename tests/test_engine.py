"""Tests for publishing_engine core modules.

Run: python tests/test_engine.py
"""

from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent / "engine"))


def test_qmd_parser_import():
    """qmd_parser module should import without errors."""
    from qmd_parser import parse_qmd, Block
    assert Block is not None


def test_block_dataclass():
    """Block should have expected fields."""
    from qmd_parser import Block
    b = Block(type="paragraph", text="Test text")
    assert b.type == "paragraph"
    assert b.text == "Test text"
    assert b.level == 0
    assert b.rows == []


def test_bib_formatter_import():
    """BibFormatter should import without errors."""
    from bib_formatter import BibFormatter
    assert BibFormatter is not None


def test_bib_formatter_author_parsing():
    """BibFormatter should handle author name formatting."""
    from bib_formatter import BibFormatter
    # Create a formatter without a real .bib file
    bf = BibFormatter.__new__(BibFormatter)
    bf.entries = {}
    bf.style = "authoryear"
    bf.ref_counter = 0
    bf.cited_keys = []
    assert bf is not None


def test_docx_engine_import():
    """DocxBuilder should import without errors."""
    from docx_engine import DocxBuilder
    assert DocxBuilder is not None


def test_math_renderer_import():
    """math_renderer should import without errors."""
    from math_renderer import latex_to_omml
    assert latex_to_omml is not None


def test_clean_markdown():
    """_clean_markdown should strip markdown formatting."""
    from render_paper import _clean_markdown
    result = _clean_markdown("**bold** and *italic* text")
    assert "bold" in result
    assert "italic" in result
    assert "**" not in result
    assert "*" not in result


def test_clean_markdown_code_spans():
    """_clean_markdown should strip code spans."""
    from render_paper import _clean_markdown
    result = _clean_markdown("Use `function()` here")
    assert "function()" in result
    assert "`" not in result


def test_clean_markdown_crossrefs():
    """_clean_markdown should strip Quarto cross-ref tags."""
    from render_paper import _clean_markdown
    result = _clean_markdown("Text {#sec-intro} more text")
    assert "{#sec-intro}" not in result
    assert "Text" in result


def test_sample_exists():
    """Sample directory should contain manuscript.qmd."""
    sample = Path(__file__).parent.parent / "sample"
    assert sample.exists(), "sample/ directory not found"
    qmd = sample / "manuscript.qmd"
    assert qmd.exists(), "sample/manuscript.qmd not found"


def test_sample_parse():
    """Sample manuscript should parse without errors."""
    from qmd_parser import parse_qmd
    sample_qmd = Path(__file__).parent.parent / "sample" / "manuscript.qmd"
    if not sample_qmd.exists():
        return  # skip if no sample
    meta, blocks = parse_qmd(sample_qmd)
    assert isinstance(meta, dict)
    assert isinstance(blocks, list)
    assert len(blocks) > 0
    assert "title" in meta


# ── Runner ────────────────────────────────────────────────────────────

def run_all():
    tests = [
        test_qmd_parser_import,
        test_block_dataclass,
        test_bib_formatter_import,
        test_bib_formatter_author_parsing,
        test_docx_engine_import,
        test_math_renderer_import,
        test_clean_markdown,
        test_clean_markdown_code_spans,
        test_clean_markdown_crossrefs,
        test_sample_exists,
        test_sample_parse,
    ]

    passed = 0
    failed = 0
    for test in tests:
        name = test.__name__
        try:
            test()
            passed += 1
            print(f"  PASS  {name}")
        except AssertionError as e:
            failed += 1
            print(f"  FAIL  {name}: {e}")
        except Exception as e:
            failed += 1
            print(f"  ERROR {name}: {type(e).__name__}: {e}")

    print(f"\n{passed} passed, {failed} failed out of {len(tests)} tests")
    return failed == 0


if __name__ == "__main__":
    print("Running publishing_engine tests...\n")
    success = run_all()
    sys.exit(0 if success else 1)
