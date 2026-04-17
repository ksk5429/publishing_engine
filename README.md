# Publishing Engine

Python-based manuscript rendering pipeline for generating publication-quality DOCX files for SCI journal submission. No Quarto or LaTeX dependency -- builds Word documents entirely from Python.

## Features

- **Native OMML equations** -- editable in Word, rendered via MathML to OMML
- **Booktabs-style tables** -- navy headers, alternating rows, no vertical borders
- **Embedded figures** -- PNG/JPG with auto-numbered captions
- **Bibliography from .bib** -- pybtex parses BibTeX; supports `(Author, Year)` or `[1]` numbered styles
- **Auto cross-references** -- `@fig-label`, `@eq-label`, `@tbl-label` to `Fig. 3`, `Eq. (2)`, `Table 1`
- **Track Changes XML** -- revised paragraphs marked with `w:ins` elements
- **Word/page count validation** -- auto-check against journal limits
- **Pre-render AI validation** -- `--validate` flag runs ai_style_checker before rendering
- **7 output documents** per manuscript:

| Output | Description |
|--------|-------------|
| `manuscript_full.docx` | Complete manuscript with inline figures |
| `manuscript_manuscript.docx` | Text-only with figure placeholders |
| `manuscript_revised.docx` | Revised with red text + Track Changes |
| `figures_and_tables.docx` | Companion document -- full-width figures/tables |
| `response_to_reviewers.docx` | Point-by-point with grey comment boxes |
| `supplementary_material.docx` | Extended tables and methods |
| `cover_letter.docx` | Formal cover letter |

## Quick Start

```bash
pip install python-docx pybtex latex2mathml lxml pypandoc

# Render all 7 documents for a manuscript
python engine/render_paper.py sample

# Render with AI style validation (blocks if score > threshold)
python engine/render_paper.py sample --validate --threshold 30

# Render all papers with validation
python engine/render_paper.py --all --validate

# Render companion documents
python engine/render_figures_tables.py sample
python engine/render_response.py sample
python engine/render_coverletter.py sample
python engine/render_supplementary.py sample
```

## Pre-render Validation Hook

The `--validate` flag automatically runs [ai_style_checker](https://github.com/ksk5429/ai_style_checker) before rendering:

```bash
# Check + render (continues regardless of score)
python engine/render_paper.py paperB --validate

# Gate: block render if AI score exceeds threshold
python engine/render_paper.py paperB --validate --threshold 30

# Works with --all (skips papers exceeding threshold)
python engine/render_paper.py --all --validate --threshold 30
```

The hook:
- Auto-discovers ai_style_checker in sibling directories
- Returns `True` (passed), `False` (blocked), or `None` (checker not available)
- Saves `style_report.json` to the paper's `_output/` directory
- Logs warnings at WARNING level on checker failures

## Architecture

```
publishing_engine/
├── engine/                      # Core rendering modules
│   ├── docx_engine.py           # DocxBuilder class (title page, abstract, TOC)
│   ├── qmd_parser.py            # .qmd to structured blocks
│   ├── bib_formatter.py         # .bib to citations + bibliography
│   ├── math_renderer.py         # LaTeX to MathML to OMML
│   ├── equation_handler.py      # Math renderer + pandoc fallback
│   ├── render_paper.py          # Main renderer + --validate hook
│   ├── render_figures_tables.py # Companion figures DOCX
│   ├── render_response.py       # Response to reviewers
│   ├── render_supplementary.py  # Supplementary material
│   └── render_coverletter.py    # Cover letter
├── protocol.py                  # ManuscriptProtocol for pipeline integration
├── sample/                      # Demo manuscript
├── requirements.txt
└── README.md
```

## Manuscript Source Format

Content is written in `.qmd` (Quarto-compatible Markdown) with YAML frontmatter:

```yaml
---
title: "Paper Title"
author:
  - name: Author Name
    affiliation: University
keywords: [keyword1, keyword2]
highlights: [highlight1, highlight2]
nomenclature:
  - ["S/D", "--", "Scour depth ratio"]
limits:
  max_words: 9000
citation_style: authoryear  # or "numbered"
bibliography: references.bib
---
```

## Dependencies

- Python 3.10+
- python-docx, pybtex, latex2mathml, lxml, pypandoc

## Ecosystem

| Repo | Purpose |
|------|---------|
| [ai_style_checker](https://github.com/ksk5429/ai_style_checker) | 12-checker AI detection + fingerprinting |
| [sentence_evolver](https://github.com/ksk5429/sentence_evolver) | 10-persona sentence rewriting + A/B scoring |
| **publishing_engine** | .qmd to publication DOCX (this repo) |
| [manuscript_pipeline](https://github.com/ksk5429/manuscript_pipeline) | Orchestrator chaining all engines |
| [pdf_search_engine](https://github.com/ksk5429/pdf_search_engine) | Academic PDF discovery + markdown conversion |

## License

Apache 2.0
