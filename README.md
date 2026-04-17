# Publishing Engine

A Python-based manuscript rendering pipeline for generating publication-quality DOCX files for SCI journal submission. No Quarto or LaTeX dependency — builds Word documents entirely from Python with full control over every element.

## Features

- **Native OMML equations** — editable in Word, rendered via MathML → OMML (Microsoft's MML2OMML.XSL)
- **Booktabs-style tables** — navy headers, alternating rows, center-aligned, no vertical borders
- **Embedded figures** — PNG/JPG with auto-numbered captions
- **Bibliography from .bib** — pybtex parses BibTeX; supports `(Author, Year)` or `[1]` numbered styles
- **Auto cross-references** — `@fig-label`, `@eq-label`, `@tbl-label` → `Fig. 3`, `Eq. (2)`, `Table 1`
- **Track Changes XML** — revised paragraphs marked with `w:ins` elements (author + timestamp)
- **Word/page count validation** — auto-check against journal limits
- **Nomenclature table** — symbol/units/description from YAML
- **7 output documents** per manuscript:

| Output | Description |
|--------|-------------|
| `manuscript_full.docx` | Complete manuscript with inline figures |
| `manuscript_manuscript.docx` | Text-only with figure placeholders |
| `manuscript_revised.docx` | Revised with red text + Track Changes on addressed paragraphs |
| `figures_and_tables.docx` | Companion document — each figure/table at full page width |
| `response_to_reviewers.docx` | Point-by-point with grey comment boxes + blue responses |
| `supplementary_material.docx` | Extended tables and methods |
| `cover_letter.docx` | Formal cover letter |

## Quick Start

```bash
pip install python-docx pybtex latex2mathml lxml pypandoc

# Render all 7 documents for the sample manuscript
cd publishing_engine
python engine/render_paper.py sample
python engine/render_figures_tables.py sample
python engine/render_response.py sample
python engine/render_supplementary.py sample
python engine/render_coverletter.py sample
```

## Architecture

```
publishing_engine/
├── engine/                      # Core rendering modules
│   ├── docx_engine.py           # DocxBuilder class (title page, abstract, TOC, etc.)
│   ├── qmd_parser.py            # .qmd → structured blocks
│   ├── bib_formatter.py         # .bib → citations + bibliography
│   ├── math_renderer.py         # LaTeX → MathML → OMML (from markdocx)
│   ├── equation_handler.py      # Integrates math_renderer + pandoc fallback
│   ├── render_paper.py          # Main renderer (full/manuscript/revised)
│   ├── render_figures_tables.py # Companion figures DOCX
│   ├── render_response.py       # Response to reviewers
│   ├── render_supplementary.py  # Supplementary material
│   └── render_coverletter.py    # Cover letter
│
├── sample/                      # Demo manuscript
│   ├── manuscript.qmd           # Source content (Quarto-compatible markdown)
│   ├── references.bib           # BibTeX bibliography
│   ├── figures/                 # PNG figures
│   ├── Reviewers_Comments.txt   # Sample reviewer comments
│   ├── revision_marks.py        # Keywords for red-marking revised paragraphs
│   └── _output/                 # Generated DOCX files
│
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
    email: author@university.edu
keywords: [keyword1, keyword2]
highlights: [highlight1, highlight2]
nomenclature:
  - ["S/D", "—", "Scour depth ratio"]
limits:
  max_words: 9000
  max_pages: 30
citation_style: authoryear  # or "numbered"
bibliography: references.bib
---
```

## Dependencies

- Python 3.12+
- python-docx 1.2.0
- pybtex 0.26.1
- latex2mathml 3.78.1
- lxml 6.0.2
- pandoc 3.9+ (for fallback equation rendering)
- Microsoft Office (for MML2OMML.XSL — optional, improves equation quality)

## Credits

- Equation rendering adapted from [markdocx](https://github.com/shynneri-source/markdocx)
- Inspired by [Pandoc Scholar](https://pandoc-scholar.github.io/) and [Manubot](https://manubot.org/)

## License

Apache 2.0
