# Changelog

## 2026.04.2 (2026-04-17)

### Added
- `--validate` flag: run ai_style_checker before rendering
- `--threshold` flag: block render if AI score exceeds limit
- 3-state return from `_run_pre_render_check()` (True/False/None)
- `--all` mode correctly blocks papers exceeding threshold
- protocol.py: ManuscriptProtocol for pipeline integration

## 2026.04.1 (2026-04-17)

### Added
- Initial release: .qmd to 7 DOCX document types
- Native OMML equations (MathML to OMML)
- Booktabs-style tables, embedded figures
- Bibliography from .bib (author-year and numbered styles)
- Auto cross-references, Track Changes XML
- Response to reviewers with grey comment boxes
