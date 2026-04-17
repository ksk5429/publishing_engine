"""
Unified paper renderer v2: .qmd → publication-quality DOCX.

Fixes from v1:
  1. Abstract: reads YAML string directly (no markdown parsing artifacts)
  2. References: parses .bib with pybtex, renders formatted bibliography
  3. Equations: native OMML via latex2mathml (fallback: italic text)
  4. Citations: @citekey ��� (Author, Year) using bib data
  5. Figures/tables: referenced by number only in main text
     (companion figures_and_tables.docx generated separately)

Usage:
    python _shared/render_paper.py paperB_buckingham_pi
    python _shared/render_paper.py --all
"""
from __future__ import annotations

import argparse
import logging
import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
from docx_engine import DocxBuilder
from qmd_parser import parse_qmd, Block
from bib_formatter import BibFormatter

PAPERS_DIR = Path(__file__).parent.parent

SHARED_BLOCKS = {
    "credit": (
        "Kyeong-Sun Kim: conceptualisation, methodology, software, formal analysis, "
        "investigation, data curation, validation, writing \u2014 original draft, visualisation. "
        "Sung-Ryul Kim: conceptualisation, resources, supervision, writing \u2014 review and "
        "editing, funding acquisition."
    ),
    "coi": (
        "The authors declare no competing financial or personal interests that could "
        "have influenced the work reported in this manuscript."
    ),
    "ai_disclosure": (
        "During the preparation of this work, the authors used generative AI "
        "(Anthropic Claude) for academic writing coaching and scaffolding of Python "
        "figure-generation scripts. All numerical values were independently computed "
        "from the analysis scripts; no AI-generated numerical data appear in the paper. "
        "All cited references were verified against CrossRef. The authors reviewed and "
        "edited all content and take full responsibility for the publication."
    ),
    "acknowledgements": (
        "This work was supported by the KEPCO Research Institute through the collaborative "
        "research agreement with Seoul National University. The authors thank the MMB "
        "consortium for site geotechnical investigation data and Unison Heavy Industries "
        "Corporation for turbine operational data and engineering support."
    ),
}

logging.basicConfig(level=logging.INFO, format="[%(asctime)s] %(message)s", datefmt="%H:%M:%S")
log = logging.getLogger("render_paper")


def _clean_markdown(text: str, bib: BibFormatter | None = None) -> str:
    # Bold/italic markdown
    text = re.sub(r'\*\*(.+?)\*\*', r'\1', text)
    text = re.sub(r'\*(.+?)\*', r'\1', text)
    # Code spans
    text = re.sub(r'`(.+?)`', r'\1', text)
    # Inline math: $...$ → keep the content but make it readable
    # Common substitutions for readability in plain text
    def _math_to_text(m):
        s = m.group(1)
        s = s.replace(r'\pi_6', '\u03C0\u2086')
        s = s.replace(r'\pi', '\u03C0')
        s = s.replace(r'\Pi_1', '\u03A0\u2081')
        s = s.replace(r'\Pi_2', '\u03A0\u2082')
        s = s.replace(r'\Pi_3', '\u03A0\u2083')
        s = s.replace(r'\Pi_4', '\u03A0\u2084')
        s = s.replace(r'\Pi_5', '\u03A0\u2085')
        s = s.replace(r'\Pi_6', '\u03A0\u2086')
        s = s.replace(r'\Pi_7', '\u03A0\u2087')
        s = s.replace(r'\Pi_8', '\u03A0\u2088')
        s = s.replace(r'\Pi_9', '\u03A0\u2089')
        s = s.replace(r'\varphi_1', '\u03C6\u2081')
        s = s.replace(r'\varphi_2', '\u03C6\u2082')
        s = s.replace(r'\varphi_3', '\u03C6\u2083')
        s = s.replace(r'\varphi_4', '\u03C6\u2084')
        s = s.replace(r'\varphi_5', '\u03C6\u2085')
        s = s.replace(r'\sigma', '\u03C3')
        s = s.replace(r'\gamma', '\u03B3')
        s = s.replace(r'\varepsilon', '\u03B5')
        s = s.replace(r'\leq', '\u2264')
        s = s.replace(r'\geq', '\u2265')
        s = s.replace(r'\approx', '\u2248')
        s = s.replace(r'\times', '\u00D7')
        s = s.replace(r'\pm', '\u00B1')
        s = s.replace(r'\infty', '\u221E')
        s = s.replace(r'\alpha', '\u03B1')
        s = s.replace(r'\beta', '\u03B2')
        s = s.replace(r'\delta', '\u03B4')
        s = s.replace(r'\eta', '\u03B7')
        s = s.replace(r'\theta', '\u03B8')
        s = s.replace(r'\lambda', '\u03BB')
        s = s.replace(r'\mu', '\u03BC')
        s = s.replace(r'\rho', '\u03C1')
        s = s.replace(r'\xi', '\u03BE')
        s = s.replace(r'\phi', '\u03C6')
        s = s.replace(r'\psi', '\u03C8')
        s = s.replace(r'\omega', '\u03C9')
        s = s.replace(r'\nu', '\u03BD')
        s = s.replace(r'\kappa', '\u03BA')
        # Fix em-dash rendering (-- → en-dash)
        s = s.replace('--', '\u2013')
        # Clean remaining LaTeX commands
        s = re.sub(r'\\(?:mathrm|text|mathit|mathbf|textit|textbf)\{([^}]*)\}', r'\1', s)
        s = re.sub(r'\\operatorname\{([^}]*)\}', r'\1', s)
        s = re.sub(r'\\(?:left|right|Big|big|bigg)[()|\[\]{}.]?', '', s)
        s = re.sub(r'\\(?:quad|qquad|,|;|!)', ' ', s)
        s = re.sub(r'\\frac\{([^}]*)\}\{([^}]*)\}', r'(\1/\2)', s)
        s = re.sub(r'\\sqrt\{([^}]*)\}', r'sqrt(\1)', s)
        s = re.sub(r'\\bar\{([^}]*)\}', r'\1', s)
        s = re.sub(r'\\hat\{([^}]*)\}', r'\1', s)
        s = re.sub(r'\\tilde\{([^}]*)\}', r'\1', s)
        s = re.sub(r'\\overline\{([^}]*)\}', r'\1', s)
        # Subscript: _{text} → _text (simplified)
        s = re.sub(r'_\{([^}]*)\}', r'_\1', s)
        # Superscript: ^{text} → ^text
        s = re.sub(r'\^\{([^}]*)\}', r'^\1', s)
        # Remove remaining backslash commands
        s = re.sub(r'\\[a-zA-Z]+', '', s)
        # Remove remaining braces
        s = re.sub(r'[{}]', '', s)
        s = s.replace(r'\mathrm{', '').replace('}', '')
        s = re.sub(r'\\(?:text|mathit|mathbf)\{([^}]*)\}', r'\1', s)
        s = s.replace(r'\,', ' ')
        s = s.replace(r'\\', '')
        s = re.sub(r'_\{([^}]*)\}', lambda m2: ''.join(
            chr(0x2080 + int(c)) if c.isdigit() else c for c in m2.group(1)
        ) if m2.group(1).isdigit() else '_' + m2.group(1), s)
        s = re.sub(r'_(\d)', lambda m2: chr(0x2080 + int(m2.group(1))), s)
        s = re.sub(r'\^(\d)', lambda m2: chr(0x2070 + int(m2.group(1))) if int(m2.group(1)) not in (1,) else '\u00B9', s)
        s = s.replace('^2', '\u00B2').replace('^3', '\u00B3')
        s = re.sub(r'\\[a-zA-Z]+', '', s)
        s = re.sub(r'[{}]', '', s)
        return s.strip()
    text = re.sub(r'\$(.+?)\$', _math_to_text, text)
    # Quarto cross-ref and attribute tags
    text = re.sub(r'\{#[\w-]+\}', '', text)
    text = re.sub(r'\{[^}]*width=[^}]*\}', '', text)
    # Citations
    if bib:
        text = bib.resolve_citations(text)
    else:
        text = re.sub(r'\[@[\w;,\s@]+\]', '', text)
        text = re.sub(r'@(\w+)', r'[\1]', text)
    # Cleanup
    text = re.sub(r'\s+', ' ', text).strip()
    return text


def _resolve_crossrefs(text: str, fig_map: dict, tbl_map: dict, eq_map: dict) -> str:
    """Replace @fig-label, @tbl-label, @eq-label with auto-numbered references."""
    def _fig_repl(m):
        label = m.group(1)
        num = fig_map.get(label, "?")
        return f"Fig. {num}"
    def _tbl_repl(m):
        label = m.group(1)
        num = tbl_map.get(label, "?")
        return f"Table {num}"
    def _eq_repl(m):
        label = m.group(1)
        num = eq_map.get(label, "?")
        return f"Eq. ({num})"
    text = re.sub(r'@fig-([\w-]+)', _fig_repl, text)
    text = re.sub(r'@tbl-([\w-]+)', _tbl_repl, text)
    text = re.sub(r'@eq-([\w-]+)', _eq_repl, text)
    # Clean section references: §\ref{sec-*}, [sec]-*, §sec-*
    text = re.sub(r'§?\\ref\{sec-([\w-]+)\}', r'§\1', text)
    text = re.sub(r'\[sec\]-([\w-]+)', r'§\1', text)
    text = re.sub(r'§sec-([\w-]+)', r'§\1', text)
    text = re.sub(r'#sec-([\w-]+)', '', text)
    return text


def render_one(paper_name: str, include_figures: bool = True,
               revision_marks: dict[str, list[str]] | None = None) -> Path:
    """Render a manuscript.

    revision_marks: if provided, maps comment IDs to lists of keyword phrases.
    Any paragraph containing one of those phrases will be rendered in red,
    producing a tracked-changes-style revised manuscript.
    """
    paper_dir = PAPERS_DIR / paper_name
    qmd_path = paper_dir / "manuscript.qmd"
    if not qmd_path.exists():
        raise FileNotFoundError(f"No manuscript.qmd in {paper_dir}")

    meta, blocks = parse_qmd(qmd_path)
    log.info("parsed %s: %d blocks", paper_name, len(blocks))

    # Load bibliography with citation style from YAML
    bib_path = paper_dir / "references.bib"
    citation_style = meta.get("citation_style", "authoryear")
    bib = BibFormatter(bib_path, style=citation_style) if bib_path.exists() else None
    if bib:
        log.info("  loaded %d bib entries from %s", len(bib.db.entries), bib_path.name)

    title = meta.get("title", paper_name)
    authors_raw = meta.get("author", [])
    authors = []
    for a in authors_raw:
        if isinstance(a, dict):
            email = a.get("email", "")
            if "ksk5429" in email:
                email = "kyeongsunkim@snu.ac.kr"
            authors.append({
                "name": a.get("name", ""),
                "affiliation": a.get("affiliation", "Seoul National University"),
                "email": email,
                "corresponding": bool(email),
            })
    if not authors:
        authors = [
            {"name": "Kyeong-Sun Kim", "affiliation": "Seoul National University",
             "email": "kyeongsunkim@snu.ac.kr", "corresponding": True},
            {"name": "Sung-Ryul Kim", "affiliation": "Seoul National University"},
        ]

    # FIX 1: abstract — try YAML first, then extract from body blocks
    abstract = meta.get("abstract", "")
    if isinstance(abstract, str):
        # YAML multi-line strings have literal \n — join into one paragraph
        abstract = " ".join(abstract.strip().split())
    if not abstract:
        # Extract from # Abstract section in body
        in_abstract = False
        abstract_parts = []
        for block in blocks:
            if block.type == "heading" and block.text.lower().strip() == "abstract":
                in_abstract = True
                continue
            if in_abstract:
                if block.type == "heading":
                    break
                if block.type == "paragraph":
                    abstract_parts.append(block.text)
        if abstract_parts:
            abstract = " ".join(abstract_parts)
            abstract = _clean_markdown(abstract, bib)
            log.info("  extracted abstract from body (%d chars)", len(abstract))

    keywords = meta.get("keywords", [])
    highlights = meta.get("highlights", [])

    nomenclature_raw = meta.get("nomenclature", [])
    nomenclature = []
    for item in nomenclature_raw:
        if isinstance(item, (list, tuple)) and len(item) >= 3:
            nomenclature.append((str(item[0]), str(item[1]), str(item[2])))

    short_title = title[:70] + "..." if len(title) > 70 else title
    b = DocxBuilder(
        title=title,
        authors=authors,
        running_header=short_title,
        abstract=abstract,
        keywords=keywords,
        highlights=highlights,
    )

    # Nomenclature table (after ToC, before body)
    if nomenclature:
        b.nomenclature(nomenclature)
        b.page_break()

    # Pre-scan blocks to build cross-reference maps
    fig_map: dict[str, int] = {}
    tbl_map: dict[str, int] = {}
    eq_map: dict[str, int] = {}
    _fig_n = _tbl_n = _eq_n = 0
    for block in blocks:
        if block.type == "figure":
            _fig_n += 1
            label = re.search(r'\{#fig-([\w-]+)\}', block.caption + " " + block.label) if block.label else None
            if not label and block.path:
                label_guess = Path(block.path).stem.replace("fig_", "").replace("fig", "")
                fig_map[label_guess] = _fig_n
            if label:
                fig_map[label.group(1)] = _fig_n
        elif block.type == "table" and block.headers:
            _tbl_n += 1
        elif block.type == "equation":
            _eq_n += 1
            if block.label:
                eq_map[block.label] = _eq_n

    skip_headings = {"abstract", "highlights"}
    eq_num = 0

    for block in blocks:
        if block.type == "heading":
            h_lower = block.text.lower().strip()
            if h_lower in skip_headings:
                continue
            if "acknowledge" in h_lower:
                b.acknowledgements(SHARED_BLOCKS["acknowledgements"])
                continue
            if "generative ai" in h_lower or "ai-assisted" in h_lower:
                continue  # Removed per KSK review — no AI disclosure section
            if "credit" in h_lower:
                b.credit(SHARED_BLOCKS["credit"])
                continue
            if "competing" in h_lower or "conflict" in h_lower:
                b.coi(SHARED_BLOCKS["coi"])
                continue
            if h_lower == "references":
                # FIX 2: render bibliography from .bib
                if bib:
                    refs = bib.get_cited_references()
                    if not refs:
                        refs = bib.get_all_references()
                    b.references_list(refs)
                else:
                    b.heading("References", level=1, numbered=False)
                continue
            b.heading(block.text, level=block.level, numbered=block.numbered)

        elif block.type == "paragraph":
            text = block.text
            if text.startswith("[") and text.endswith("]"):
                continue
            # FIX 4: resolve citations + cross-references
            text = _clean_markdown(text, bib)
            text = _resolve_crossrefs(text, fig_map, tbl_map, eq_map)
            if text:
                # Check if this paragraph matches any revision marker
                is_revised = False
                if revision_marks:
                    text_lower = text.lower()
                    for cid, phrases in revision_marks.items():
                        for phrase in phrases:
                            if phrase.lower() in text_lower:
                                is_revised = True
                                break
                        if is_revised:
                            break

                if is_revised:
                    from docx.shared import Pt as _Pt, RGBColor as _RGB
                    from docx.enum.text import WD_ALIGN_PARAGRAPH as _ALIGN
                    from docx.oxml.ns import qn as _qn
                    from docx.oxml import OxmlElement as _El
                    from datetime import datetime as _dt

                    p = b.doc.add_paragraph()
                    p.paragraph_format.first_line_indent = _Pt(0)
                    p.alignment = _ALIGN.JUSTIFY

                    # Add Track Changes w:ins wrapper
                    ins = _El("w:ins")
                    ins.set(_qn("w:id"), str(hash(text) % 100000))
                    ins.set(_qn("w:author"), "Kyeong-Sun Kim")
                    ins.set(_qn("w:date"), _dt.now().strftime("%Y-%m-%dT%H:%M:%SZ"))

                    r_el = _El("w:r")
                    rPr = _El("w:rPr")
                    rFonts = _El("w:rFonts")
                    rFonts.set(_qn("w:ascii"), "Times New Roman")
                    rFonts.set(_qn("w:hAnsi"), "Times New Roman")
                    rPr.append(rFonts)
                    sz = _El("w:sz")
                    sz.set(_qn("w:val"), "22")
                    rPr.append(sz)
                    color = _El("w:color")
                    color.set(_qn("w:val"), "C62828")
                    rPr.append(color)
                    r_el.append(rPr)
                    t = _El("w:t")
                    t.set(_qn("xml:space"), "preserve")
                    t.text = text
                    r_el.append(t)
                    ins.append(r_el)
                    p._p.append(ins)
                else:
                    b.paragraph(text, indent=True)

        elif block.type == "figure":
            fig_path = paper_dir / block.path
            clean_caption = _clean_markdown(block.caption, bib)
            if include_figures and fig_path.exists():
                b.figure(fig_path, clean_caption)
            else:
                b._fig_num += 1
                b.paragraph(
                    f"[Fig. {b._fig_num}. {clean_caption[:80]} — see companion figures document]",
                    indent=False, italic=True,
                )

        elif block.type == "table":
            if block.headers:
                # Clean LaTeX in table cells
                clean_headers = [_clean_markdown(h) for h in block.headers]
                clean_rows = [[_clean_markdown(cell) for cell in row] for row in block.rows]
                b.table(clean_headers, clean_rows, caption_text=block.caption)

        elif block.type == "equation":
            b.equation(block.latex)

    # Add end-matter if not already added by heading detection
    b.credit(SHARED_BLOCKS["credit"])
    b.coi(SHARED_BLOCKS["coi"])

    # Journal word/page limits (configurable per paper via YAML `limits:`)
    limits = meta.get("limits", None)
    b.add_stats_footer(limits=limits)

    out_dir = paper_dir / "_output"
    out_dir.mkdir(parents=True, exist_ok=True)
    if revision_marks:
        suffix = "_revised"
    elif include_figures:
        suffix = "_full"
    else:
        suffix = "_manuscript"
    out_path = out_dir / f"manuscript{suffix}.docx"
    b.save(out_path)
    log.info("rendered %s -> %s (%d bytes)", paper_name, out_path, out_path.stat().st_size)
    return out_path


PAPER_NAMES = [
    "paperB_buckingham_pi",
    "paperA_dt_decision",
    "paperV2_shm_sage",
    "paperOp3_aes",
]


def _run_pre_render_check(paper_name: str, threshold: float = 0.0) -> bool | None:
    """Pre-render hook: run ai_style_checker before rendering.

    Returns:
        True  — checker ran and passed
        False — checker ran and FAILED threshold
        None  — checker not available (silent pass)
    """
    import subprocess, json, os

    qmd_path = PAPERS_DIR / paper_name / "manuscript.qmd"
    if not qmd_path.exists():
        return None

    checker_candidates = [
        PAPERS_DIR.parent.parent / "ai_style_checker",
        PAPERS_DIR.parent / "ai_style_checker",
        Path(__file__).parent.parent.parent / "ai_style_checker",
    ]
    checker_path = next((p for p in checker_candidates if (p / "cli.py").exists()), None)
    if not checker_path:
        log.debug("[PRE-CHECK] ai_style_checker not found in sibling directories")
        return None

    try:
        result = subprocess.run(
            [sys.executable, str(checker_path / "cli.py"), str(qmd_path), "--format", "json"],
            capture_output=True, text=True, cwd=str(checker_path),
            env={**os.environ, "PYTHONUTF8": "1"}, timeout=30,
        )
        if result.returncode != 0:
            log.warning("[PRE-CHECK] checker exited with code %d: %s", result.returncode, result.stderr[:200])

        data = json.loads(result.stdout)
        score = data.get("ai_score", {}).get("score", 0)
        label = data.get("ai_score", {}).get("label", "")
        n_issues = sum(r.get("issue_count", 0) for r in data.get("results", []))

        log.info("[PRE-CHECK] %s: AI score %.1f/100 (%s), %d issues", paper_name, score, label, n_issues)

        report_path = PAPERS_DIR / paper_name / "_output" / "style_report.json"
        report_path.parent.mkdir(parents=True, exist_ok=True)
        report_path.write_text(json.dumps(data, indent=2), encoding="utf-8")

        if threshold > 0 and score > threshold:
            log.warning("[PRE-CHECK] BLOCKED: AI score %.1f exceeds threshold %.1f", score, threshold)
            return False

        return True
    except Exception as e:
        log.warning("[PRE-CHECK] ai_style_checker failed: %s", e)
        return None


def main():
    parser = argparse.ArgumentParser(description="Render .qmd -> styled DOCX (v2)")
    parser.add_argument("paper", nargs="?", help="Paper directory name")
    parser.add_argument("--all", action="store_true", help="Render all papers")
    parser.add_argument("--validate", action="store_true",
                        help="Run ai_style_checker before rendering")
    parser.add_argument("--threshold", type=float, default=0.0,
                        help="Block render if AI score exceeds this (requires --validate)")
    args = parser.parse_args()

    if args.all:
        for name in PAPER_NAMES:
            try:
                if args.validate:
                    check_result = _run_pre_render_check(name, args.threshold)
                    if check_result is False:
                        log.error("BLOCKED %s: exceeded threshold. Skipping.", name)
                        continue
                render_one(name, include_figures=True)
                render_one(name, include_figures=False)
            except Exception as e:
                log.error("FAILED %s: %s", name, e)
                import traceback
                traceback.print_exc()
    elif args.paper:
        if args.validate:
            check_result = _run_pre_render_check(args.paper, args.threshold)
            if check_result is False:
                log.error("Pre-render validation BLOCKED: AI score exceeded threshold. Aborting.")
                sys.exit(1)

        render_one(args.paper, include_figures=True)
        render_one(args.paper, include_figures=False)
        marks_path = (PAPERS_DIR / args.paper / "revision_marks.py")
        if marks_path.exists():
            import importlib.util
            spec = importlib.util.spec_from_file_location("marks", str(marks_path))
            mod = importlib.util.module_from_spec(spec)
            spec.loader.exec_module(mod)
            marks = getattr(mod, "REVISION_MARKS", {})
            render_one(args.paper, include_figures=True, revision_marks=marks)
    else:
        parser.print_help()


if __name__ == "__main__":
    main()
