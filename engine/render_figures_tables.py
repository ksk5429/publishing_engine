"""
Generate a companion figures_and_tables.docx for each paper.

Each figure gets a full page with:
  - Figure number (bold, centered)
  - Image (full width)
  - Caption (italic)

Each table gets a full page with:
  - Table number (bold)
  - Booktabs-styled table
  - Caption (italic)

Usage:
    python _shared/render_figures_tables.py paperB_buckingham_pi
    python _shared/render_figures_tables.py --all
"""
from __future__ import annotations

import argparse
import logging
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
from docx_engine import DocxBuilder
from qmd_parser import parse_qmd

PAPERS_DIR = Path(__file__).parent.parent

logging.basicConfig(level=logging.INFO, format="[%(asctime)s] %(message)s", datefmt="%H:%M:%S")
log = logging.getLogger("fig_tables")


def render_figures_tables(paper_name: str) -> Path:
    paper_dir = PAPERS_DIR / paper_name
    qmd_path = paper_dir / "manuscript.qmd"
    meta, blocks = parse_qmd(qmd_path)

    title = meta.get("title", paper_name)
    b = DocxBuilder(
        title=f"Figures and Tables — {title[:60]}...",
        authors=[{"name": "Kyeong-Sun Kim", "affiliation": "Seoul National University",
                  "email": "kyeongsunkim@snu.ac.kr", "corresponding": True},
                 {"name": "Sung-Ryul Kim", "affiliation": "Seoul National University"}],
        running_header=f"Figures & Tables — {paper_name}",
        abstract="This companion document contains all figures and tables for the manuscript, "
                 "presented at full page width for review. Figure and table numbers correspond "
                 "to those cited in the main text.",
    )

    fig_count = 0
    tbl_count = 0

    for block in blocks:
        if block.type == "figure":
            fig_count += 1
            fig_path = paper_dir / block.path
            b.page_break()
            b.figure(fig_path, block.caption, width=6.3)

        elif block.type == "table" and block.headers:
            tbl_count += 1
            b.page_break()
            b.table(block.headers, block.rows, caption_text=block.caption)

    out_dir = paper_dir / "_output"
    out_dir.mkdir(parents=True, exist_ok=True)
    out_path = out_dir / "figures_and_tables.docx"
    b.save(out_path)
    log.info("%s: %d figures, %d tables → %s (%d bytes)",
             paper_name, fig_count, tbl_count, out_path.name, out_path.stat().st_size)
    return out_path


PAPER_NAMES = ["paperB_buckingham_pi", "paperV2_shm_sage", "paperOp3_aes"]


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("paper", nargs="?")
    parser.add_argument("--all", action="store_true")
    args = parser.parse_args()

    if args.all:
        for name in PAPER_NAMES:
            try:
                render_figures_tables(name)
            except Exception as e:
                log.error("FAILED %s: %s", name, e)
    elif args.paper:
        render_figures_tables(args.paper)
    else:
        parser.print_help()


if __name__ == "__main__":
    main()
