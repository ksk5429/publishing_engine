"""
Generate supplementary material DOCX for a paper.

Reads supplementary content from manuscript.qmd YAML frontmatter
field `supplementary:` or from a separate supplementary.yaml.

Usage:
    python _shared/render_supplementary.py _shared/sample
"""
from __future__ import annotations
import argparse, logging, sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent))
from docx_engine import DocxBuilder
from qmd_parser import parse_qmd

PAPERS_DIR = Path(__file__).parent.parent
logging.basicConfig(level=logging.INFO, format="[%(asctime)s] %(message)s", datefmt="%H:%M:%S")
log = logging.getLogger("supplementary")

SAMPLE_SUPPLEMENTARY = {
    "tables": [
        {
            "caption": "Complete 64-feature ranking across four soil conditions with per-series relative margins, sign consistency, and rank score.",
            "headers": ["Rank", "Feature", "Sign consistent", "Min margin (%)", "Mean margin (%)", "Max margin (%)"],
            "rows": [
                ["1", "coh_ma_band", "Yes", "18.4", "96.3", "231.6"],
                ["2", "coh_ma_f1", "Yes", "11.7", "599.5", "1586.8"],
                ["3", "coh_mt_f1", "Yes", "9.7", "76.9", "262.1"],
                ["4", "coh_mt_band", "Yes", "8.5", "32.0", "92.6"],
                ["5", "coh_ba_band", "Yes", "4.0", "110.0", "304.6"],
                ["6", "rms_mid", "Yes", "2.0", "3.2", "5.0"],
                ["7", "coh_bt_band", "Yes", "1.3", "23.4", "79.7"],
                ["8", "rms_mid_f1", "Yes", "0.9", "2.9", "5.0"],
                ["9", "coh_ba_f1", "Yes", "0.5", "250.7", "881.5"],
                ["10", "coh_bt_f1", "Yes", "0.4", "4.7", "8.4"],
            ],
        },
        {
            "caption": "Per-fold LOSO classification results across three protocols and two feature sets.",
            "headers": ["Protocol", "Feature set", "Held out", "TN", "FP", "FN", "TP", "Accuracy"],
            "rows": [
                ["absolute_cos", "coh_ma_f1", "T2", "1", "0", "3", "0", "0.250"],
                ["absolute_cos", "coh_ma_f1", "T3", "1", "0", "3", "0", "0.250"],
                ["absolute_cos", "coh_ma_f1", "T4", "0", "1", "0", "3", "0.750"],
                ["absolute_cos", "coh_ma_f1", "T5", "1", "0", "3", "0", "0.250"],
                ["baselined_cos", "coh_ma_f1", "ALL", "4", "0", "9", "3", "0.438"],
                ["baselined_thresh", "top-5", "ALL", "4", "0", "9", "3", "0.438"],
            ],
        },
    ],
    "figures": [],
    "methods": [
        "The EWMA smoothing at span 48 corresponds to an 8-hour averaging window at 10-minute cadence. "
        "This span was selected in the companion V1 paper as the detection time-scale that balances sensitivity "
        "to persistent scour events (days) against rejection of short-lived environmental transients (minutes to hours).",
        "Bootstrap confidence intervals were computed using 10,000 iterations with numpy random seed 42. "
        "The Wilson score interval was used for the LOSO accuracy CI because it provides better coverage "
        "properties than the Wald interval at small sample sizes (n=16).",
    ],
}


def render_supplementary(paper_name: str) -> Path:
    paper_dir = PAPERS_DIR / paper_name
    meta, _ = parse_qmd(paper_dir / "manuscript.qmd")
    title = meta.get("title", paper_name)

    b = DocxBuilder(
        title=f"Supplementary Material\n{title[:80]}",
        authors=[
            {"name": "Kyeong-Sun Kim", "affiliation": "Seoul National University",
             "email": "kyeongsunkim@snu.ac.kr", "corresponding": True},
            {"name": "Sung-Ryul Kim", "affiliation": "Seoul National University"},
        ],
        running_header="Supplementary Material",
        abstract="This document contains supplementary tables, extended methods, "
                 "and additional data supporting the main manuscript.",
    )

    supp = meta.get("supplementary", SAMPLE_SUPPLEMENTARY)

    # Extended methods
    methods = supp.get("methods", [])
    if methods:
        b.heading("Supplementary Methods", level=1)
        for m in methods:
            b.paragraph(m)

    # Extended tables
    tables = supp.get("tables", [])
    for tbl_data in tables:
        b.page_break()
        b.table(tbl_data["headers"], tbl_data["rows"], caption_text=tbl_data.get("caption", ""))

    out_dir = paper_dir / "_output"
    out_dir.mkdir(parents=True, exist_ok=True)
    out_path = out_dir / "supplementary_material.docx"
    b.save(out_path)
    log.info("wrote %s (%d bytes)", out_path, out_path.stat().st_size)
    return out_path


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("paper")
    args = parser.parse_args()
    render_supplementary(args.paper)

if __name__ == "__main__":
    main()
