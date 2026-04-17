"""
AI-style text checker for manuscripts.

Scans a manuscript.qmd for common AI-generated writing patterns
and flags them for human revision.

Usage:
    python _shared/style_checker.py paperB_buckingham_pi
"""
from __future__ import annotations
import argparse, logging, re
from pathlib import Path

PAPERS_DIR = Path(__file__).parent.parent
logging.basicConfig(level=logging.INFO, format="[%(asctime)s] %(message)s", datefmt="%H:%M:%S")
log = logging.getLogger("style_check")

AI_PATTERNS = [
    (r"\b(Critically|Notably|Importantly|Remarkably|Interestingly)\b,", "AI transition word with comma"),
    (r"\b(Furthermore|Moreover|Additionally|Consequently)\b,", "AI transition word"),
    (r"\w+:\s+\w+", "Colon overuse (Noun: Explanation)"),
    (r"\s—\s", "Em-dash (convert to comma or parenthetical)"),
    (r"\b(pave the way|masquerade|harness|uniquely positioned|spearhead)\b", "Flowery AI verb"),
    (r"\b(is suggested to be|could potentially|may possibly)\b", "Hedging language"),
    (r"\b(revolutionary|groundbreaking|game-changing|cutting-edge|state-of-the-art)\b", "Promotional language"),
    (r"The ranking —.*— is", "AI-typical em-dash list pattern"),
    (r"\b(delve|leverage|foster|bolster|underscore|highlight)\b", "AI-typical verb"),
    (r"\b(In this context|In this regard|To this end|In light of)\b", "AI phrase"),
]

STRUCTURAL_CHECKS = [
    ("abstract_length", "Abstract should be 6-8 sentences, <200 words"),
    ("paragraph_length", "Paragraphs should be 4-6 sentences"),
    ("acronym_undefined", "Acronyms should be defined on first use"),
]


def check_style(paper_name: str) -> list[dict]:
    qmd = PAPERS_DIR / paper_name / "manuscript.qmd"
    text = qmd.read_text(encoding="utf-8", errors="replace")
    if text.startswith("---"):
        text = text[text.find("---", 3) + 3:]

    issues = []
    for line_num, line in enumerate(text.split("\n"), 1):
        for pattern, description in AI_PATTERNS:
            for m in re.finditer(pattern, line, re.IGNORECASE):
                issues.append({
                    "line": line_num,
                    "pattern": description,
                    "match": m.group()[:50],
                    "context": line.strip()[:80],
                })

    log.info("%s: %d AI-style issues found", paper_name, len(issues))
    for issue in issues[:10]:
        log.info("  L%d [%s]: '%s'", issue["line"], issue["pattern"], issue["match"])
    if len(issues) > 10:
        log.info("  ... and %d more", len(issues) - 10)

    return issues


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("paper", nargs="?")
    parser.add_argument("--all", action="store_true")
    args = parser.parse_args()

    papers = ["paperB_buckingham_pi", "paperA_dt_decision", "paperV1_jcshm",
              "paperV2_shm_sage", "paperOp3_aes"]
    targets = papers if args.all else ([args.paper] if args.paper else papers)

    total = 0
    for p in targets:
        issues = check_style(p)
        total += len(issues)
    log.info("Total AI-style issues across %d papers: %d", len(targets), total)


if __name__ == "__main__":
    main()
