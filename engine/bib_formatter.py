"""
BibTeX parser + citation resolver for the DOCX engine.

Reads a .bib file, provides:
  - format_citation(key) → "[Author, Year]" for inline text
  - format_reference(key) → full hanging-indent reference string
  - resolve_all_citations(text) → text with @citekey replaced by [Author, Year]
"""
from __future__ import annotations

import re
from pathlib import Path

from pybtex.database import parse_file as parse_bib_file


class BibFormatter:
    def __init__(self, bib_path: Path, style: str = "authoryear"):
        """style: 'authoryear' → (Author, Year) or 'numbered' → [1]"""
        from pybtex import errors as _pybtex_errors
        _pybtex_errors.set_strict_mode(False)
        self.db = parse_bib_file(str(bib_path))
        self.style = style
        self._cited: list[str] = []
        self._cited_set: set[str] = set()

    def _author_str(self, persons) -> str:
        if not persons:
            return ""
        names = []
        for p in persons:
            last = " ".join(p.last_names)
            firsts = " ".join(n[0] + "." for n in p.first_names if n)
            names.append(f"{last}, {firsts}".strip(", "))
        if len(names) <= 3:
            return ", ".join(names)
        return f"{names[0]} et al."

    def _short_author(self, persons) -> str:
        if not persons:
            return "?"
        first = persons[0]
        last = " ".join(first.last_names)
        if len(persons) > 2:
            return f"{last} et al."
        elif len(persons) == 2:
            second = " ".join(persons[1].last_names)
            return f"{last} and {second}"
        return last

    def format_citation(self, key: str) -> str:
        if key not in self._cited_set:
            self._cited_set.add(key)
            self._cited.append(key)
        entry = self.db.entries.get(key)
        if not entry:
            return f"[{key}]"
        if self.style == "numbered":
            idx = self._cited.index(key) + 1
            return f"[{idx}]"
        persons = entry.persons.get("author", [])
        year = entry.fields.get("year", "?")
        return f"({self._short_author(persons)}, {year})"

    def format_reference(self, key: str) -> str:
        entry = self.db.entries.get(key)
        if not entry:
            return f"[{key}] — entry not found in .bib"
        persons = entry.persons.get("author", [])
        fields = entry.fields
        year = fields.get("year", "?")
        title = fields.get("title", "").strip("{}")
        journal = fields.get("journal", fields.get("booktitle", "")).strip("{}")
        volume = fields.get("volume", "")
        pages = fields.get("pages", "").replace("--", "\u2013")
        doi = fields.get("doi", "")
        note = fields.get("note", "")

        # Clean LaTeX special chars
        def _clean_latex(s: str) -> str:
            s = s.replace(r"{\~n}", "\u00f1").replace(r"{\'e}", "\u00e9")
            s = s.replace(r"{\'a}", "\u00e1").replace(r"{\'i}", "\u00ed")
            s = s.replace(r"{\'o}", "\u00f3").replace(r"{\'u}", "\u00fa")
            s = s.replace(r"{\~a}", "\u00e3").replace(r"{\o}", "\u00f8")
            s = s.replace(r"{\~A}", "\u00c3").replace(r"{\O}", "\u00d8")
            s = s.replace(r"\textendash", "\u2013").replace("--", "\u2013")
            # Remove all remaining LaTeX braces: {API} -> API, {DNV GL} -> DNV GL
            s = re.sub(r'\{([^}]*)\}', r'\1', s)
            return s

        parts = [_clean_latex(self._author_str(persons))]
        parts.append(f"({year}).")
        parts.append(f"{_clean_latex(title)}.")
        if journal:
            j_str = _clean_latex(journal)
            if volume:
                j_str += f", {volume}"
            if pages:
                j_str += f", {pages}"
            parts.append(f"{j_str}.")
        if doi:
            parts.append(f"https://doi.org/{doi}")
        elif note:
            parts.append(_clean_latex(note))
        return " ".join(parts)

    def resolve_citations(self, text: str) -> str:
        def _replace_bracket(m):
            keys = [k.strip().lstrip("@") for k in m.group(1).split(";")]
            cites = [self.format_citation(k) for k in keys]
            return "; ".join(cites)

        def _replace_single(m):
            key = m.group(1)
            return self.format_citation(key)

        text = re.sub(r'\[(@[\w]+(?:;\s*@[\w]+)*)\]', _replace_bracket, text)
        text = re.sub(r'@([\w]+)', _replace_single, text)
        return text

    def get_cited_references(self) -> list[str]:
        refs = []
        for i, key in enumerate(self._cited):
            ref_text = self.format_reference(key)
            if self.style == "numbered":
                refs.append(f"[{i+1}] {ref_text}")
            else:
                refs.append(ref_text)
        return refs

    def get_all_references(self) -> list[str]:
        refs = []
        for key in sorted(self.db.entries.keys()):
            refs.append(self.format_reference(key))
        return refs
