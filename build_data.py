#!/usr/bin/env python3
"""
Build pipeline for the type-specimen bibliographic directory.

Takes:
  - a Zotero BibTeX export (publications, tagged with verbatim institution/country tags)
  - an xlsx lookup table (verbatim_tag -> canonical_city, canonical_country,
    canonical_institution, canonical_abbreviation)

Produces (in OUTPUT_DIR):
  - data/institutions.json   one record per institution or bare-country tag
  - data/publications.json   one record per publication, linked to institution IDs
  - reports/tag_gaps.csv          tags found in the library that don't resolve cleanly
  - reports/duplicate_abbreviations.csv   abbreviations mapped to >1 distinct institution

Nothing is silently dropped: every publication and every tag ends up in the
output somewhere, even if that "somewhere" is a gap report.

Re-run this any time the .bib or .xlsx changes; it fully regenerates all outputs.
"""

import csv
import json
import re
import unicodedata
from collections import defaultdict
from pathlib import Path

import bibtexparser
import openpyxl

# ---------------------------------------------------------------------------
# CONFIG — the only block you should need to edit
# ---------------------------------------------------------------------------
BIB_PATH = "/mnt/user-data/uploads/BirdTypePublications_20260908.bib"
XLSX_PATH = "/mnt/user-data/uploads/all_tags_canonical_20260914.xlsx"
XLSX_SHEET = "all_tags_canonical"

BUILD_ROOT = Path("/home/claude/build/output")
SITE_DIR = BUILD_ROOT / "site"  # public - this whole folder is what you push to GitHub Pages
REPORTS_DIR = BUILD_ROOT / "reports_internal"  # for you only - don't publish this folder

# Bracket markers in the `pages` field that are link/reference artifacts
# (stripped from display text). Anything else in brackets (e.g. "In Russian.")
# is treated as a genuine content note and kept.
LINK_MARKER_RE = re.compile(r"^(HYPERLINK|LINK\d*|BHL\d*|DOI|NHL)\b", re.IGNORECASE)
HYPERLINK_URL_RE = re.compile(r'HYPERLINK\s+"([^"]+)"')
BRACKET_RE = re.compile(r"\[([^\]]*)\]")
BRACE_RE = re.compile(r"[{}]")

# ---------------------------------------------------------------------------


def fix_mojibake(s):
    """Repair the UTF-8-decoded-as-Windows-1252 mangling present in the xlsx.
    Falls back to the 1:1 byte-value mapping for the handful of codepoints
    (0x81, 0x8D, 0x8F, 0x90, 0x9D) that cp1252 leaves undefined."""
    if not isinstance(s, str):
        return s
    if "Ã" not in s and "â€" not in s and "Â" not in s:
        return s
    barr = bytearray()
    for c in s:
        try:
            barr += c.encode("cp1252")
        except UnicodeEncodeError:
            if ord(c) < 256:
                barr.append(ord(c))
            else:
                return s  # unrecognised pattern - leave untouched
    try:
        return barr.decode("utf-8")
    except UnicodeDecodeError:
        return s


def slugify(s):
    s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode("ascii")
    s = re.sub(r"[^a-zA-Z0-9]+", "-", s).strip("-").lower()
    return s or "x"


def clean_url(url):
    """Strip LaTeX-style backslash-escapes (\\_, \\#, \\%, \\~ etc.) that
    Zotero's BibTeX export leaves in URLs, and add a missing https:// scheme
    where the source just recorded a bare domain/path."""
    if not url:
        return url
    url = re.sub(r"\\([_#%&~^])", r"\1", url).strip()
    if url and not re.match(r"^[a-zA-Z][a-zA-Z0-9+.-]*://", url):
        url = "https://" + url
    return url


def clean_braces(s):
    return BRACE_RE.sub("", s or "").strip()


def clean_pages(raw_pages):
    """Return (display_text, link, notes[]) from a raw `pages` field that may
    contain HYPERLINK/DOI/LINK/BHL artifacts and/or genuine language notes."""
    if not raw_pages:
        return "", None, []
    link = None
    notes = []

    def repl(m):
        nonlocal link
        content = m.group(1)
        hyper_url = HYPERLINK_URL_RE.search(content)
        if hyper_url:
            if link is None:
                link = hyper_url.group(1)
            return ""
        if LINK_MARKER_RE.match(content):
            return ""
        # genuine note (language, availability, etc.)
        notes.append(content.strip())
        return ""

    display = BRACKET_RE.sub(repl, raw_pages)
    # nested brackets (e.g. "[LINK1 [LINK2]]") can leave a stray bracket
    # behind after the inner pair is consumed - by this point every real
    # bracket pair has already been handled, so any leftover [ or ] is a
    # malformed link-marker artifact, not article content
    display = re.sub(r"[\[\]]", "", display)
    display = re.sub(r"\s{2,}", " ", display).strip(" .")
    return display, link, notes


def load_lookup(xlsx_path, sheet):
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    ws = wb[sheet]
    rows = list(ws.iter_rows(values_only=True))
    header = rows[0]
    lookup = {}
    raw_rows = []
    for r in rows[1:]:
        d = {k: fix_mojibake(v) for k, v in zip(header, r)}
        tag = (d.get("verbatim_tag") or "").strip()
        if not tag:
            continue
        d["verbatim_tag"] = tag
        raw_rows.append(d)
        lookup[tag.lower()] = d
    return lookup, raw_rows


def classify_lookup_row(d):
    """institution | country | unresolved"""
    if d.get("canonical_abbreviation"):
        return "institution"
    if d.get("canonical_country"):
        return "country"
    return "unresolved"


def resolve_raw_tag(raw_tag, lookup):
    """Try a direct match; then try splitting on '/' into sub-tags and
    requiring every sub-tag to resolve. Returns a list of lookup rows, or
    None if it can't be resolved at all."""
    key = raw_tag.strip().lower()
    if key in lookup:
        return [lookup[key]]
    parts = [p.strip() for p in raw_tag.split("/") if p.strip()]
    if len(parts) > 1:
        hits = []
        for p in parts:
            hit = lookup.get(p.lower())
            if hit is None:
                return None
            hits.append(hit)
        return hits
    return None


def institution_id(row):
    return slugify(f"{row['canonical_abbreviation']}-{row.get('canonical_city') or ''}")


def country_id(row):
    return "country-" + slugify(row["canonical_country"])


def main():
    (SITE_DIR / "data").mkdir(parents=True, exist_ok=True)
    (REPORTS_DIR).mkdir(parents=True, exist_ok=True)

    lookup, raw_rows = load_lookup(XLSX_PATH, XLSX_SHEET)

    with open(BIB_PATH, encoding="utf-8") as f:
        bib = bibtexparser.load(f)

    # --- institution / country directory, seeded from the lookup table ----
    institutions = {}
    for row in raw_rows:
        kind = classify_lookup_row(row)
        if kind == "institution":
            iid = institution_id(row)
            institutions.setdefault(
                iid,
                {
                    "id": iid,
                    "type": "institution",
                    "name": row["canonical_institution"],
                    "abbreviation": row["canonical_abbreviation"],
                    "city": row.get("canonical_city"),
                    "country": row.get("canonical_country"),
                    "verbatim_tags": set(),
                    "publications": set(),
                },
            )
            institutions[iid]["verbatim_tags"].add(row["verbatim_tag"])
        elif kind == "country":
            cid = country_id(row)
            institutions.setdefault(
                cid,
                {
                    "id": cid,
                    "type": "country",
                    "name": row["canonical_country"],
                    "abbreviation": None,
                    "city": None,
                    "country": row["canonical_country"],
                    "verbatim_tags": set(),
                    "publications": set(),
                },
            )
            institutions[cid]["verbatim_tags"].add(row["verbatim_tag"])

    # --- walk publications, resolving keyword tags -------------------------
    publications = []
    tag_gap_rows = []  # for reports/tag_gaps.csv
    ungrouped = []  # citekeys with zero institution/country link

    for e in bib.entries:
        citekey = e["ID"]
        raw_kw = e.get("keywords", "")
        raw_tags = [t.strip() for t in raw_kw.split(",") if t.strip()]

        inst_links = set()
        country_links = set()
        unresolved_tags = []

        for raw_tag in raw_tags:
            hits = resolve_raw_tag(raw_tag, lookup)
            if hits is None:
                unresolved_tags.append(raw_tag)
                tag_gap_rows.append(
                    {
                        "citekey": citekey,
                        "raw_tag": raw_tag,
                        "issue": "not found in lookup table",
                    }
                )
                continue
            for row in hits:
                kind = classify_lookup_row(row)
                if kind == "institution":
                    inst_links.add(institution_id(row))
                elif kind == "country":
                    country_links.add(country_id(row))
                else:
                    tag_gap_rows.append(
                        {
                            "citekey": citekey,
                            "raw_tag": raw_tag,
                            "issue": "present in lookup table but canonical fields are blank",
                        }
                    )

        # only attach country-level links for pubs with NO institution-level
        # link, so institution cards aren't duplicated onto a generic country
        # bucket
        linked_countries = set() if inst_links else country_links

        for iid in inst_links:
            institutions[iid]["publications"].add(citekey)
        for cid in linked_countries:
            institutions[cid]["publications"].add(citekey)

        if not inst_links and not linked_countries:
            ungrouped.append(
                {"citekey": citekey, "raw_keywords": raw_kw or "(no keywords field)"}
            )

        pages_display, pages_link, notes = clean_pages(e.get("pages", ""))
        link = clean_url(e.get("url")) or clean_url(pages_link)

        publications.append(
            {
                "id": citekey,
                "type": e["ENTRYTYPE"],
                "title": clean_braces(e.get("title", "")) or "[No title recorded]",
                "authors": e.get("author", ""),
                "year": e.get("year", ""),
                "venue": clean_braces(e.get("journal") or e.get("booktitle") or e.get("publisher") or ""),
                "volume": e.get("volume", ""),
                "number": e.get("number", ""),
                "pages": pages_display,
                "link": link,
                "notes": notes,
                "institutions": sorted(inst_links),
                "countries_only": sorted(linked_countries),
            }
        )

    # --- finalise institutions (sets -> sorted lists) -----------------------
    institutions_out = []
    for row in institutions.values():
        row = dict(row)
        row["verbatim_tags"] = sorted(row["verbatim_tags"])
        row["publications"] = sorted(row["publications"])
        row["publication_count"] = len(row["publications"])
        institutions_out.append(row)
    # drop country-only cards that never end up standing alone for any
    # publication - they're tags in the vocabulary but always co-occur with
    # an institution tag in this library, so a bare card would be a dead end
    institutions_out = [
        r for r in institutions_out if not (r["type"] == "country" and r["publication_count"] == 0)
    ]
    institutions_out.sort(key=lambda r: (r["type"] != "institution", r["name"] or ""))

    # --- duplicate-abbreviation report --------------------------------------
    by_abbr = defaultdict(set)
    for row in raw_rows:
        if row.get("canonical_abbreviation"):
            by_abbr[row["canonical_abbreviation"]].add(
                (row["canonical_institution"], row.get("canonical_city"), row.get("canonical_country"))
            )
    dup_report = [
        {"abbreviation": abbr, "variant": f"{name} | {city} | {country}"}
        for abbr, variants in by_abbr.items()
        if len(variants) > 1
        for (name, city, country) in variants
    ]

    # --- write outputs -------------------------------------------------------
    with open(SITE_DIR / "data" / "institutions.json", "w", encoding="utf-8") as f:
        json.dump(institutions_out, f, ensure_ascii=False, indent=1)

    with open(SITE_DIR / "data" / "publications.json", "w", encoding="utf-8") as f:
        json.dump(publications, f, ensure_ascii=False, indent=1)

    with open(REPORTS_DIR / "tag_gaps.csv", "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=["citekey", "raw_tag", "issue"])
        w.writeheader()
        w.writerows(tag_gap_rows)

    with open(REPORTS_DIR / "duplicate_abbreviations.csv", "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=["abbreviation", "variant"])
        w.writeheader()
        w.writerows(dup_report)

    with open(REPORTS_DIR / "ungrouped_publications.csv", "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=["citekey", "raw_keywords"])
        w.writeheader()
        w.writerows(ungrouped)

    # --- summary ---------------------------------------------------------
    n_inst = sum(1 for r in institutions_out if r["type"] == "institution")
    n_country = sum(1 for r in institutions_out if r["type"] == "country")
    print(f"Publications:            {len(publications)}")
    print(f"Institution cards:       {n_inst}")
    print(f"Country-only cards:      {n_country}")
    print(f"Tag gap rows:            {len(tag_gap_rows)}")
    print(f"Duplicate-abbrev rows:   {len(dup_report)}")
    print(f"Ungrouped publications:  {len(ungrouped)}")


if __name__ == "__main__":
    main()
