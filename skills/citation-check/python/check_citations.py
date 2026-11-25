#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
check_citations.py -- a citation checker for lecture outlines, slide decks,
chapter drafts and manuscripts.

WHAT IT DOES (plain language)
-----------------------------
1. Reads your document (.docx, .pptx, .md, or .txt), including table cells,
   slide text, and speaker notes.
2. Finds the reference list and looks up every DOI on Crossref to confirm that
   the paper is real and that the authors, year, title, journal, volume, and
   pages in your reference match the publisher's record.
3. Finds every in-text citation, e.g. "(Koch 2001)", "Nixon et al. (2009)",
   "(Lucas et al. 1999a, b)", and grabs the sentence it supports (the "claim").
4. Matches each in-text citation to a reference entry, and lists citations
   with no reference and references never cited.
5. Pulls each paper's abstract (Crossref, then OpenAlex, then Semantic Scholar)
   and, if you have the PDFs, the passages in the paper that best match each
   claim, with page numbers.
6. Writes an Excel workbook (plus a short Markdown summary) where each claim
   sits next to the evidence, with blank columns for YOU to record whether
   the citation is being interpreted correctly.

WHAT IT CANNOT DO
-----------------
It cannot judge whether a claim is a fair reading of the paper. It only puts
the evidence side by side so a human can decide quickly.

It never changes your document.

Run  `python3 check_citations.py --help`  for usage.

Part of research-skills (https://github.com/research-reclaimed/research-skills).
MIT license.
"""

import argparse
import datetime as dt
import difflib
import html
import json
import math
import os
import re
import subprocess
import sys
import time
import unicodedata
import urllib.parse
from collections import Counter, defaultdict
from pathlib import Path

# ---------------------------------------------------------------------------
# Basic settings
# ---------------------------------------------------------------------------

# Web APIs ask callers to identify themselves. Crossref's "polite pool" gives
# faster, more reliable answers to callers who include a contact email.
# This script never sends an email address unless YOU set one, by putting it in
# an environment variable named CITATION_CHECK_MAILTO before running, e.g.
#   export CITATION_CHECK_MAILTO=you@university.edu
_MAILTO = os.environ.get("CITATION_CHECK_MAILTO", "").strip()
USER_AGENT = ("research-skills-citation-check/1.0 "
              "(+https://github.com/research-reclaimed/research-skills"
              + (f"; mailto:{_MAILTO}" if _MAILTO else "") + ")")

# Minimum gap between two calls to the same website, in seconds.
# Crossref throttles bursts, so we stay slow and polite.
MIN_SECONDS_BETWEEN_CALLS = 1.1

# How long abstracts may be in the Claims sheet (the References sheet keeps
# the full text, up to Excel's cell limit).
ABSTRACT_CHARS_IN_CLAIMS = 1500
EXCEL_CELL_LIMIT = 32000


# ===========================================================================
# PART 1. Small text helpers
# ===========================================================================

def strip_accents(s):
    """'Möller' -> 'Moller', 'Norström' -> 'Norstrom'."""
    return "".join(c for c in unicodedata.normalize("NFKD", s)
                   if not unicodedata.combining(c))


def norm_name(s):
    """Normalize a surname for comparison: no accents, no spaces/punctuation,
    lower case. 'van der Heide' -> 'vanderheide', "D'Odorico" -> 'dodorico'."""
    s = strip_accents(s or "").lower()
    return re.sub(r"[^a-z]", "", s)


def norm_text(s):
    """Normalize a title or journal name: strip HTML/JATS tags, accents,
    punctuation, and case, and treat '&' as 'and'."""
    s = re.sub(r"<[^>]+>", " ", s or "")
    s = html.unescape(s)
    s = strip_accents(s).lower().replace("&", " and ")
    s = re.sub(r"[^a-z0-9]+", " ", s)
    return re.sub(r"\s+", " ", s).strip()


def clean_abstract(s):
    """Crossref abstracts come wrapped in JATS XML tags (<jats:p> etc.).
    Remove the tags and a leading 'Abstract' label."""
    if not s:
        return ""
    s = re.sub(r"<[^>]+>", " ", s)
    s = html.unescape(s)
    s = re.sub(r"\s+", " ", s).strip()
    s = re.sub(r"^(abstract|summary)[\s.:]*", "", s, flags=re.I)
    return s


def similarity(a, b):
    """0-1 similarity of two normalized strings (difflib ratio)."""
    a, b = norm_text(a), norm_text(b)
    if not a or not b:
        return 0.0
    return difflib.SequenceMatcher(None, a, b).ratio()


def fix_dashes(s):
    """Turn en/em dashes and similar into plain hyphens so '1–17' == '1-17'."""
    return re.sub(r"[‐-―−]", "-", s or "")


def truncate(s, n):
    s = s or ""
    return s if len(s) <= n else s[: n - 3].rstrip() + "..."


# ---------------------------------------------------------------------------
# Sentence splitting that does not break on "et al.", "e.g.", or initials.
# ---------------------------------------------------------------------------
_ABBREVIATIONS = ["et al.", "e.g.", "i.e.", "vs.", "cf.", "Fig.", "Figs.", "ca.",
                  "approx.", "pp.", "p.", "Eds.", "Ed.", "No.", "Vol.", "St.",
                  "Dr.", "U.S.", "Ch.", "sp.", "spp.", "etc.", "Proc.", "Natl.",
                  "Acad.", "Sci.", "Mar.", "Ecol.", "Prog.", "Ser.", "Jan.",
                  "Feb.", "Aug.", "Sept.", "Oct.", "Nov.", "Dec."]
_PROTECT = "⁣"   # an invisible character used as a temporary stand-in for "."


def split_sentences(text):
    """Split text into sentences, keeping abbreviations and initials intact."""
    t = text
    for ab in _ABBREVIATIONS:
        t = t.replace(ab, ab.replace(".", _PROTECT))
    # Single capital initials such as "J." in "Cloern, J.E."
    t = re.sub(r"\b([A-Z])\.(?=\s?[A-Z]?)", r"\1" + _PROTECT, t)
    # Decimal numbers such as "1.07"
    t = re.sub(r"(\d)\.(\d)", r"\1" + _PROTECT + r"\2", t)
    parts = re.split(r"(?<=[.!?])[\"”’)]?\s+(?=[\"“(•\[A-Z0-9])", t)
    return [p.replace(_PROTECT, ".").strip() for p in parts if p.strip()]


# ===========================================================================
# PART 2. Reading documents into "blocks" of text with a location label
# ===========================================================================
# Every input format is turned into a list of blocks. A block is a dict:
#   text      the paragraph / table cell / slide paragraph text
#   location  human-readable place, e.g. "2.5 Methods > table 7, row 'A2 ...'"
#   heading   True if the block is a heading
#   level     heading level (1, 2, 3...) or None
#   context   extra text shown to you (e.g. the whole table row)

REF_HEADING_RE = re.compile(
    r"^\s*(references|reference list|literature cited|bibliography|works cited|"
    r"sources cited|citations)\b", re.I)


def read_docx(path):
    """Read a Word file in document order: paragraphs AND tables.
    python-docx's doc.paragraphs skips tables, so we walk the XML body."""
    try:
        import docx
        from docx.oxml.ns import qn
        from docx.table import Table
        from docx.text.paragraph import Paragraph
    except ImportError:
        sys.exit("Reading .docx needs python-docx. Install it with:\n"
                 "  python3 -m pip install python-docx")

    d = docx.Document(str(path))
    blocks = []
    heading_stack = []          # list of (level, text)
    table_counter = [0]

    def where():
        return " > ".join(h[1] for h in heading_stack[-2:]) or "(top of document)"

    def handle_paragraph(p):
        text = p.text.strip()
        style = (p.style.name if p.style is not None else "") or ""
        m = re.match(r"(Heading|Title)\s*(\d*)", style)
        if m and text:
            level = int(m.group(2)) if m.group(2) else 0
            while heading_stack and heading_stack[-1][0] >= level:
                heading_stack.pop()
            heading_stack.append((level, text))
            blocks.append(dict(text=text, location=where(), heading=True,
                               level=level, context=text))
        elif text:
            blocks.append(dict(text=text, location=where(), heading=False,
                               level=None, context=text))

    def handle_table(tbl):
        table_counter[0] += 1
        tno = table_counter[0]
        rows = tbl.rows
        header = []
        for r_i, row in enumerate(rows):
            # Merged cells are returned repeatedly by python-docx; skip repeats.
            seen, cells = set(), []
            for c in row.cells:
                if id(c._tc) in seen:
                    continue
                seen.add(id(c._tc))
                cells.append(c)
            texts = [c.text.strip() for c in cells]
            if r_i == 0:
                header = texts
            label = texts[0] if texts else ""
            # Short first cells (e.g. "15-26" minutes) are more useful with
            # the second cell attached ("15-26 | A2 Producers push back").
            if len(label) <= 8 and len(texts) > 1:
                label = f"{label} | {texts[1]}"
            label = truncate(label.replace("\n", " / "), 80)
            row_text = truncate(" | ".join(t.replace("\n", " / ") for t in texts), 1500)
            for c_i, cell in enumerate(cells):
                col = header[c_i] if c_i < len(header) else f"col {c_i + 1}"
                col = truncate(col.replace("\n", " "), 40)
                loc = f"{where()} > table {tno}, row '{label}', column '{col}'"
                # Nested tables inside a cell
                for inner in cell.tables:
                    handle_table(inner)
                # Each line in a cell (bullets separated by line breaks) is
                # its own block so the claim text stays focused.
                for line in cell.text.split("\n"):
                    line = line.strip()
                    if line:
                        blocks.append(dict(text=line, location=loc, heading=False,
                                           level=None, context=row_text,
                                           header_row=(r_i == 0)))

    for el in d.element.body.iterchildren():
        if el.tag == qn("w:p"):
            handle_paragraph(Paragraph(el, d))
        elif el.tag == qn("w:tbl"):
            handle_table(Table(el, d))
    return blocks


def read_pptx(path):
    """Read a PowerPoint deck: slide text, tables, grouped shapes, speaker notes.
    Each slide title acts as a heading (level 1)."""
    try:
        from pptx import Presentation
    except ImportError:
        sys.exit("Reading .pptx needs python-pptx. Install it with:\n"
                 "  python3 -m pip install python-pptx")
    prs = Presentation(str(path))
    blocks = []

    def shape_texts(shape):
        """Yield (text, context) pairs from a shape, recursing into groups."""
        if shape.shape_type == 6 and hasattr(shape, "shapes"):   # group
            for s in shape.shapes:
                yield from shape_texts(s)
        if getattr(shape, "has_table", False) and shape.has_table:
            for row in shape.table.rows:
                cells = [c.text.strip() for c in row.cells]
                ctx = " | ".join(cells)
                for c in cells:
                    for line in c.split("\n"):
                        if line.strip():
                            yield line.strip(), ctx
        if getattr(shape, "has_text_frame", False) and shape.has_text_frame:
            for para in shape.text_frame.paragraphs:
                t = "".join(r.text for r in para.runs).strip()
                if t:
                    yield t, t

    for n, slide in enumerate(prs.slides, start=1):
        title = ""
        if slide.shapes.title is not None and slide.shapes.title.has_text_frame:
            title = slide.shapes.title.text_frame.text.strip()
        label = f"Slide {n}" + (f" ({truncate(title, 60)})" if title else "")
        blocks.append(dict(text=title or label, location=label, heading=True,
                           level=1, context=title))
        for shape in slide.shapes:
            if slide.shapes.title is not None and shape.shape_id == slide.shapes.title.shape_id:
                continue
            for t, ctx in shape_texts(shape):
                blocks.append(dict(text=t, location=label, heading=False,
                                   level=None, context=ctx))
        if slide.has_notes_slide:
            notes = slide.notes_slide.notes_text_frame.text
            for line in notes.split("\n"):
                if line.strip():
                    blocks.append(dict(text=line.strip(), location=label + " [speaker notes]",
                                       heading=False, level=None, context=line.strip()))
    return blocks


def read_text(path):
    """Read Markdown or plain text. Markdown '#' lines are headings; a plain
    line reading 'References' also counts as a heading. Markdown table rows
    ('| a | b |') become cells."""
    blocks = []
    heading_stack = []
    header = None
    for raw in Path(path).read_text(encoding="utf-8", errors="replace").splitlines():
        line = raw.strip()
        if not line:
            header = None
            continue
        m = re.match(r"^(#{1,6})\s+(.*)", line)
        if m or REF_HEADING_RE.match(line) and len(line) < 60:
            level = len(m.group(1)) if m else 1
            text = m.group(2).strip() if m else line
            while heading_stack and heading_stack[-1][0] >= level:
                heading_stack.pop()
            heading_stack.append((level, text))
            blocks.append(dict(text=text, location=" > ".join(h[1] for h in heading_stack[-2:]),
                               heading=True, level=level, context=text))
            continue
        where = " > ".join(h[1] for h in heading_stack[-2:]) or "(top of document)"
        if line.startswith("|"):
            cells = [c.strip() for c in line.strip("|").split("|")]
            if all(re.fullmatch(r":?-{2,}:?", c) for c in cells if c):
                continue
            if header is None:
                header = cells
            loc = f"{where} > table row '{truncate(cells[0], 60)}'"
            for c in cells:
                if c:
                    blocks.append(dict(text=c, location=loc, heading=False,
                                       level=None, context=" | ".join(cells)))
            continue
        line = re.sub(r"^([-*+]|\d+[.)])\s+", "", line)   # drop list markers
        blocks.append(dict(text=line, location=where, heading=False, level=None, context=line))
    return blocks


def read_document(path):
    ext = Path(path).suffix.lower()
    if ext == ".docx":
        return read_docx(path)
    if ext == ".pptx":
        return read_pptx(path)
    if ext in (".md", ".markdown", ".txt"):
        return read_text(path)
    sys.exit(f"Sorry, I can't read '{ext}' files. Use .docx, .pptx, .md, or .txt.")


# ===========================================================================
# PART 3. Parsing the reference list
# ===========================================================================

DOI_RE = re.compile(r"(?:https?://(?:dx\.)?doi\.org/|\bdoi:\s*)?(10\.\d{4,9}/[^\s\"<>]+)", re.I)
URL_RE = re.compile(r"https?://[^\s\"<>]+")
# "Surname, I. ... (1996)." at the start of a paragraph = looks like a reference.
LOOKS_LIKE_REF_RE = re.compile(
    r"^[^\W\d_][^:()]{0,300}(?:\(Eds?\.\)\s*)?\((?:1[5-9]|20)\d{2}[a-z]?[^)]*\)\.")
INITIALS_RE = re.compile(r"^(?:[A-Z][a-z]?\.\s*-?\s*)+$")


def clean_doi(doi):
    """Remove trailing punctuation that is part of the sentence, not the DOI."""
    doi = doi.strip().rstrip(".,;")
    # Drop a trailing ')' only if it is unbalanced (DOIs can contain parens).
    while doi.endswith(")") and doi.count("(") < doi.count(")"):
        doi = doi[:-1]
    while doi.endswith("]") and doi.count("[") < doi.count("]"):
        doi = doi[:-1]
    return doi


def parse_reference(text, index, location):
    """Split one reference string into its parts (best effort)."""
    ref = dict(index=index, text=text, location=location)
    m = DOI_RE.search(text)
    ref["doi"] = clean_doi(m.group(1)) if m else ""
    urls = [u.rstrip(".,;)") for u in URL_RE.findall(text)]
    ref["url"] = next((u for u in urls if "doi.org" not in u), "")

    # Year: the first "(1996)" / "(1999a)" / "(2003, September 3)".
    ym = re.search(r"\(((?:1[5-9]|20)\d{2})([a-z]?)[^)]*\)", text)
    ref["year"] = int(ym.group(1)) if ym else None
    ref["suffix"] = ym.group(2) if ym else ""

    # Authors: the text before "(Ed.)" or "(year".
    am = re.search(r"\((?:Eds?\.|(?:1[5-9]|20)\d{2})", text)
    author_seg = text[: am.start()].strip() if am else ""
    ref["et_al"] = bool(re.search(r"et\s+al", author_seg))
    seg = re.sub(r"et\s+al\.?", "", author_seg)
    seg = seg.replace("&", ",").replace(" and ", ",")
    tokens = [t.strip() for t in seg.split(",") if t.strip()]
    surnames = [t for t in tokens if not INITIALS_RE.match(t)]
    ref["surnames"] = surnames
    ref["corporate"] = len(surnames) == 1 and len(surnames[0].split()) >= 3 and not tokens[1:]
    # Keys an in-text citation may use for the first author.
    keys = set()
    if surnames:
        keys.add(norm_name(surnames[0]))
    br = re.search(r"\[([A-Za-z]{2,})\]", author_seg)          # "[EPA]"
    if br:
        keys.add(norm_name(br.group(1)))
    if ref["corporate"]:
        words = re.sub(r"\[.*?\]", "", surnames[0]).split()
        keys.add(norm_name("".join(w[0] for w in words if w[0].isupper())))
    ref["keys"] = keys

    # Title: text after "(year). " up to the next sentence end.
    rest = text[ym.end():].lstrip(" .") if ym else ""
    rest_nodoi = DOI_RE.sub("", URL_RE.sub("", rest)).strip()
    ref["rest"] = rest_nodoi
    tm = re.search(r"(?<!\b[A-Z])[.?!](?=\s+(?!I{1,3}:|IV:|V:)[A-Z(“\"])", rest_nodoi)
    ref["title"] = rest_nodoi[: tm.start() + (1 if tm and rest_nodoi[tm.start()] == "?" else 0)].strip() if tm else rest_nodoi
    ref["tail"] = rest_nodoi[tm.end():].strip() if tm else ""
    # Journal: the tail up to the first ", <number>" or " <number>(".
    jm = re.match(r"(.+?)(?:,\s*\d|\s+\d+\s*\(|,\s*e\w*\d|\.\s*$|$)", ref["tail"])
    ref["journal"] = jm.group(1).strip(" ,.") if jm else ""
    ref["short"] = short_label(ref)
    return ref


def short_label(ref):
    """'Nixon et al. 2009', 'Peterson & Fry 1987', 'EPA 2003'."""
    s = ref.get("surnames") or ["?"]
    first = s[0]
    br = re.search(r"\[([A-Za-z]{2,})\]", first)
    if br:
        first = br.group(1)
    if len(s) == 1 and not ref.get("et_al"):
        a = first
    elif len(s) == 2 and not ref.get("et_al"):
        a = f"{first} & {s[1]}"
    else:
        a = f"{first} et al."
    return f"{a} {ref.get('year') or 'n.d.'}{ref.get('suffix', '')}"


def extract_references(blocks, refs_blocks=None):
    """Find the reference list.
    - If a heading like 'References' exists, take the paragraphs under it
      (stopping at the next heading) that look like references.
    - Otherwise, take any paragraph that looks like a reference AND has a
      DOI or URL.
    Returns (references, set of block ids that belong to the reference section
    so in-text citation parsing can skip them)."""
    source = refs_blocks if refs_blocks is not None else blocks
    refs, skip_ids = [], set()
    in_refs, ref_level, collecting = False, None, False
    found_heading = any(b["heading"] and REF_HEADING_RE.match(b["text"]) for b in source)

    for i, b in enumerate(source):
        if b["heading"]:
            if REF_HEADING_RE.match(b["text"]):
                in_refs, ref_level, collecting = True, b["level"], True
                if refs_blocks is None:
                    skip_ids.add(i)
                continue
            if in_refs and b["level"] is not None and ref_level is not None and b["level"] <= ref_level:
                in_refs = False        # left the reference section
            elif in_refs:
                collecting = False     # a sub-heading such as "Verification notes"
            if in_refs and refs_blocks is None:
                skip_ids.add(i)
            continue
        if found_heading:
            if in_refs:
                if refs_blocks is None:
                    skip_ids.add(i)     # back matter: do not parse as claims
                if collecting and (LOOKS_LIKE_REF_RE.match(b["text"]) or DOI_RE.search(b["text"])):
                    refs.append(parse_reference(b["text"], len(refs) + 1, b["location"]))
        else:
            t = b["text"]
            if LOOKS_LIKE_REF_RE.match(t) and (DOI_RE.search(t) or URL_RE.search(t)):
                refs.append(parse_reference(t, len(refs) + 1, b["location"]))
                if refs_blocks is None:
                    skip_ids.add(i)
    return refs, skip_ids


# ===========================================================================
# PART 4. Web lookups with caching (curl via subprocess)
# ===========================================================================

class WebCache:
    """Stores every API answer in a JSON cache file, keyed by URL, so reruns
    are fast and --no-network works."""

    def __init__(self, path, offline=False, refresh=False):
        self.path = Path(path)
        self.offline = offline
        self.refresh = refresh
        self.data = {}
        self.last_call = {}
        self.new_calls = 0
        if self.path.exists():
            try:
                self.data = json.loads(self.path.read_text(encoding="utf-8"))
            except json.JSONDecodeError:
                print("  ! cache file was unreadable; starting a fresh cache")

    def save(self):
        self.path.parent.mkdir(parents=True, exist_ok=True)
        tmp = self.path.with_suffix(".tmp")
        tmp.write_text(json.dumps(self.data, ensure_ascii=False, indent=0), encoding="utf-8")
        tmp.replace(self.path)

    def get_json(self, url):
        """Return {'status': int, 'body': dict or None}.
        status 0 = network problem; status -1 = offline and not cached."""
        if url in self.data and not self.refresh:
            return self.data[url]
        if self.offline:
            return {"status": -1, "body": None}

        host = urllib.parse.urlparse(url).netloc
        result = {"status": 0, "body": None}
        for attempt in range(4):
            # Rate limit: wait so calls to the same host are >= ~1 s apart.
            wait = MIN_SECONDS_BETWEEN_CALLS - (time.time() - self.last_call.get(host, 0))
            if wait > 0:
                time.sleep(wait)
            self.last_call[host] = time.time()
            cmd = ["curl", "-sS", "-g", "-L", "--max-time", "40",
                   "-A", USER_AGENT, "-H", "Accept: application/json",
                   "-w", "\n__HTTP_STATUS__%{http_code}", url]
            try:
                proc = subprocess.run(cmd, capture_output=True, text=True, timeout=60)
                out = proc.stdout
            except subprocess.TimeoutExpired:
                out = ""
            m = re.search(r"\n__HTTP_STATUS__(\d{3})\s*$", out)
            code = int(m.group(1)) if m else 0
            body_text = out[: m.start()] if m else ""
            if code == 200:
                try:
                    result = {"status": 200, "body": json.loads(body_text)}
                except json.JSONDecodeError:
                    result = {"status": 200, "body": None}
                break
            if code in (404, 400, 410):
                result = {"status": code, "body": None}
                break
            # 429 (too many requests), 5xx, or network trouble: back off, retry.
            result = {"status": code, "body": None}
            time.sleep(3 * (2 ** attempt))
        self.new_calls += 1
        if result["status"] in (200, 404, 400, 410):
            self.data[url] = result
            self.save()
        return result


def quote_doi(doi):
    return urllib.parse.quote(doi, safe="/")


def crossref_lookup(web, doi):
    return web.get_json(f"https://api.crossref.org/works/{quote_doi(doi)}")


def openalex_lookup(web, doi):
    return web.get_json(f"https://api.openalex.org/works/doi:{quote_doi(doi)}")


def semanticscholar_lookup(web, doi):
    return web.get_json(
        f"https://api.semanticscholar.org/graph/v1/paper/DOI:{quote_doi(doi)}"
        "?fields=abstract,title,year,isOpenAccess,openAccessPdf")


def openalex_abstract(work):
    """OpenAlex stores abstracts as an 'inverted index' {word: [positions]}.
    Put the words back in order."""
    inv = (work or {}).get("abstract_inverted_index")
    if not inv:
        return ""
    pos = {}
    for word, places in inv.items():
        for p in places:
            pos[p] = word
    return " ".join(pos[k] for k in sorted(pos))


# ===========================================================================
# PART 5. Comparing a reference with its Crossref record
# ===========================================================================

def crossref_years(msg):
    out = {}
    for key, label in [("published-print", "print"), ("published-online", "online"),
                       ("issued", "issued")]:
        try:
            out[label] = int(msg[key]["date-parts"][0][0])
        except (KeyError, IndexError, TypeError, ValueError):
            pass
    return out


def acronym(s):
    """'Proceedings of the National Academy of Sciences' -> 'pnas'."""
    small = {"of", "the", "and", "in", "for", "on", "a", "an", "to"}
    return "".join(w[0] for w in norm_text(s).split() if w not in small)


def compare_reference(ref, msg):
    """Field-by-field comparison. Fills in ref['checks'] and ref['status']."""
    problems_major, problems_minor, notes = [], [], []
    ref_norm = norm_text(ref["text"])

    # --- First author ---
    authors = msg.get("author") or []
    cr_families = [a.get("family") or a.get("name") or "" for a in authors]
    cr_first = cr_families[0] if cr_families else ""
    ref_first = ref["surnames"][0] if ref["surnames"] else ""
    if not cr_first:
        editors = msg.get("editor") or []
        cr_first = (editors[0].get("family") or "") if editors else ""
    first_ok = bool(cr_first) and norm_name(ref_first) == norm_name(cr_first)
    if not first_ok and cr_first and norm_name(cr_first) and norm_name(cr_first) in norm_name(ref_first):
        first_ok = True
    ref["cr_first_author"] = cr_first
    ref["cr_author_count"] = len(authors)
    ref["check_author"] = "OK" if first_ok else ("no author in Crossref" if not cr_first else "MISMATCH")

    # Other listed authors: are they all present, in the same order?
    cr_norm = [norm_name(f) for f in cr_families]
    listed = [norm_name(s) for s in ref["surnames"]]
    missing = [s for s, n in zip(ref["surnames"][1:], listed[1:]) if n not in cr_norm]
    if cr_norm and len(listed) > 1 and not missing:
        if listed != cr_norm[: len(listed)]:
            problems_minor.append("author order differs from Crossref")
    if missing and cr_norm:
        problems_minor.append("author(s) not in Crossref list: " + ", ".join(missing))
    if cr_norm and not ref["et_al"] and len(listed) != len(cr_norm) and not ref["corporate"]:
        problems_minor.append(f"reference lists {len(listed)} authors, Crossref has {len(cr_norm)}")

    # --- Year (print vs online can differ by a year) ---
    years = crossref_years(msg)
    ref["cr_years"] = "; ".join(f"{k} {v}" for k, v in years.items()) or "none"
    yr = ref["year"]
    if not years:
        ref["check_year"] = "not in Crossref (unverified)"
    elif yr is None:
        ref["check_year"] = "no year in reference"
        problems_minor.append("year missing")
    elif yr in (years.get("print"), years.get("issued")):
        ref["check_year"] = "OK"
    elif years and any(abs(yr - v) <= 1 for v in years.values()):
        ref["check_year"] = f"within 1 year ({ref['cr_years']})"
        if yr != years.get("online"):
            problems_minor.append(f"year {yr} vs Crossref {ref['cr_years']}")
    else:
        ref["check_year"] = "MISMATCH"
        problems_minor.append(f"year {yr} vs Crossref {ref['cr_years']}")

    # --- Title ---
    cr_title = " ".join(msg.get("title") or [])
    cr_sub = " ".join(msg.get("subtitle") or [])
    cr_full = f"{cr_title}: {cr_sub}" if cr_sub else cr_title
    ref["cr_title"] = clean_abstract(cr_full)
    rest_norm = norm_text(ref["rest"])
    cands = [similarity(ref["title"], cr_full), similarity(ref["title"], cr_title)]
    for t in (cr_full, cr_title):
        tn = norm_text(t)
        if tn and tn in rest_norm:
            cands.append(1.0)
        elif tn:
            cands.append(difflib.SequenceMatcher(None, tn, rest_norm[: len(tn) + 5]).ratio())
    sim = max(cands) if cr_title else 0.0
    ref["title_similarity"] = round(sim, 2)
    if sim < 0.6:
        problems_major.append(f"title does not match (similarity {sim:.2f})")
    elif sim < 0.9:
        problems_minor.append(f"title differs somewhat (similarity {sim:.2f})")

    # --- Journal / book (container) ---
    containers = (msg.get("container-title") or []) + (msg.get("short-container-title") or [])
    ref["cr_journal"] = (msg.get("container-title") or [""])[0]
    ctype = msg.get("type", "")
    ref["cr_type"] = ctype
    if not containers:
        ref["check_journal"] = "n/a (no container in Crossref)" if ctype != "journal-article" else "not in Crossref"
    else:
        ok = False
        tail_norm = norm_text(ref["tail"]) or rest_norm
        for c in containers:
            cn = norm_text(c)
            if cn and (cn in ref_norm):
                ok = True
            if ref["journal"] and (norm_text(ref["journal"]) and norm_text(ref["journal"]) in cn):
                ok = True
            if ref["journal"] and similarity(ref["journal"], c) >= 0.85:
                ok = True
            # Fuzzy "is it somewhere in the reference?" check, which tolerates
            # typos in Crossref such as 'ECOLOGYAND' (compare a same-length window).
            if cn and len(cn) >= 8 and not ok:
                sm = difflib.SequenceMatcher(None, ref_norm, cn, autojunk=False)
                i, j, size = sm.find_longest_match(0, len(ref_norm), 0, len(cn))
                start = max(0, i - j)
                window = ref_norm[start: start + len(cn) + 3]
                if difflib.SequenceMatcher(None, window, cn).ratio() >= 0.85:
                    ok = True
            if acronym(c) and len(acronym(c)) >= 3 and re.search(rf"\b{acronym(c)}\b", tail_norm):
                ok = True
        ref["check_journal"] = "OK" if ok else "MISMATCH"
        if not ok:
            problems_minor.append(f"journal/book '{ref['journal']}' vs Crossref '{ref['cr_journal']}'")

    # Text after the title, without DOI/URL, is where volume and pages live.
    tail = fix_dashes(ref["tail"] or ref["rest"])

    # --- Volume ---
    vol = (msg.get("volume") or "").strip()
    ref["cr_volume"] = vol
    if not vol:
        ref["check_volume"] = "not in Crossref"
    elif ctype != "journal-article":
        ref["check_volume"] = "n/a (not a journal article)"
    elif re.search(rf"(?<![\w-]){re.escape(vol)}(?![\w-])", tail):
        ref["check_volume"] = "OK"
    else:
        ref["check_volume"] = "MISMATCH"
        problems_minor.append(f"volume: Crossref says {vol}")

    # --- Pages or article number ---
    page = fix_dashes((msg.get("page") or "").strip())
    artno = (msg.get("article-number") or "").strip()
    ref["cr_pages"] = page or (f"article {artno}" if artno else "")
    if page:
        first, _, last = page.partition("-")
        if last and len(last) < len(first) and last.isdigit() and first.isdigit():
            last = first[: len(first) - len(last)] + last      # "1367-81" -> "1367-1381"
        ref_range = re.search(rf"(?<![\w-]){re.escape(first)}\s*-\s*(\w+)", tail)
        if last and ref_range and ref_range.group(1) == last:
            ref["check_pages"] = "OK"
        elif last and ref_range:
            ref["check_pages"] = "MISMATCH"
            problems_minor.append(f"pages: reference has {first}-{ref_range.group(1)}, Crossref {first}-{last}")
        elif not last and re.search(rf"(?<![\w-]){re.escape(first)}(?![\w])", tail):
            ref["check_pages"] = "first page OK (Crossref gives only the first page; last page unverified)"
        elif last and ctype in ("book-chapter", "book") and first in tail:
            ref["check_pages"] = "OK"
        else:
            ref["check_pages"] = "MISMATCH"
            problems_minor.append(f"pages: Crossref says {page}")
    elif artno and re.search(rf"(?<![\w-]){re.escape(artno)}(?![\w-])", tail):
        ref["check_pages"] = "OK"
    elif artno and artno.lower() in ref["doi"].lower():
        # e.g. AGU papers: Crossref's "article number" is just the DOI suffix,
        # not the citation number printed on the paper (e.g. G03011).
        ref["check_pages"] = ("unverified: Crossref's article ID is the DOI suffix; "
                              "the paper's own citation number is not in Crossref")
    elif artno:
        if re.search(rf"(?<![\w-]){re.escape(artno)}(?![\w-])", tail):
            ref["check_pages"] = "OK"
        else:
            ref["check_pages"] = "MISMATCH"
            problems_minor.append(f"article number: Crossref says {artno}")
    else:
        ref["check_pages"] = "not in Crossref (unverified)"

    # --- Status ---
    if not first_ok and cr_first:
        if sim < 0.85:
            problems_major.append(f"first author {ref_first} vs Crossref {cr_first}")
        else:
            problems_minor.append(f"first author {ref_first} vs Crossref {cr_first}")
    if problems_major:
        ref["status"] = "MAJOR MISMATCH"
    elif problems_minor:
        ref["status"] = "MINOR MISMATCH"
    else:
        ref["status"] = "OK"
    ref["problems"] = "; ".join(problems_major + problems_minor + notes)


def check_references(refs, web, want_abstracts=True):
    """Look up every reference online (or in the cache) and compare."""
    for n, ref in enumerate(refs, start=1):
        ref.update(abstract="", abstract_source="", oa_pdf="", cr_title="",
                   cr_journal="", cr_years="", cr_volume="", cr_pages="",
                   cr_first_author="", cr_type="", title_similarity="",
                   check_author="", check_year="", check_journal="",
                   check_volume="", check_pages="", problems="")
        label = f"[{n}/{len(refs)}] {ref['short']}"
        if not ref["doi"]:
            ref["status"] = "MANUAL CHECK"
            ref["problems"] = ("no DOI: check by hand" +
                               (f" at {ref['url']}" if ref["url"] else " (no URL given)"))
            print(f"  {label}: no DOI -> manual check")
            continue
        cr = crossref_lookup(web, ref["doi"])
        if cr["status"] == -1:
            ref["status"] = "NOT CHECKED"
            ref["problems"] = "not in cache; rerun without --no-network"
            print(f"  {label}: not in cache")
            continue
        if cr["status"] != 200 or not cr["body"]:
            ref["status"] = "NOT FOUND" if cr["status"] in (404, 400, 410) else "NOT CHECKED"
            ref["problems"] = (f"DOI not found on Crossref (HTTP {cr['status']})"
                               if ref["status"] == "NOT FOUND"
                               else f"network problem (HTTP {cr['status']}); rerun later")
            print(f"  {label}: {ref['status']}")
            continue
        msg = cr["body"].get("message", {})
        compare_reference(ref, msg)
        print(f"  {label}: {ref['status']}" + (f"  ({ref['problems']})" if ref["problems"] else ""))

        if not want_abstracts:
            continue
        # Abstract: Crossref first ...
        ab = clean_abstract(msg.get("abstract", ""))
        if ab:
            ref["abstract"], ref["abstract_source"] = ab, "Crossref"
        # ... OpenAlex (also gives an open-access PDF link) ...
        oa = openalex_lookup(web, ref["doi"])
        work = oa.get("body") or {}
        if not ref["abstract"]:
            ab = openalex_abstract(work)
            if ab:
                ref["abstract"], ref["abstract_source"] = ab, "OpenAlex"
        loc = work.get("best_oa_location") or {}
        ref["oa_pdf"] = loc.get("pdf_url") or (work.get("open_access") or {}).get("oa_url") or ""
        # ... then Semantic Scholar.
        if not ref["abstract"] or not ref["oa_pdf"]:
            s2 = semanticscholar_lookup(web, ref["doi"])
            body = s2.get("body") or {}
            if not ref["abstract"] and body.get("abstract"):
                ref["abstract"], ref["abstract_source"] = body["abstract"].strip(), "Semantic Scholar"
            if not ref["oa_pdf"]:
                ref["oa_pdf"] = ((body.get("openAccessPdf") or {}).get("url")) or ""
        if not ref["abstract"]:
            ref["abstract_source"] = "none found (read the paper)"


# ===========================================================================
# PART 6. Finding in-text citations
# ===========================================================================

_UP = "A-ZÀ-ÖØ-Þ"                          # capital letters incl. accented
_PARTICLE = r"(?:(?:[Vv]an|[Vv]on|[Dd]e|[Dd]el|[Dd]er|[Dd]en|[Dd]a|[Dd]i|[Dd]u|[Ll]e|[Ll]a|[Tt]er|dos|das)\s+)"
SURNAME = rf"(?<![\w'’\-]){_PARTICLE}{{0,2}}[{_UP}](?:[^\W\d_]|['’\-](?=[^\W\d_]))+"
YEAR = r"(?:1[5-9]|20)\d{2}"
YEARLIST = rf"(?P<years>{YEAR}[a-z]?(?:\s*,\s*(?:{YEAR}[a-z]?|[a-z])(?!\w))*)"
AUTHORS = (rf"(?P<a1>{SURNAME})(?:\s+(?P<etal>et\s+al\.?)|(?:\s*&\s*|\s+and\s+)(?P<a2>{SURNAME}))?")
# "Nixon et al. (2009)"
NARRATIVE_RE = re.compile(AUTHORS + r"\s*\(\s*" + YEARLIST + r"(?:\s*,\s*pp?\.[^)]*)?\s*\)")
# "(Koch 2001)", "Koch 2001; Cloern 1996", "Lucas et al. 1999a, b"
PAREN_RE = re.compile(AUTHORS + r",?\s+" + YEARLIST + r"(?![\d–\-])")
# Full reference style used in reading lists: "Cloern, J.E. (1996)",
# "Smayda, T.J. & Borkman, D.G. (2008)"
_INITS = r"(?:[A-Z][a-z]?\.\s?-?\s?)+"
REFSTYLE_RE = re.compile(
    rf"(?P<rs>{SURNAME},\s*{_INITS}(?:,\s*{SURNAME},\s*{_INITS})*"
    rf"(?:,?\s*(?:&|and)\s*{SURNAME},\s*{_INITS})?(?:,?\s*et\s+al\.)?)\s*\((?P<years>{YEAR}[a-z]?)\)")

# Capitalized words that sit next to a year but are not authors.
STOPWORDS = {w.lower() for w in """
January February March April May June July August September October November December
Jan Feb Mar Apr Jun Jul Aug Sep Sept Oct Nov Dec Spring Summer Fall Autumn Winter
Since In By From Through Between And Or The Of To For On At Before After Until During
Circa Around About Year Years Fig Figure Table Tables Group Lecture Slide Slides Reading
Draft Version Revised Updated Accessed Copyright Est Est. Term Semester Week Class
Monday Tuesday Wednesday Thursday Friday Saturday Sunday Present Today Circa FY CY
""".split()}


def expand_years(years_text):
    """'1999a, b' -> [(1999,'a'), (1999,'b')]; '2017, 2022' -> [(2017,''),(2022,'')]."""
    out, last_year = [], None
    for part in re.split(r"\s*,\s*", years_text.strip()):
        m = re.fullmatch(rf"({YEAR})([a-z]?)", part)
        if m:
            last_year = int(m.group(1))
            out.append((last_year, m.group(2)))
        elif re.fullmatch(r"[a-z]", part) and last_year:
            out.append((last_year, part))
    return out


def find_citations(text):
    """Return a list of citation dicts found in one block of text."""
    found, taken = [], []

    def overlaps(s, e):
        return any(s < te and e > ts for ts, te in taken)

    # 1) Reference-style mentions first ("Cloern, J.E. (1996)").
    for m in REFSTYLE_RE.finditer(text):
        seg = re.sub(r"et\s+al\.?", "", m.group("rs")).replace("&", ",").replace(" and ", ",")
        names = [t.strip() for t in seg.split(",") if t.strip() and not INITIALS_RE.match(t.strip())]
        etal = bool(re.search(r"et\s+al", m.group("rs"))) or len(names) >= 3
        for y, suf in expand_years(m.group("years")):
            found.append(dict(raw=m.group(0), a1=names[0], a2=names[1] if len(names) == 2 else "",
                              et_al=etal, year=y, suffix=suf, start=m.start()))
        taken.append((m.start(), m.end()))

    # 2) Narrative "(Author year)" then parenthetical "Author year".
    for rx in (NARRATIVE_RE, PAREN_RE):
        for m in rx.finditer(text):
            if overlaps(m.start(), m.end()):
                continue
            a1 = m.group("a1")
            last_word = a1.split()[-1]
            if last_word.lower().strip(".") in STOPWORDS:
                continue
            if m.group("a2") and m.group("a2").split()[-1].lower() in STOPWORDS:
                continue
            for y, suf in expand_years(m.group("years")):
                found.append(dict(raw=m.group(0), a1=a1, a2=m.group("a2") or "",
                                  et_al=bool(m.group("etal")), year=y, suffix=suf,
                                  start=m.start()))
            taken.append((m.start(), m.end()))
    found.sort(key=lambda c: c["start"])
    return found


def cite_label(c):
    a = c["a1"] + (" et al." if c["et_al"] else (f" & {c['a2']}" if c["a2"] else ""))
    return f"{a} {c['year']}{c['suffix']}"


def collect_claims(blocks, skip_ids):
    """One row per in-text citation, with the sentence around it."""
    claims = []
    for i, b in enumerate(blocks):
        if i in skip_ids or b["heading"] or b.get("header_row"):
            continue
        cites = find_citations(b["text"])
        if not cites:
            continue
        sentences = split_sentences(b["text"])
        # Work out where each sentence starts and ends in the original text,
        # so a citation is paired with the sentence at ITS position (the same
        # citation can appear twice in one paragraph).
        spans, pos = [], 0
        for s in sentences:
            start = b["text"].find(s, pos)
            if start < 0:
                start = pos
            spans.append((start, start + len(s), s))
            pos = start + len(s)
        for c in cites:
            claim = next((s for st, en, s in spans if st <= c["start"] < en), b["text"])
            # A full reference written out in a reading list ("Cloern, J.E.
            # (1996). Title...") is best shown whole.
            if REFSTYLE_RE.match(c["raw"]) and len(claim) < 60:
                claim = b["text"]
            claims.append(dict(citation=cite_label(c), raw=c["raw"], cite=c,
                               location=b["location"], claim=claim,
                               context=b.get("context") or b["text"]))
    return claims


# ===========================================================================
# PART 7. Matching in-text citations to references
# ===========================================================================

def match_claims(claims, refs):
    """Surname + year, tolerant of a/b suffixes and +/-1 year. Uses 'et al.'
    and '&' to break ties (e.g. Nixon 1988 vs Nixon et al. 2009)."""
    by_key = defaultdict(list)
    for r in refs:
        for k in r["keys"]:
            by_key[k].append(r)

    for cl in claims:
        c = cl["cite"]
        cands = by_key.get(norm_name(c["a1"]), [])
        # A surname with a particle might be keyed without it (or vice versa).
        if not cands:
            bare = norm_name(re.sub(_PARTICLE, "", c["a1"]))
            cands = [r for r in refs if any(k.endswith(bare) for k in r["keys"])] if len(bare) > 2 else []
        exact = [r for r in cands if r["year"] == c["year"] and (not c["suffix"] or r["suffix"] == c["suffix"])]
        if not exact and c["suffix"]:
            exact = [r for r in cands if r["year"] == c["year"]]
        note = []
        pool = exact
        if not pool:
            pool = [r for r in cands if r["year"] and abs(r["year"] - c["year"]) == 1]
            if pool:
                note.append(f"year differs: in text {c['year']}, reference {pool[0]['year']}")

        def form_fits(r):
            n = len(r["surnames"])
            if c["a2"]:
                return n == 2 and not r["et_al"] and norm_name(r["surnames"][1]) == norm_name(c["a2"])
            if c["et_al"]:
                return n >= 3 or r["et_al"]
            return (n == 1 and not r["et_al"]) or r["corporate"]

        if len(pool) > 1:
            fitting = [r for r in pool if form_fits(r)]
            if len(fitting) >= 1:
                pool = fitting
        if len(pool) > 1:
            cl["ref"] = None
            cl["match_status"] = "AMBIGUOUS"
            cl["match_note"] = "could be: " + "; ".join(r["short"] for r in pool) + \
                               " (add a/b year letters or co-author names)"
            continue
        if not pool:
            cl["ref"] = None
            cl["match_status"] = "NO REFERENCE"
            others = sorted({r["short"] for r in cands})
            cl["match_note"] = ("no reference with this author and year" +
                                (f"; same author in list: {', '.join(others)}" if others else ""))
            continue
        r = pool[0]
        cl["ref"] = r
        if c["suffix"] != r["suffix"]:
            if r["suffix"] and not c["suffix"]:
                note.append(f"reference is {r['year']}{r['suffix']}: add the letter in text")
            elif c["suffix"] and not r["suffix"]:
                note.append(f"in-text letter '{c['suffix']}' but reference has no letter")
        if not r["corporate"] and not form_fits(r):
            n = len(r["surnames"])
            want = (r["surnames"][0] if n == 1 and not r["et_al"] else
                    f"{r['surnames'][0]} & {r['surnames'][1]}" if n == 2 and not r["et_al"] else
                    f"{r['surnames'][0]} et al.")
            note.append(f"citation form: reference suggests '{want} {r['year']}'")
        cl["match_status"] = "MATCHED" if not note else "MATCHED (check note)"
        cl["match_note"] = "; ".join(note)
        r.setdefault("cited_in", []).append(cl["location"])
    return claims


# ===========================================================================
# PART 8. PDFs: match to references, then find passages for each claim
# ===========================================================================

STOP = set("""a about above after again against all also am an and any are as at be because been
before being below between both but by can could did do does doing down during each few for from
further had has have having here how i if in into is it its itself just more most no nor not now of
off on once only or other our out over own same should so some such than that the their them then
there these they this those through to too under until up very was we were what when where which
while who whom why will with would you your et al e g i ie eg vs using use used can may might shows
show showed shown one two three via per within""".split())


def tokens(text):
    words = re.findall(r"[a-z][a-z0-9\-]{2,}", strip_accents(text.lower()))
    out = []
    for w in words:
        if w in STOP:
            continue
        if len(w) > 4 and w.endswith("s") and not w.endswith("ss"):
            w = w[:-1]                    # crude plural stripping: blooms -> bloom
        out.append(w)
    return out


def load_pdfs(pdf_dir):
    """Extract text from every PDF (recursively). Returns list of dicts."""
    try:
        from pypdf import PdfReader
    except ImportError:
        sys.exit("Searching PDFs needs pypdf. Install it with:\n"
                 "  python3 -m pip install pypdf\n"
                 "or run again with --no-pdfs.")
    pdfs = []
    for path in sorted(Path(pdf_dir).rglob("*.pdf")):
        try:
            reader = PdfReader(str(path))
            pages = []
            for p in reader.pages:
                try:
                    pages.append(p.extract_text() or "")
                except Exception:
                    pages.append("")
        except Exception as e:
            print(f"  ! could not read {path.name}: {e}")
            continue
        chars = sum(len(p) for p in pages)
        pdfs.append(dict(path=path, pages=pages, has_text=chars > 500))
        print(f"  read {path.name}: {len(pages)} pages" + ("" if chars > 500 else " (little/no text: scanned? needs OCR)"))
    return pdfs


def match_pdfs(pdfs, refs):
    """DOI on the first pages, then title similarity, then 'Author_Year' in
    the file name."""
    for pdf in pdfs:
        front = " ".join(pdf["pages"][:2])
        front_norm = norm_text(front)
        match, how = None, ""
        dois = {clean_doi(d).lower() for d in DOI_RE.findall(front)}
        for r in refs:
            if r["doi"] and r["doi"].lower() in dois:
                match, how = r, "DOI on first pages"
                break
        if not match:
            best, best_score = None, 0.0
            for r in refs:
                title = r.get("cr_title") or r["title"]
                tn = norm_text(title)
                if len(tn) < 15:
                    continue
                if tn in front_norm:
                    score = 1.0
                else:
                    sm = difflib.SequenceMatcher(None, front_norm[:4000], tn, autojunk=False)
                    blk = sm.find_longest_match(0, min(len(front_norm), 4000), 0, len(tn))
                    score = blk.size / len(tn)
                if score > best_score:
                    best, best_score = r, score
            if best and best_score >= 0.6:
                match, how = best, f"title on first pages ({best_score:.2f})"
        if not match:
            fname = norm_text(pdf["path"].stem)
            for r in refs:
                if r["surnames"] and r["year"] and norm_name(r["surnames"][0]) in fname.replace(" ", "") \
                        and str(r["year"]) in fname:
                    match, how = r, "author + year in file name"
                    break
        pdf["ref"], pdf["how"] = match, how
        if match:
            match.setdefault("pdfs", []).append(pdf)
            print(f"  {pdf['path'].name} -> {match['short']} ({how})")
        else:
            print(f"  {pdf['path'].name} -> no matching reference")


def build_windows(pdf):
    """Cut the paper into overlapping 3-sentence windows, remembering pages."""
    if "windows" in pdf:
        return pdf["windows"]
    sents = []
    for pno, text in enumerate(pdf["pages"], start=1):
        text = re.sub(r"-\s*\n\s*", "", text)          # re-join hyphenated words
        text = re.sub(r"\s+", " ", text)
        for s in split_sentences(text):
            if len(s) > 25:
                sents.append((pno, s))
    windows = []
    for i in range(len(sents)):
        chunk = sents[i:i + 3]
        pages = sorted({p for p, _ in chunk})
        txt = " ".join(s for _, s in chunk)
        windows.append(dict(i=i, pages=pages, text=txt, toks=Counter(tokens(txt))))
    # Inverse document frequency: rare words in this paper count more.
    df = Counter()
    for w in windows:
        df.update(set(w["toks"]))
    n = max(len(windows), 1)
    pdf["idf"] = {t: math.log((n + 1) / (c + 0.5)) for t, c in df.items()}
    pdf["windows"] = windows
    return windows


def best_passages(pdf, query, k=3):
    windows = build_windows(pdf)
    q = set(tokens(query))
    if not q or not windows:
        return []
    idf = pdf["idf"]
    scored = []
    for w in windows:
        shared = q & set(w["toks"])
        if not shared:
            continue
        score = sum(idf.get(t, 0) * (1 + math.log(w["toks"][t])) for t in shared)
        score /= math.sqrt(sum(w["toks"].values()) + 10)
        scored.append((score, w))
    scored.sort(key=lambda x: -x[0])
    chosen = []
    for score, w in scored:
        if all(abs(w["i"] - c["i"]) >= 3 for _, c in chosen):     # no overlapping windows
            chosen.append((score, w))
        if len(chosen) == k:
            break
    out = []
    for score, w in chosen:
        pg = (f"p. {w['pages'][0]}" if len(w["pages"]) == 1 else f"pp. {w['pages'][0]}-{w['pages'][-1]}")
        out.append(f"[{pg}] {truncate(w['text'], 900)}")
    return out


def attach_passages(claims):
    for cl in claims:
        cl["passages"], cl["pdf_name"] = [], ""
        r = cl.get("ref")
        if not r or not r.get("pdfs"):
            continue
        pdf = next((p for p in r["pdfs"] if p["has_text"]), r["pdfs"][0])
        cl["pdf_name"] = pdf["path"].name
        if not pdf["has_text"]:
            cl["passages"] = ["(PDF has no extractable text; it may be a scan that needs OCR)"]
            continue
        # Search with the claim minus the citations themselves.
        query = re.sub(r"\([^)]*\d{4}[^)]*\)", " ", cl["claim"])
        if len(tokens(query)) < 4:
            query = cl["context"]
        cl["passages"] = best_passages(pdf, query)


# ===========================================================================
# PART 9. Writing the Excel report and Markdown summary
# ===========================================================================

def write_excel(path, refs, claims, run_info):
    try:
        from openpyxl import Workbook
        from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
        from openpyxl.styles import Alignment, Font, PatternFill
        from openpyxl.utils import get_column_letter
        from openpyxl.worksheet.datavalidation import DataValidation
    except ImportError:
        sys.exit("Writing the Excel report needs openpyxl. Install it with:\n"
                 "  python3 -m pip install openpyxl")

    GREEN, AMBER, RED, GREY, BLUE = "C6EFCE", "FFEB9C", "FFC7CE", "E7E6E6", "DDEBF7"
    STATUS_COLOR = {"OK": GREEN, "MATCHED": GREEN, "MINOR MISMATCH": AMBER,
                    "MATCHED (check note)": AMBER, "MANUAL CHECK": AMBER, "AMBIGUOUS": AMBER,
                    "MAJOR MISMATCH": RED, "NOT FOUND": RED, "NO REFERENCE": RED,
                    "NOT CHECKED": GREY, "NOT CITED": BLUE}
    wrap = Alignment(wrap_text=True, vertical="top")

    def safe(v):
        if isinstance(v, str):
            return truncate(ILLEGAL_CHARACTERS_RE.sub("", v), EXCEL_CELL_LIMIT)
        return v

    def sheet(wb, title, headers, rows, widths, status_cols=()):
        ws = wb.create_sheet(title)
        ws.append(headers)
        for c in ws[1]:
            c.font = Font(bold=True, color="FFFFFF")
            c.fill = PatternFill("solid", fgColor="1F4E78")
            c.alignment = wrap
        for row in rows:
            ws.append([safe(v) for v in row])
        for col_idx, w in enumerate(widths, start=1):
            ws.column_dimensions[get_column_letter(col_idx)].width = w
        for row in ws.iter_rows(min_row=2):
            for c in row:
                c.alignment = wrap
            for sc in status_cols:
                cell = row[sc]
                color = STATUS_COLOR.get(str(cell.value))
                if color:
                    cell.fill = PatternFill("solid", fgColor=color)
        ws.freeze_panes = "B2" if title != "Summary" else "A2"
        ws.auto_filter.ref = ws.dimensions
        return ws

    wb = Workbook()
    wb.remove(wb.active)

    # (d) Summary first so it opens on top.
    sheet(wb, "Summary", ["Item", "Count / value"], run_info["summary_rows"], [55, 70])

    # (a) References
    ref_rows = []
    for r in refs:
        cites = len(r.get("cited_in", []))
        ref_rows.append([
            r["index"], r["status"], r["problems"], r["text"], r["doi"] or r["url"],
            r["surnames"][0] if r["surnames"] else "", r.get("cr_first_author", ""), r.get("check_author", ""),
            r["year"], r.get("cr_years", ""), r.get("check_year", ""),
            r.get("title_similarity", ""), r.get("cr_title", ""),
            r.get("journal", ""), r.get("cr_journal", ""), r.get("check_journal", ""),
            r.get("cr_volume", ""), r.get("check_volume", ""),
            r.get("cr_pages", ""), r.get("check_pages", ""), r.get("cr_type", ""),
            cites, r.get("oa_pdf", ""), "; ".join(p["path"].name for p in r.get("pdfs", [])),
            r.get("abstract_source", ""), r.get("abstract", "")])
    sheet(wb, "References",
          ["#", "Status", "Problems found", "Reference (as written)", "DOI / URL",
           "First author (ref)", "First author (Crossref)", "Author check",
           "Year (ref)", "Year (Crossref)", "Year check",
           "Title similarity (0-1)", "Title (Crossref)",
           "Journal/book (ref, parsed)", "Journal/book (Crossref)", "Journal check",
           "Volume (Crossref)", "Volume check", "Pages (Crossref)", "Pages check", "Crossref type",
           "Times cited in text", "Open-access PDF link", "Matched local PDF",
           "Abstract source", "Abstract"],
          ref_rows,
          [5, 16, 40, 60, 30, 14, 14, 10, 8, 22, 14, 10, 45, 25, 25, 10, 9, 10, 12, 18, 14, 8, 30, 25, 14, 80],
          status_cols=(1,))

    # (b) Claims to verify
    claim_rows = []
    for n, cl in enumerate(claims, start=1):
        r = cl.get("ref")
        ps = cl.get("passages", [])
        claim_rows.append([
            n, cl["citation"], cl["location"], cl["claim"], truncate(cl["context"], 1500),
            cl["match_status"], cl["match_note"],
            (f"{r['short']}: {r.get('cr_title') or r['title']}" if r else ""),
            r["status"] if r else "",
            truncate(r.get("abstract", ""), ABSTRACT_CHARS_IN_CLAIMS) if r else "",
            r.get("abstract_source", "") if r else "",
            r.get("oa_pdf", "") if r else "",
            cl.get("pdf_name", ""),
            ps[0] if len(ps) > 0 else "", ps[1] if len(ps) > 1 else "", ps[2] if len(ps) > 2 else "",
            "", "", ""])
    ws = sheet(wb, "Claims_to_verify",
               ["#", "Citation (as parsed)", "Location in document", "Claim (sentence citing it)",
                "Full paragraph / table row", "Match", "Match note", "Matched reference",
                "Reference status", "Abstract (truncated)", "Abstract source", "Open-access PDF link",
                "Local PDF", "Best PDF passage 1", "Best PDF passage 2", "Best PDF passage 3",
                "Interpretation correct? (Y/N/Partly)", "Notes", "Action"],
               claim_rows,
               [5, 22, 32, 55, 40, 14, 30, 40, 14, 70, 12, 25, 20, 50, 50, 50, 16, 35, 30],
               status_cols=(5, 8))
    # Highlight the reviewer's columns (yours) and offer a Y/N/Partly drop-down.
    for col in ("Q", "R", "S"):
        for c in ws[col][1:]:
            c.fill = PatternFill("solid", fgColor="FFF2CC")
    if len(claim_rows):
        dv = DataValidation(type="list", formula1='"Y,N,Partly"', allow_blank=True)
        ws.add_data_validation(dv)
        dv.add(f"Q2:Q{len(claim_rows) + 1}")

    # (c) Unmatched
    un_rows = []
    for cl in claims:
        if cl["match_status"] in ("NO REFERENCE", "AMBIGUOUS", "MATCHED (check note)"):
            un_rows.append([cl["match_status"], cl["citation"], cl["location"], cl["claim"], cl["match_note"]])
    for r in refs:
        if not r.get("cited_in"):
            un_rows.append(["NOT CITED", r["short"], "reference list",
                            r["text"], "listed but never cited in the text (may be fine: background/optional reading)"])
    sheet(wb, "Unmatched", ["Type", "Citation / reference", "Location", "Claim or reference text", "Note"],
          un_rows, [22, 28, 40, 70, 50], status_cols=(0,))

    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(str(path))


def write_markdown(path, refs, claims, run_info):
    L = []
    L.append(f"# Citation check: {run_info['doc_name']}")
    L.append("")
    L.append(f"Run {run_info['when']} | network: {run_info['network']} | PDFs: {run_info['pdf_note']}")
    L.append(f"Full report: `{run_info['xlsx_name']}`")
    L.append("")
    L.append("## Counts")
    L.append("")
    for k, v in run_info["summary_rows"]:
        if k and not str(k).startswith("---"):
            L.append(f"- {k}: {v}")
    L.append("")

    def section(title, rows):
        L.append(f"## {title}")
        L.append("")
        if not rows:
            L.append("None.")
        L.extend(rows)
        L.append("")

    section("References needing attention",
            [f"- **{r['short']}** ({r['status']}): {r['problems']}" for r in refs
             if r["status"] not in ("OK",)])
    section("In-text citations with no reference",
            [f"- {cl['citation']} ({cl['location']}): {cl['match_note']}" for cl in claims
             if cl["match_status"] == "NO REFERENCE"])
    section("Ambiguous or noted matches",
            [f"- {cl['citation']} ({cl['location']}): {cl['match_note']}" for cl in claims
             if cl["match_status"] in ("AMBIGUOUS", "MATCHED (check note)")])
    section("References never cited in text (may be fine)",
            [f"- {r['short']}" for r in refs if not r.get("cited_in")])
    L.append("Reminder: this tool checks that citations exist and are described correctly. "
             "Whether each claim fairly represents the paper is for a human to judge in the "
             "Claims_to_verify sheet.")
    path.write_text("\n".join(L) + "\n", encoding="utf-8")


# ===========================================================================
# PART 10. The command line
# ===========================================================================

HELP_TEXT = """\
Check the citations in a lecture outline, slide deck, chapter draft or manuscript.

It confirms that every reference is a real paper and is described correctly
(authors, year, title, journal, volume, pages, DOI), then lines up every
in-text citation with the paper's abstract (and, if you have the PDFs, the
best-matching passages) so you can confirm you are interpreting it correctly.

Output: an Excel workbook and a short Markdown summary, saved in a folder
called citation_reports/ next to your document. Nothing in your document is
changed.
"""

EPILOG = """\
examples:
  # Check a Word document, online:
  python3 check_citations.py --doc "My Chapter Draft.docx"

  # Also search a folder of PDFs for the passages behind each claim:
  python3 check_citations.py --doc lecture_outline.docx --pdfs readings/

  # Slides usually have no reference list: borrow the outline's list
  python3 check_citations.py --doc lecture_slides.pptx --refs lecture_outline.docx

  # Rerun quickly without internet, using saved answers from earlier runs:
  python3 check_citations.py --doc lecture_outline.docx --no-network

optional setting:
  Crossref answers faster if you include a contact email. To do that, set
  CITATION_CHECK_MAILTO before running, e.g.
    export CITATION_CHECK_MAILTO=you@university.edu
  Nothing is sent unless you set it.
"""


def main():
    ap = argparse.ArgumentParser(description=HELP_TEXT, epilog=EPILOG,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--doc", required=True,
                    help="the document to check (.docx, .pptx, .md, .txt)")
    ap.add_argument("--pdfs", default=None,
                    help="folder of PDFs to search (sub-folders included). Optional.")
    ap.add_argument("--no-pdfs", action="store_true", help="skip the PDF search entirely")
    ap.add_argument("--refs", default=None,
                    help="take the reference list from THIS file instead (useful for slide decks "
                         "that cite the outline's references)")
    ap.add_argument("--out", default=None,
                    help="folder for the report (default: citation_reports/ next to your document)")
    ap.add_argument("--no-network", action="store_true",
                    help="do not go online; reuse answers saved in the cache from earlier runs")
    ap.add_argument("--refresh", action="store_true",
                    help="ignore saved answers and look everything up again")
    args = ap.parse_args()

    doc = Path(args.doc).expanduser().resolve()
    if not doc.exists():
        sys.exit(f"Can't find the document: {doc}")
    print(f"Document: {doc}")

    out_dir = (Path(args.out).expanduser().resolve() if args.out
               else doc.parent / "citation_reports")
    cache_file = out_dir / "cache" / "api_cache.json"

    # 1. Read the document and find the reference list.
    blocks = read_document(doc)
    if args.refs:
        refs_path = Path(args.refs).expanduser().resolve()
        if not refs_path.exists():
            sys.exit(f"Can't find the --refs file: {refs_path}")
        refs, _ = extract_references(blocks, refs_blocks=read_document(refs_path))
        _, skip_ids = extract_references(blocks)
    else:
        refs, skip_ids = extract_references(blocks)
    print(f"Found {len(refs)} references.")
    if not refs:
        print("  ! No reference list found. Add a 'References' heading, or use --refs.")

    # 2. Look up references.
    web = WebCache(cache_file, offline=args.no_network, refresh=args.refresh)
    print("Checking references" + (" (offline, cache only)" if args.no_network else " (online; ~2-4 s each)") + "...")
    check_references(refs, web)

    # 3-4. In-text citations and matching.
    claims = collect_claims(blocks, skip_ids)
    match_claims(claims, refs)
    print(f"Found {len(claims)} in-text citations.")

    # 5. PDFs.
    pdf_dir = None if (args.no_pdfs or not args.pdfs) else Path(args.pdfs).expanduser().resolve()
    pdfs = []
    if pdf_dir and pdf_dir.exists():
        print(f"Reading PDFs in {pdf_dir} ...")
        pdfs = load_pdfs(pdf_dir)
        match_pdfs(pdfs, refs)
    elif args.pdfs and not args.no_pdfs:
        print(f"  ! PDF folder not found: {pdf_dir}")
    attach_passages(claims)

    # 6. Reports.
    today = dt.date.today().isoformat()
    stem = re.sub(r"[^\w\-]+", "_", doc.stem)
    xlsx = out_dir / f"citation_report_{stem}_{today}.xlsx"
    md = out_dir / f"citation_report_{stem}_{today}.md"
    # Never overwrite an existing report (it may hold your notes): add _run2, _run3...
    n = 2
    while xlsx.exists():
        xlsx = out_dir / f"citation_report_{stem}_{today}_run{n}.xlsx"
        md = out_dir / f"citation_report_{stem}_{today}_run{n}.md"
        n += 1

    st = Counter(r["status"] for r in refs)
    ms = Counter(cl["match_status"] for cl in claims)
    uncited = sum(1 for r in refs if not r.get("cited_in"))
    matched_pdfs = sum(1 for p in pdfs if p.get("ref"))
    summary_rows = [
        ["Document", str(doc)],
        ["Run date", today],
        ["Network", "offline (cache only)" if args.no_network else "online"],
        ["--- References ---", ""],
        ["References in list", len(refs)],
        ["OK", st.get("OK", 0)],
        ["MINOR MISMATCH", st.get("MINOR MISMATCH", 0)],
        ["MAJOR MISMATCH (probably the wrong paper)", st.get("MAJOR MISMATCH", 0)],
        ["NOT FOUND (DOI does not resolve)", st.get("NOT FOUND", 0)],
        ["MANUAL CHECK (no DOI)", st.get("MANUAL CHECK", 0)],
        ["NOT CHECKED (network problem or not cached)", st.get("NOT CHECKED", 0)],
        ["References with an abstract found", sum(1 for r in refs if r.get("abstract"))],
        ["References with an open-access PDF link", sum(1 for r in refs if r.get("oa_pdf"))],
        ["--- In-text citations ---", ""],
        ["In-text citations parsed (one per citation, per location)", len(claims)],
        ["Matched to a reference", ms.get("MATCHED", 0) + ms.get("MATCHED (check note)", 0)],
        ["  of which have a note (year off by 1, a/b letter, et al./& form)", ms.get("MATCHED (check note)", 0)],
        ["No matching reference", ms.get("NO REFERENCE", 0)],
        ["Ambiguous (more than one possible reference)", ms.get("AMBIGUOUS", 0)],
        ["References never cited in text", uncited],
        ["--- PDFs ---", ""],
        ["PDF folder", str(pdf_dir) if pdf_dir and pdf_dir.exists() else "none"],
        ["PDFs read", len(pdfs)],
        ["PDFs matched to a reference", matched_pdfs],
    ]
    run_info = dict(summary_rows=summary_rows, doc_name=doc.name, when=today,
                    network="offline" if args.no_network else "online",
                    pdf_note=f"{len(pdfs)} read, {matched_pdfs} matched",
                    xlsx_name=xlsx.name)
    write_excel(xlsx, refs, claims, run_info)
    write_markdown(md, refs, claims, run_info)
    if web.new_calls:
        web.save()

    print()
    for k, v in summary_rows:
        if not str(k).startswith("---") and k not in ("Document",):
            print(f"  {k}: {v}")
    print(f"\nReport:  {xlsx}\nSummary: {md}")


if __name__ == "__main__":
    main()
