---
name: citation-check
description: Check every citation in a lecture outline, slide deck, chapter draft or manuscript - confirms each DOI is real and matches the reference as written (authors, year, title, journal, volume, pages), matches in-text citations to the reference list, and builds an Excel sheet that puts each cited claim next to the paper's abstract (and PDF passages, if the PDFs are available) so a person can judge whether the paper really says that. Use when the user asks to check, verify or audit citations or references, asks whether references are real or AI-made-up, or runs /citation-check with a file.
---

# Citation check

Checks the citations in one document and writes a report. **It never edits the document.** It
cannot decide whether a claim is a fair reading of a paper; it lines up the evidence so the user can.

Explain each step in plain language. Assume the user is a researcher, not a programmer.

## Step 1. Get the inputs (one question set)
1. **The document** (.docx, .pptx, .md or .txt). If they named a PDF, say the tool reads
   Word, PowerPoint, Markdown and text files only, and ask for the Word version.
2. **A folder of the cited PDFs** (optional). With it, the report shows the three passages in each
   paper that best match each claim, with page numbers. Without it, abstracts only.
3. **Slides with no reference list?** Ask which file holds the references (for `--refs`).
4. **R or Python?** Use whichever the user prefers. If they have no preference, check what is
   installed (`Rscript --version`, `python3 --version`) and use what is there.

## Step 2. Check the tools are installed
The scripts live next to this file: `python/check_citations.py` and `R/check_citations.R`.
- Python needs: `python3 -m pip install python-docx python-pptx pypdf openpyxl`
- R needs the packages listed at the top of `R/check_citations.R` (one `install.packages()` line).
Ask before installing anything, and show the command first.

## Step 3. Run it
From the folder that holds the document:
```
python3 <skill folder>/python/check_citations.py --doc "<document>" [--pdfs "<pdf folder>"] [--refs "<file with references>"]
Rscript <skill folder>/R/check_citations.R --doc "<document>" [--pdfs "<pdf folder>"] [--refs "<file with references>"]
```
- It looks up each DOI on Crossref, about 2-4 seconds per reference. Tell the user how long to
  expect (references x 3 s). Run long jobs in the background.
- Answers are saved in `citation_reports/cache/`, so a rerun is fast and `--no-network` works offline.
- Crossref answers faster when the request includes a contact email. **Never add the user's email
  yourself.** If they want to, they set it once: `export CITATION_CHECK_MAILTO=you@example.edu`.

## Step 4. Report back
The report goes to `citation_reports/` next to the document: an Excel workbook and a short
Markdown summary. Read the summary and tell the user, in this order:
1. **References that need attention**: MAJOR MISMATCH (the DOI points to a different paper) and
   NOT FOUND (the DOI does not exist) first, then MINOR MISMATCH with what differs.
2. **In-text citations with no reference**, and ambiguous ones (two possible papers).
3. **MANUAL CHECK**: references with no DOI (reports, some books). The tool can't check these.
4. What to do next: open the `Claims_to_verify` sheet and fill in the yellow columns
   (Interpretation correct? Y/N/Partly, Notes, Action) for each claim.

Never call a reference "fake" or "fabricated". Say what the tool found ("the DOI resolves to a
different paper: <title>") and let the user decide. A mismatch is often a typo.

## Limits to mention when relevant
- Abstracts often leave out the specific number being cited. If so, the user must read the paper.
- Numbered citation styles ([1], footnotes) are not parsed; author-year styles are.
- Scanned PDFs without a text layer can't be searched; they need OCR first.
- A capitalized word next to a year is occasionally picked up as a citation. These show as
  NO REFERENCE and are easy to spot.

## Try it first
`example/sample_lecture_notes.md` has two planted mistakes (a wrong volume number and a citation
with no reference). Running the check on it is a two-minute demo of what the report looks like.
