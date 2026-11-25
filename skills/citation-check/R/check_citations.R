#!/usr/bin/env Rscript
# =============================================================================
# check_citations.R -- a citation checker for lecture outlines, slide decks,
# chapter drafts and manuscripts.
#
# This is the R version of python/check_citations.py. It is a line-for-line
# translation: same command line, same rules, same thresholds, same report.
# Pick whichever language you are more comfortable reading.
#
# WHAT IT DOES (plain language)
# -----------------------------
# 1. Reads your document (.docx, .pptx, .md, or .txt), including table cells,
#    slide text, and speaker notes.
# 2. Finds the reference list and looks up every DOI on Crossref to confirm
#    that the paper is real and that the authors, year, title, journal,
#    volume, and pages in your reference match the publisher's record.
# 3. Finds every in-text citation, e.g. "(Koch 2001)", "Nixon et al. (2009)",
#    "(Lucas et al. 1999a, b)", and grabs the sentence it supports (the
#    "claim").
# 4. Matches each in-text citation to a reference entry, and lists citations
#    with no reference and references never cited.
# 5. Pulls each paper's abstract (Crossref, then OpenAlex, then Semantic
#    Scholar) and, if you have the PDFs, the passages in the paper that best
#    match each claim, with page numbers.
# 6. Writes an Excel workbook (plus a short Markdown summary) where each claim
#    sits next to the evidence, with blank columns for YOU to record whether
#    the citation is being interpreted correctly.
#
# WHAT IT CANNOT DO
# -----------------
# It cannot judge whether a claim is a fair reading of the paper. It only puts
# the evidence side by side so a human can decide quickly.
# It never changes your document.
#
# HOW TO RUN IT
# -------------
# From a terminal (not from inside RStudio's console):
#   Rscript check_citations.R --doc "My Chapter Draft.docx"
#   Rscript check_citations.R --help        (shows every option)
#
# File paths come from the command line and are read relative to the folder
# you run the command from. That is why this script does NOT use
# here::here(): your document can live anywhere, not inside an R project.
#
# PACKAGES IT NEEDS
# -----------------
#   xml2       reads the XML inside .docx and .pptx files
#   httr2      makes the web calls to Crossref, OpenAlex, Semantic Scholar
#   jsonlite   reads and writes JSON (the web answers and the cache file)
#   openxlsx2  writes the Excel workbook
#   stringi    removes accents (e.g. "Moller" from "Moeller" with umlaut)
#   pdftools   reads PDFs; only needed if you use --pdfs
# Install them once with:
#   install.packages(c("xml2", "httr2", "jsonlite", "openxlsx2", "stringi", "pdftools"))
#
# A NOTE ON STYLE
# ---------------
# The data here are nested lists (a reference has a list of authors, a claim
# points at a reference), not tidy tables, so the code uses base R lists and
# small functions rather than dplyr. Each section below has the same PART
# number as the Python script so you can read the two side by side.
#
# Part of research-skills (https://github.com/research-reclaimed/research-skills).
# MIT license.
# =============================================================================


# -----------------------------------------------------------------------------
# Basic settings
# -----------------------------------------------------------------------------

# Web APIs ask callers to identify themselves. Crossref's "polite pool" gives
# faster, more reliable answers to callers who include a contact email.
# This script never sends an email address unless YOU set one, by putting it
# in an environment variable named CITATION_CHECK_MAILTO before running, e.g.
#   export CITATION_CHECK_MAILTO=you@university.edu
MAILTO <- trimws(Sys.getenv("CITATION_CHECK_MAILTO", ""))
USER_AGENT <- paste0(
  "research-skills-citation-check/1.0 ",
  "(+https://github.com/research-reclaimed/research-skills",
  if (nzchar(MAILTO)) paste0("; mailto:", MAILTO) else "",
  ")"
)

# Minimum gap between two calls to the same website, in seconds.
# Crossref throttles bursts, so we stay slow and polite.
MIN_SECONDS_BETWEEN_CALLS <- 1.1

# How long abstracts may be in the Claims sheet (the References sheet keeps
# the full text, up to Excel's cell limit).
ABSTRACT_CHARS_IN_CLAIMS <- 1500
EXCEL_CELL_LIMIT <- 32000


# =============================================================================
# PART 0. R helpers that copy Python's behavior exactly
# =============================================================================
# Python and R handle text slightly differently. These small helpers make R
# behave like Python so the two scripts give the same answers. You do not
# need to read them closely; each one says which Python feature it copies.

# Stop with a friendly message (like Python's sys.exit("message")).
stop_quietly <- function(msg, status = 1) {
  cat(msg, "\n", sep = "", file = stderr())
  quit(save = "no", status = status)
}

# Python's "truthiness": None, "", empty list, 0 and False all count as
# "nothing". py_or(a, b) returns a unless a is "nothing", then b
# (Python's `a or b`).
truthy <- function(x) {
  if (is.null(x) || length(x) == 0) return(FALSE)
  if (is.list(x)) return(TRUE)
  if (length(x) == 1) {
    if (is.na(x)) return(FALSE)
    if (is.character(x)) return(nzchar(x))
    if (is.logical(x)) return(isTRUE(x))
    if (is.numeric(x)) return(x != 0)
  }
  TRUE
}
py_or <- function(a, b) if (truthy(a)) a else b

# Read one field from a list, with a default when it is missing
# (Python's dict.get(key, default)).
get_field <- function(x, key, default = NULL) {
  if (is.null(x) || !is.list(x) || is.null(x[[key]])) default else x[[key]]
}

# --- Regular expressions -----------------------------------------------------
# Python's `re` module and R's PCRE engine (perl = TRUE) use almost the same
# pattern language. Two switches make PCRE behave like Python 3:
#   (*UCP)  makes \w, \d, \s and \b understand accented letters, as Python does.
#   (?i)    case-insensitive matching (Python's re.I).
rx <- function(pattern, ignore_case = FALSE) {
  paste0("(*UCP)", if (ignore_case) "(?i)" else "", pattern)
}

# Turn one regmatches result into a Python-like match object:
#   start (0-based, like Python), end, text, and g = the captured groups
#   (named when the pattern names them). A group that did not take part in
#   the match is NA (Python's None).
match_from <- function(x, m, k) {
  st <- m[k]
  len <- attr(m, "match.length")[k]
  out <- list(start = st - 1L, end = st - 1L + len,
              text = substr(x, st, st + len - 1L), g = character(0))
  cs <- attr(m, "capture.start")
  if (!is.null(cs) && ncol(cs) > 0) {
    cl <- attr(m, "capture.length")
    g <- vapply(seq_len(ncol(cs)), function(j) {
      s <- cs[k, j]
      l <- cl[k, j]
      if (s < 1 || l < 0) NA_character_ else substr(x, s, s + l - 1L)
    }, character(1))
    names(g) <- attr(m, "capture.names")
    out$g <- g
  }
  out
}

# re.search: first match anywhere, or NULL.
rx_search <- function(pattern, x, ignore_case = FALSE) {
  m <- regexpr(rx(pattern, ignore_case), x, perl = TRUE)
  if (m[1] == -1) return(NULL)
  match_from(x, m, 1)
}

# re.match: match only at the very start of the text.
rx_match_start <- function(pattern, x, ignore_case = FALSE) {
  rx_search(paste0("^(?:", pattern, ")"), x, ignore_case)
}

# re.fullmatch: the pattern must cover the whole text.
rx_full <- function(pattern, x, ignore_case = FALSE) {
  !is.null(rx_search(paste0("^(?:", pattern, ")\\z"), x, ignore_case))
}

# TRUE/FALSE: does the pattern occur anywhere? (bool(re.search(...)))
rx_detect <- function(pattern, x, ignore_case = FALSE) {
  grepl(rx(pattern, ignore_case), x, perl = TRUE)
}

# re.finditer: every non-overlapping match, left to right.
rx_find_all <- function(pattern, x, ignore_case = FALSE) {
  m <- gregexpr(rx(pattern, ignore_case), x, perl = TRUE)[[1]]
  if (m[1] == -1) return(list())
  lapply(seq_along(m), function(k) match_from(x, m, k))
}

# re.sub: replace every match. Back-references are written "\\1" in R.
rx_replace <- function(pattern, replacement, x, ignore_case = FALSE) {
  gsub(rx(pattern, ignore_case), replacement, x, perl = TRUE)
}

# re.split. (R's own strsplit() cannot see text before the current piece, so
# patterns that "look behind" would split in the wrong places. We split by
# hand from the match positions instead.)
rx_split <- function(pattern, x) {
  m <- gregexpr(rx(pattern), x, perl = TRUE)[[1]]
  if (m[1] == -1) return(x)
  starts <- as.integer(m)
  lens <- attr(m, "match.length")
  pieces <- character(0)
  prev <- 1L
  for (k in seq_along(starts)) {
    pieces <- c(pieces, substr(x, prev, starts[k] - 1L))
    prev <- starts[k] + lens[k]
  }
  c(pieces, substr(x, prev, nchar(x)))
}

# Wrap literal text so it can sit inside a pattern (Python's re.escape).
rx_literal <- function(s) {
  paste0("\\Q", gsub("\\E", "\\E\\\\E\\Q", s, fixed = TRUE), "\\E")
}

# --- Python string methods -----------------------------------------------------
# str.strip() with no argument: remove whitespace at both ends (Unicode-aware).
py_strip <- function(x) rx_replace("^\\s+|\\s+$", "", x)
py_rstrip <- function(x) rx_replace("\\s+$", "", x)
# str.strip(chars) / str.rstrip(chars) / str.lstrip(chars): remove any of the
# given characters from the ends.
strip_chars <- function(x, chars, left = TRUE, right = TRUE) {
  cls <- paste0("[", gsub("([\\]\\[\\\\^-])", "\\\\\\1", chars), "]+")
  if (left) x <- gsub(paste0("^", cls), "", x, perl = TRUE)
  if (right) x <- gsub(paste0(cls, "$"), "", x, perl = TRUE)
  x
}
# str.split(sep): unlike R's strsplit(), keeps empty pieces, including a
# trailing one ("a|b|" -> "a", "b", "").
py_split <- function(x, sep) strsplit(paste0(x, sep), sep, fixed = TRUE)[[1]]
# str.split() with no argument: split on runs of whitespace.
py_split_ws <- function(x) {
  out <- regmatches(x, gregexpr(rx("\\S+"), x, perl = TRUE))[[1]]
  if (length(out) == 0) character(0) else out
}
# Python slice s[start:end] with 0-based positions (end not included).
py_slice <- function(s, start, end = NULL) {
  n <- nchar(s)
  if (is.null(end)) end <- n
  start <- max(0, min(start, n))
  end <- max(0, min(end, n))
  if (end <= start) return("")
  substr(s, start + 1, end)
}
# str.find(sub, start): 0-based position of sub, or -1.
py_find <- function(s, sub, start = 0) {
  rest <- py_slice(s, start)
  hit <- regexpr(sub, rest, fixed = TRUE)
  if (hit[1] == -1) -1L else as.integer(start + hit[1] - 1L)
}
# The `in` test for strings: is `needle` somewhere inside `haystack`?
str_has <- function(haystack, needle) grepl(needle, haystack, fixed = TRUE)
starts_with <- function(x, prefix) startsWith(x, prefix)
ends_with <- function(x, suffix) endsWith(x, suffix)
# Count characters of one kind, e.g. how many "(" are in a DOI.
count_char <- function(x, ch) lengths(regmatches(x, gregexpr(ch, x, fixed = TRUE)))
# Python's sorted() of strings: plain code-point order, not your computer's
# alphabetical "locale" order.
py_sort <- function(x) sort(x, method = "radix")


# =============================================================================
# PART 0b. Python's difflib.SequenceMatcher, rewritten in R
# =============================================================================
# The thresholds in this script (0.6, 0.85, 0.9) were tuned on Python's
# difflib "ratio", so we copy that exact measure rather than use a different
# string distance. How it works:
#   1. Find the longest run of characters the two texts share.
#   2. Repeat on the pieces to the left and to the right of that run.
#   3. ratio = 2 * (characters matched) / (total characters in both texts).
# Identical texts score 1.0; texts with nothing in common score 0.0.
#
# "autojunk": when the second text is 200+ characters long, Python ignores
# characters that make up more than 1% of it (usually spaces) while searching
# for runs. We copy that too.

seq_matcher <- function(a, b, autojunk = TRUE) {
  a_codes <- utf8ToInt(enc2utf8(a))
  b_codes <- utf8ToInt(enc2utf8(b))
  if (length(a_codes) == 1 && is.na(a_codes)) a_codes <- integer(0)
  if (length(b_codes) == 1 && is.na(b_codes)) b_codes <- integer(0)
  alphabet <- unique(b_codes)
  # b2j: for each distinct character of b, the (0-based) places it occurs.
  b2j <- split(seq_along(b_codes) - 1L, factor(b_codes, levels = alphabet))
  n <- length(b_codes)
  if (autojunk && n >= 200) {
    ntest <- n %/% 100 + 1
    popular <- lengths(b2j) > ntest
    b2j[popular] <- list(integer(0))
  }
  list(a = a_codes, b = b_codes, b2j = unname(b2j),
       a_key = match(a_codes, alphabet))
}

# Longest block shared by a[alo:ahi] and b[blo:bhi]. Returns 0-based i, j and
# the block size, exactly as Python's find_longest_match does.
find_longest_match <- function(sm, alo, ahi, blo, bhi) {
  a <- sm$a
  b <- sm$b
  besti <- alo
  bestj <- blo
  bestsize <- 0L
  lb <- length(b)
  prev <- integer(lb)          # prev[j + 1] = run length ending at b[j]
  if (ahi > alo) {
    for (i in alo:(ahi - 1L)) {
      key <- sm$a_key[i + 1L]
      js <- if (is.na(key)) integer(0) else sm$b2j[[key]]
      js <- js[js >= blo & js < bhi]
      new <- integer(lb)
      if (length(js) > 0) {
        k <- c(0L, prev)[js + 1L] + 1L
        top <- max(k)
        if (top > bestsize) {
          w <- which.max(k)
          besti <- i - top + 1L
          bestj <- js[w] - top + 1L
          bestsize <- top
        }
        new[js + 1L] <- k
      }
      prev <- new
    }
  }
  # Python then extends the block over neighbouring equal characters that
  # the "autojunk" rule had hidden.
  while (besti > alo && bestj > blo && a[besti] == b[bestj]) {
    besti <- besti - 1L
    bestj <- bestj - 1L
    bestsize <- bestsize + 1L
  }
  while (besti + bestsize < ahi && bestj + bestsize < bhi &&
         a[besti + bestsize + 1L] == b[bestj + bestsize + 1L]) {
    bestsize <- bestsize + 1L
  }
  list(i = besti, j = bestj, size = bestsize)
}

# Python's SequenceMatcher(None, a, b).ratio().
seq_ratio <- function(a, b, autojunk = TRUE) {
  sm <- seq_matcher(a, b, autojunk)
  la <- length(sm$a)
  lb <- length(sm$b)
  if (la + lb == 0) return(1)
  matched <- 0L
  queue <- list(c(0L, la, 0L, lb))
  while (length(queue) > 0) {
    q <- queue[[length(queue)]]
    queue[[length(queue)]] <- NULL
    x <- find_longest_match(sm, q[1], q[2], q[3], q[4])
    if (x$size > 0) {
      matched <- matched + x$size
      if (q[1] < x$i && q[3] < x$j) queue[[length(queue) + 1]] <- c(q[1], x$i, q[3], x$j)
      if (x$i + x$size < q[2] && x$j + x$size < q[4]) {
        queue[[length(queue) + 1]] <- c(x$i + x$size, q[2], x$j + x$size, q[4])
      }
    }
  }
  2 * matched / (la + lb)
}


# =============================================================================
# PART 1. Small text helpers
# =============================================================================

# 'Moller' from 'Möller', 'Norstrom' from 'Norström': split each accented
# letter into letter + accent mark, then drop the accent marks.
strip_accents <- function(s) {
  s <- stringi::stri_trans_nfkd(s)
  stringi::stri_replace_all_regex(s, "[^\\p{ccc=0}]", "")
}

# Lower case the way Python does (full Unicode rules).
py_lower <- function(s) stringi::stri_trans_tolower(s, locale = "")

# Normalize a surname for comparison: no accents, no spaces/punctuation,
# lower case. 'van der Heide' -> 'vanderheide', "D'Odorico" -> 'dodorico'.
norm_name <- function(s) {
  if (is.null(s) || length(s) == 0 || is.na(s)) s <- ""
  s <- py_lower(strip_accents(s))
  gsub("[^a-z]", "", s, perl = TRUE)
}

# Turn HTML codes such as "&amp;" or "&#233;" back into characters
# (Python's html.unescape). Named codes are looked up with xml2's HTML parser.
.entity_cache <- new.env()
decode_entity <- function(ent) {
  body <- sub("^&", "", ent)
  if (starts_with(body, "#")) {
    num_txt <- sub(";$", "", sub("^#", "", body))
    num <- if (grepl("^[xX]", num_txt)) strtoi(sub("^[xX]", "", num_txt), 16L) else
      suppressWarnings(as.numeric(num_txt))
    if (is.na(num)) return(ent)
    if (num == 0 || (num >= 0xD800 && num <= 0xDFFF) || num > 0x10FFFF) return("\uFFFD")
    if (num == 0x0D) return("\r")
    if (num >= 0x80 && num <= 0x9F) {
      # Old Windows-1252 codes, as Python maps them.
      cp <- iconv(rawToChar(as.raw(num)), "CP1252", "UTF-8")
      if (!is.na(cp)) return(cp)
      return("")
    }
    if ((num >= 0x01 && num <= 0x08) || num == 0x0B || (num >= 0x0E && num <= 0x1F) ||
        num == 0x7F || (num >= 0xFDD0 && num <= 0xFDEF) || (num %% 0x10000) %in% c(0xFFFE, 0xFFFF)) {
      return("")
    }
    return(intToUtf8(num))
  }
  # Named code, e.g. "&amp;". Only the form with a closing ";" is decoded.
  if (!ends_with(body, ";")) return(ent)
  if (!is.null(.entity_cache[[body]])) return(.entity_cache[[body]])
  val <- tryCatch(
    xml2::xml_text(xml2::read_html(paste0("<p>&", body, "</p>"))),
    error = function(e) ent
  )
  if (!nzchar(val)) val <- ent
  assign(body, val, envir = .entity_cache)
  val
}
html_unescape <- function(s) {
  if (!str_has(s, "&")) return(s)
  pat <- "&(#[0-9]+;?|#[xX][0-9a-fA-F]+;?|[^\\t\\n\\f <&#;]{1,32};?)"
  m <- gregexpr(pat, s, perl = TRUE)
  ents <- regmatches(s, m)[[1]]
  if (length(ents) == 0) return(s)
  regmatches(s, m) <- list(vapply(ents, decode_entity, character(1), USE.NAMES = FALSE))
  s
}

# Normalize a title or journal name: strip HTML/JATS tags, accents,
# punctuation, and case, and treat '&' as 'and'.
norm_text <- function(s) {
  if (is.null(s) || length(s) == 0 || is.na(s)) s <- ""
  s <- rx_replace("<[^>]+>", " ", s)
  s <- html_unescape(s)
  s <- gsub("&", " and ", py_lower(strip_accents(s)), fixed = TRUE)
  s <- gsub("[^a-z0-9]+", " ", s, perl = TRUE)
  py_strip(rx_replace("\\s+", " ", s))
}

# Crossref abstracts come wrapped in JATS XML tags (<jats:p> etc.).
# Remove the tags and a leading 'Abstract' label.
clean_abstract <- function(s) {
  if (!truthy(s)) return("")
  s <- rx_replace("<[^>]+>", " ", s)
  s <- html_unescape(s)
  s <- py_strip(rx_replace("\\s+", " ", s))
  rx_replace("^(abstract|summary)[\\s.:]*", "", s, ignore_case = TRUE)
}

# 0-1 similarity of two normalized strings (difflib ratio).
similarity <- function(a, b) {
  a <- norm_text(a)
  b <- norm_text(b)
  if (!nzchar(a) || !nzchar(b)) return(0)
  seq_ratio(a, b)
}

# Turn en/em dashes and similar into plain hyphens so '1–17' == '1-17'.
fix_dashes <- function(s) {
  if (is.null(s) || is.na(s)) s <- ""
  gsub("[\u2010-\u2015\u2212]", "-", s, perl = TRUE)
}

# Shorten long text to n characters, ending with "...".
truncate_text <- function(s, n) {
  if (is.null(s) || length(s) == 0 || is.na(s)) s <- ""
  if (nchar(s) <= n) s else paste0(py_rstrip(substr(s, 1, n - 3)), "...")
}

# -----------------------------------------------------------------------------
# Sentence splitting that does not break on "et al.", "e.g.", or initials.
# -----------------------------------------------------------------------------
ABBREVIATIONS <- c("et al.", "e.g.", "i.e.", "vs.", "cf.", "Fig.", "Figs.", "ca.",
                   "approx.", "pp.", "p.", "Eds.", "Ed.", "No.", "Vol.", "St.",
                   "Dr.", "U.S.", "Ch.", "sp.", "spp.", "etc.", "Proc.", "Natl.",
                   "Acad.", "Sci.", "Mar.", "Ecol.", "Prog.", "Ser.", "Jan.",
                   "Feb.", "Aug.", "Sept.", "Oct.", "Nov.", "Dec.")
PROTECT <- "\u2063"   # an invisible character used as a temporary stand-in for "."

split_sentences <- function(text) {
  t <- text
  # Hide the full stops inside abbreviations so they do not end a sentence.
  for (ab in ABBREVIATIONS) {
    t <- gsub(ab, gsub(".", PROTECT, ab, fixed = TRUE), t, fixed = TRUE)
  }
  # Single capital initials such as "J." in "Cloern, J.E."
  t <- rx_replace("\\b([A-Z])\\.(?=\\s?[A-Z]?)", paste0("\\1", PROTECT), t)
  # Decimal numbers such as "1.07"
  t <- rx_replace("(\\d)\\.(\\d)", paste0("\\1", PROTECT, "\\2"), t)
  # Split after . ! or ? when the next word starts like a new sentence.
  parts <- rx_split("(?<=[.!?])[\"\u201D\u2019)]?\\s+(?=[\"\u201C(\u2022\\[A-Z0-9])", t)
  parts <- parts[nzchar(py_strip(parts))]
  py_strip(gsub(PROTECT, ".", parts, fixed = TRUE))
}


# =============================================================================
# PART 2. Reading documents into "blocks" of text with a location label
# =============================================================================
# Every input format is turned into a list of blocks. A block is a list:
#   text        the paragraph / table cell / slide paragraph text
#   location    human-readable place, e.g. "2.5 Methods > table 7, row 'A2 ...'"
#   heading     TRUE if the block is a heading
#   level       heading level (1, 2, 3...) or NA
#   context     extra text shown to you (e.g. the whole table row)
#   header_row  TRUE for the first row of a Word table

REF_HEADING_RE <- paste0("^\\s*(references|reference list|literature cited|bibliography|",
                         "works cited|sources cited|citations)\\b")
is_ref_heading <- function(text) rx_detect(REF_HEADING_RE, text, ignore_case = TRUE)

make_block <- function(text, location, heading = FALSE, level = NA_integer_,
                       context = text, header_row = FALSE) {
  list(text = text, location = location, heading = heading, level = level,
       context = context, header_row = header_row)
}

# "Section > Subsection" from the last two headings, as the Python does.
where_from <- function(heading_stack, empty = "(top of document)") {
  if (length(heading_stack) == 0) return(empty)
  texts <- vapply(heading_stack, function(h) h$text, character(1))
  paste(utils::tail(texts, 2), collapse = " > ")
}

# Push a heading onto the stack, first removing headings at the same or a
# deeper level (so "2.1" replaces "1.3" but sits under "2").
push_heading <- function(heading_stack, level, text) {
  while (length(heading_stack) > 0 && heading_stack[[length(heading_stack)]]$level >= level) {
    heading_stack[[length(heading_stack)]] <- NULL
  }
  heading_stack[[length(heading_stack) + 1]] <- list(level = level, text = text)
  heading_stack
}

# --- Reading files inside a .docx / .pptx ------------------------------------
# Word and PowerPoint files are zip folders full of XML files. These helpers
# read one XML file out of the zip without unpacking anything to disk.

zip_listing <- function(zip_path) utils::unzip(zip_path, list = TRUE)

read_zip_xml <- function(zip_path, member, listing = zip_listing(zip_path)) {
  hit <- listing[listing$Name == member, , drop = FALSE]
  if (nrow(hit) == 0) return(NULL)
  con <- unz(zip_path, member, open = "rb")
  on.exit(close(con))
  xml2::read_xml(readBin(con, "raw", n = hit$Length[1]))
}

# Tidy a path inside the zip: "ppt/slides/../notesSlides/x.xml" -> "ppt/notesSlides/x.xml".
zip_path_join <- function(base_dir, target) {
  if (starts_with(target, "/")) {
    parts <- strsplit(sub("^/", "", target), "/", fixed = TRUE)[[1]]
  } else {
    parts <- c(if (nzchar(base_dir)) strsplit(base_dir, "/", fixed = TRUE)[[1]],
               strsplit(target, "/", fixed = TRUE)[[1]])
  }
  out <- character(0)
  for (p in parts) {
    if (p == "" || p == ".") next
    if (p == "..") out <- utils::head(out, -1) else out <- c(out, p)
  }
  paste(out, collapse = "/")
}

# The "relationships" file says which XML file is the main document, which
# holds the styles, which are the slides, and so on.
read_rels <- function(zip_path, part, listing) {
  dir <- dirname(part)
  if (dir == ".") dir <- ""
  rels_name <- zip_path_join(dir, paste0("_rels/", basename(part), ".rels"))
  if (part == "") rels_name <- "_rels/.rels"
  x <- read_zip_xml(zip_path, rels_name, listing)
  empty <- data.frame(id = character(0), type = character(0), target = character(0))
  if (is.null(x)) return(empty)
  nodes <- xml2::xml_find_all(x, "//*[local-name()='Relationship']")
  if (length(nodes) == 0) return(empty)
  mode <- xml2::xml_attr(nodes, "TargetMode")
  keep <- is.na(mode) | mode != "External"
  nodes <- nodes[keep]
  data.frame(
    id = xml2::xml_attr(nodes, "Id"),
    type = xml2::xml_attr(nodes, "Type"),
    target = vapply(xml2::xml_attr(nodes, "Target"),
                    function(t) zip_path_join(dir, t), character(1), USE.NAMES = FALSE)
  )
}

rel_target_by_type <- function(rels, type_suffix) {
  hit <- rels$target[ends_with(rels$type, type_suffix)]
  if (length(hit) == 0) NA_character_ else hit[1]
}

# --- Word (.docx) -----------------------------------------------------------
NS_W <- c(w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main")

# Text of one Word paragraph, built the way python-docx builds it: the text of
# each run (and of runs inside hyperlinks), with tabs as "\t" and line breaks
# as "\n".
docx_para_text <- function(p) {
  kinds <- c("t", "tab", "br", "cr", "noBreakHyphen", "ptab")
  xp <- paste(c(paste0("./w:r/w:", kinds), paste0("./w:hyperlink/w:r/w:", kinds)),
              collapse = " | ")
  nodes <- xml2::xml_find_all(p, xp, NS_W)
  if (length(nodes) == 0) return("")
  pieces <- vapply(seq_along(nodes), function(k) {
    nd <- nodes[[k]]
    switch(xml2::xml_name(nd),
           t = xml2::xml_text(nd),
           tab = "\t",
           ptab = "\t",
           cr = "\n",
           noBreakHyphen = "-",
           br = {
             ty <- xml2::xml_attr(nd, "w:type", ns = NS_W)
             if (is.na(ty) || ty == "textWrapping") "\n" else ""
           },
           "")
  }, character(1))
  paste(pieces, collapse = "")
}

# Paragraph style names (e.g. "Heading 1"), looked up as python-docx does.
docx_style_lookup <- function(styles_xml) {
  if (is.null(styles_xml)) return(list(by_id = list(), default = ""))
  nodes <- xml2::xml_find_all(styles_xml, "./w:style", NS_W)
  ids <- xml2::xml_attr(nodes, "w:styleId", ns = NS_W)
  types <- xml2::xml_attr(nodes, "w:type", ns = NS_W)
  types[is.na(types)] <- "paragraph"
  defaults <- xml2::xml_attr(nodes, "w:default", ns = NS_W)
  name_nodes <- xml2::xml_find_first(nodes, "./w:name", NS_W)
  internal <- xml2::xml_attr(name_nodes, "w:val", ns = NS_W)
  # python-docx shows "heading 1" (the name stored in the file) as "Heading 1".
  ui <- ifelse(!is.na(internal) & grepl("^(caption|footer|header|heading [1-9])$", internal),
               sub("^(.)", "\\U\\1", internal, perl = TRUE), internal)
  ui[is.na(ui)] <- ""
  default_rows <- which(types == "paragraph" & !is.na(defaults) &
                          tolower(defaults) %in% c("1", "true", "on"))
  list(ids = ids, types = types, names = ui,
       default = if (length(default_rows)) ui[utils::tail(default_rows, 1)] else "")
}
docx_style_name <- function(p, lookup) {
  sid <- xml2::xml_attr(xml2::xml_find_first(p, "./w:pPr/w:pStyle", NS_W), "w:val", ns = NS_W)
  if (is.na(sid) || is.null(lookup$ids)) return(lookup$default)
  hit <- which(lookup$ids == sid)
  if (length(hit) == 0 || lookup$types[hit[1]] != "paragraph") return(lookup$default)
  lookup$names[hit[1]]
}

# Helpers for Word tables with merged cells.
tc_grid_span <- function(tc) {
  v <- xml2::xml_attr(xml2::xml_find_first(tc, "./w:tcPr/w:gridSpan", NS_W), "w:val", ns = NS_W)
  if (is.na(v)) 1L else as.integer(v)
}
tc_vmerge <- function(tc) {
  nd <- xml2::xml_find_first(tc, "./w:tcPr/w:vMerge", NS_W)
  if (inherits(nd, "xml_missing")) return(NA_character_)
  v <- xml2::xml_attr(nd, "w:val", ns = NS_W)
  if (is.na(v)) "continue" else v
}
tr_grid_before <- function(tr) {
  v <- xml2::xml_attr(xml2::xml_find_first(tr, "./w:trPr/w:gridBefore", NS_W), "w:val", ns = NS_W)
  if (is.na(v)) 0L else as.integer(v)
}
tc_grid_offset <- function(tc) {
  tr <- xml2::xml_parent(tc)
  before <- xml2::xml_find_all(tc, "./preceding-sibling::w:tc", NS_W)
  tr_grid_before(tr) + sum(vapply(before, tc_grid_span, integer(1)))
}
tc_at_grid_offset <- function(tr, offset) {
  remaining <- offset - tr_grid_before(tr)
  for (tc in xml2::xml_find_all(tr, "./w:tc", NS_W)) {
    if (remaining < 0) break
    if (remaining == 0) return(tc)
    remaining <- remaining - tc_grid_span(tc)
  }
  NULL
}
# A cell that continues a vertical merge shows the text of the cell above it
# (python-docx does this, so the text repeats down a merged column).
tc_root <- function(tc) {
  if (!identical(tc_vmerge(tc), "continue")) return(tc)
  tr_above <- xml2::xml_find_first(tc, "./ancestor::w:tr[1]/preceding-sibling::w:tr[1]", NS_W)
  if (inherits(tr_above, "xml_missing")) return(tc)
  above <- tc_at_grid_offset(tr_above, tc_grid_offset(tc))
  if (is.null(above)) return(tc)
  tc_root(above)
}

# Read a Word file in document order: paragraphs AND tables.
read_docx <- function(path) {
  listing <- zip_listing(path)
  main_part <- rel_target_by_type(read_rels(path, "", listing), "/officeDocument")
  if (is.na(main_part)) main_part <- "word/document.xml"
  doc <- read_zip_xml(path, main_part, listing)
  if (is.null(doc)) stop_quietly(paste0("Could not find the text inside the Word file: ", path))
  styles_part <- rel_target_by_type(read_rels(path, main_part, listing), "/styles")
  styles <- if (is.na(styles_part)) NULL else read_zip_xml(path, styles_part, listing)
  lookup <- docx_style_lookup(styles)

  # These three are shared by the helper functions below (hence <<-).
  blocks <- list()
  heading_stack <- list()
  table_counter <- 0L

  add_block <- function(b) blocks[[length(blocks) + 1]] <<- b
  where <- function() where_from(heading_stack)

  handle_paragraph <- function(p) {
    text <- py_strip(docx_para_text(p))
    style <- docx_style_name(p, lookup)
    m <- rx_match_start("(Heading|Title)\\s*(\\d*)", style)
    if (!is.null(m) && nzchar(text)) {
      level <- if (nzchar(m$g[2])) as.integer(m$g[2]) else 0L
      heading_stack <<- push_heading(heading_stack, level, text)
      add_block(make_block(text, where(), heading = TRUE, level = level, context = text))
    } else if (nzchar(text)) {
      add_block(make_block(text, where(), context = text))
    }
  }

  cell_text <- function(tc) {
    paras <- xml2::xml_find_all(tc, "./w:p", NS_W)
    paste(vapply(paras, docx_para_text, character(1)), collapse = "\n")
  }

  handle_table <- function(tbl) {
    table_counter <<- table_counter + 1L
    tno <- table_counter
    rows <- xml2::xml_find_all(tbl, "./w:tr", NS_W)
    header <- character(0)
    for (r_i in seq_along(rows)) {
      # Merged cells would otherwise appear more than once in a row; keep one.
      seen <- character(0)
      cells <- list()
      for (tc in xml2::xml_find_all(rows[[r_i]], "./w:tc", NS_W)) {
        root <- tc_root(tc)
        key <- xml2::xml_path(root)
        if (key %in% seen) next
        seen <- c(seen, key)
        cells[[length(cells) + 1]] <- root
      }
      raw_texts <- vapply(cells, cell_text, character(1))
      texts <- py_strip(raw_texts)
      if (r_i == 1) header <- texts
      label <- if (length(texts)) texts[1] else ""
      # Short first cells (e.g. "15-26" minutes) are more useful with the
      # second cell attached ("15-26 | A2 Producers push back").
      if (nchar(label) <= 8 && length(texts) > 1) label <- paste0(label, " | ", texts[2])
      label <- truncate_text(gsub("\n", " / ", label, fixed = TRUE), 80)
      row_text <- truncate_text(paste(gsub("\n", " / ", texts, fixed = TRUE), collapse = " | "), 1500)
      for (c_i in seq_along(cells)) {
        col <- if (c_i <= length(header)) header[c_i] else paste0("col ", c_i)
        col <- truncate_text(gsub("\n", " ", col, fixed = TRUE), 40)
        loc <- paste0(where(), " > table ", tno, ", row '", label, "', column '", col, "'")
        # Nested tables inside a cell come first, as in the Python.
        for (inner in xml2::xml_find_all(cells[[c_i]], "./w:tbl", NS_W)) handle_table(inner)
        # Each line in a cell (bullets separated by line breaks) is its own
        # block so the claim text stays focused.
        for (line in py_split(raw_texts[c_i], "\n")) {
          line <- py_strip(line)
          if (nzchar(line)) {
            add_block(make_block(line, loc, context = row_text, header_row = (r_i == 1)))
          }
        }
      }
    }
  }

  body <- xml2::xml_find_first(doc, "./w:body", NS_W)
  for (el in xml2::xml_find_all(body, "./w:p | ./w:tbl", NS_W)) {
    if (xml2::xml_name(el) == "p") handle_paragraph(el) else handle_table(el)
  }
  blocks
}

# --- PowerPoint (.pptx) ---------------------------------------------------------
NS_P <- c(p = "http://schemas.openxmlformats.org/presentationml/2006/main",
          a = "http://schemas.openxmlformats.org/drawingml/2006/main",
          r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships")
SHAPE_XPATH <- "./p:sp | ./p:grpSp | ./p:graphicFrame | ./p:cxnSp | ./p:pic | ./p:contentPart"

# Text of a whole text box: paragraphs joined by "\n"; a soft line break
# inside a paragraph is "\v" (as python-pptx does).
pptx_frame_text <- function(txBody) {
  if (inherits(txBody, "xml_missing")) return("")
  paras <- xml2::xml_find_all(txBody, "./a:p", NS_P)
  para_txt <- vapply(paras, function(p) {
    kids <- xml2::xml_find_all(p, "./a:r | ./a:br | ./a:fld", NS_P)
    paste(vapply(kids, function(k) {
      if (xml2::xml_name(k) == "br") return("\v")
      t <- xml2::xml_find_first(k, "./a:t", NS_P)
      if (inherits(t, "xml_missing")) "" else xml2::xml_text(t)
    }, character(1)), collapse = "")
  }, character(1))
  paste(para_txt, collapse = "\n")
}
shape_ph <- function(shape) xml2::xml_find_first(shape, "./*[1]/p:nvPr/p:ph", NS_P)
shape_id <- function(shape) xml2::xml_attr(xml2::xml_find_first(shape, "./*[1]/p:cNvPr", NS_P), "id")

# Yield (text, context) pairs from a shape, recursing into groups.
pptx_shape_texts <- function(shape) {
  out <- list()
  kind <- xml2::xml_name(shape)
  if (kind == "grpSp") {
    for (s in xml2::xml_find_all(shape, SHAPE_XPATH, NS_P)) out <- c(out, pptx_shape_texts(s))
  }
  tbl <- xml2::xml_find_first(shape, "./a:graphic/a:graphicData/a:tbl", NS_P)
  if (kind == "graphicFrame" && !inherits(tbl, "xml_missing")) {
    for (row in xml2::xml_find_all(tbl, "./a:tr", NS_P)) {
      cells <- vapply(xml2::xml_find_all(row, "./a:tc", NS_P), function(tc) {
        py_strip(pptx_frame_text(xml2::xml_find_first(tc, "./a:txBody", NS_P)))
      }, character(1))
      ctx <- paste(cells, collapse = " | ")
      for (c in cells) {
        for (line in py_split(c, "\n")) {
          if (nzchar(py_strip(line))) out[[length(out) + 1]] <- list(py_strip(line), ctx)
        }
      }
    }
  }
  txBody <- xml2::xml_find_first(shape, "./p:txBody", NS_P)
  if (kind == "sp" && !inherits(txBody, "xml_missing")) {
    for (para in xml2::xml_find_all(txBody, "./a:p", NS_P)) {
      # Only the text runs of each paragraph (fields and breaks are skipped),
      # matching the Python's "".join(r.text for r in para.runs).
      runs <- xml2::xml_find_all(para, "./a:r", NS_P)
      t <- py_strip(paste(vapply(runs, function(r) {
        tt <- xml2::xml_find_first(r, "./a:t", NS_P)
        if (inherits(tt, "xml_missing")) "" else xml2::xml_text(tt)
      }, character(1)), collapse = ""))
      if (nzchar(t)) out[[length(out) + 1]] <- list(t, t)
    }
  }
  out
}

# Read a PowerPoint deck: slide text, tables, grouped shapes, speaker notes.
# Each slide title acts as a heading (level 1).
read_pptx <- function(path) {
  listing <- zip_listing(path)
  pres_part <- rel_target_by_type(read_rels(path, "", listing), "/officeDocument")
  if (is.na(pres_part)) pres_part <- "ppt/presentation.xml"
  pres <- read_zip_xml(path, pres_part, listing)
  if (is.null(pres)) stop_quietly(paste0("Could not find the slides inside: ", path))
  pres_rels <- read_rels(path, pres_part, listing)
  slide_ids <- xml2::xml_attr(xml2::xml_find_all(pres, "./p:sldIdLst/p:sldId", NS_P),
                              "r:id", ns = NS_P)
  blocks <- list()
  for (n in seq_along(slide_ids)) {
    slide_part <- pres_rels$target[pres_rels$id == slide_ids[n]][1]
    slide <- read_zip_xml(path, slide_part, listing)
    if (is.null(slide)) next
    tree <- xml2::xml_find_first(slide, "./p:cSld/p:spTree", NS_P)
    shapes <- xml2::xml_find_all(tree, SHAPE_XPATH, NS_P)

    # The title is the first placeholder with index 0 (python-pptx's rule).
    title_shape <- NULL
    for (s in shapes) {
      ph <- shape_ph(s)
      if (!inherits(ph, "xml_missing")) {
        idx <- xml2::xml_attr(ph, "idx")
        if (is.na(idx) || idx == "0") { title_shape <- s; break }
      }
    }
    title <- ""
    if (!is.null(title_shape) && xml2::xml_name(title_shape) == "sp") {
      tb <- xml2::xml_find_first(title_shape, "./p:txBody", NS_P)
      if (!inherits(tb, "xml_missing")) title <- py_strip(pptx_frame_text(tb))
    }
    label <- paste0("Slide ", n, if (nzchar(title)) paste0(" (", truncate_text(title, 60), ")") else "")
    blocks[[length(blocks) + 1]] <- make_block(py_or(title, label), label, heading = TRUE,
                                               level = 1L, context = title)
    title_id <- if (is.null(title_shape)) NA_character_ else shape_id(title_shape)
    for (s in shapes) {
      if (!is.na(title_id) && identical(shape_id(s), title_id)) next
      for (tc in pptx_shape_texts(s)) {
        blocks[[length(blocks) + 1]] <- make_block(tc[[1]], label, context = tc[[2]])
      }
    }
    # Speaker notes live in a separate XML file linked from the slide.
    notes_part <- rel_target_by_type(read_rels(path, slide_part, listing), "/notesSlide")
    if (!is.na(notes_part)) {
      notes <- read_zip_xml(path, notes_part, listing)
      if (!is.null(notes)) {
        ntree <- xml2::xml_find_first(notes, "./p:cSld/p:spTree", NS_P)
        for (s in xml2::xml_find_all(ntree, SHAPE_XPATH, NS_P)) {
          ph <- shape_ph(s)
          if (!inherits(ph, "xml_missing") && identical(xml2::xml_attr(ph, "type"), "body")) {
            notes_text <- pptx_frame_text(xml2::xml_find_first(s, "./p:txBody", NS_P))
            for (line in py_split(notes_text, "\n")) {
              if (nzchar(py_strip(line))) {
                blocks[[length(blocks) + 1]] <- make_block(
                  py_strip(line), paste0(label, " [speaker notes]"), context = py_strip(line))
              }
            }
            break
          }
        }
      }
    }
  }
  blocks
}

# --- Markdown / plain text --------------------------------------------------
# Read the file as UTF-8 (bad bytes become a replacement character) and split
# it into lines the way Python's str.splitlines() does.
read_lines_like_python <- function(path) {
  raw <- readBin(path, "raw", n = file.info(path)$size)
  txt <- rawToChar(raw)
  Encoding(txt) <- "UTF-8"
  if (!validUTF8(txt)) txt <- iconv(txt, "UTF-8", "UTF-8", sub = "\uFFFD")
  breaks <- "\r\n|[\n\r\v\f\u001c\u001d\u001e\u0085\u2028\u2029]"
  if (!nzchar(txt)) return(character(0))
  lines <- rx_split(breaks, txt)
  # A final line break does not start a new (empty) line.
  if (length(lines) > 0 && rx_detect(paste0("(", breaks, ")$"), txt)) lines <- utils::head(lines, -1)
  lines
}

# Read Markdown or plain text. Markdown '#' lines are headings; a plain line
# reading 'References' also counts as a heading. Markdown table rows
# ('| a | b |') become cells.
read_text_file <- function(path) {
  blocks <- list()
  heading_stack <- list()
  header <- NULL
  for (raw in read_lines_like_python(path)) {
    line <- py_strip(raw)
    if (!nzchar(line)) {
      header <- NULL
      next
    }
    m <- rx_match_start("^(#{1,6})\\s+(.*)", line)
    if (!is.null(m) || (is_ref_heading(line) && nchar(line) < 60)) {
      level <- if (!is.null(m)) nchar(m$g[1]) else 1L
      text <- if (!is.null(m)) py_strip(m$g[2]) else line
      heading_stack <- push_heading(heading_stack, level, text)
      blocks[[length(blocks) + 1]] <- make_block(text, where_from(heading_stack, ""),
                                                 heading = TRUE, level = level, context = text)
      next
    }
    where <- where_from(heading_stack)
    if (starts_with(line, "|")) {
      cells <- py_strip(py_split(strip_chars(line, "|"), "|"))
      filled <- cells[nzchar(cells)]
      if (all(vapply(filled, function(c) rx_full(":?-{2,}:?", c), logical(1)))) next
      if (is.null(header)) header <- cells
      loc <- paste0(where, " > table row '", truncate_text(cells[1], 60), "'")
      for (c in cells) {
        if (nzchar(c)) {
          blocks[[length(blocks) + 1]] <- make_block(c, loc, context = paste(cells, collapse = " | "))
        }
      }
      next
    }
    line <- rx_replace("^([-*+]|\\d+[.)])\\s+", "", line)   # drop list markers
    blocks[[length(blocks) + 1]] <- make_block(line, where, context = line)
  }
  blocks
}

# Python's Path.suffix / Path.stem: the last ".xyz" of a file name.
path_suffix <- function(path) {
  name <- basename(path)
  i <- regexpr("\\.[^.]*$", name)
  if (i <= 1 || i == nchar(name)) "" else substring(name, i)
}
path_stem <- function(path) {
  name <- basename(path)
  suf <- path_suffix(path)
  if (nzchar(suf)) substr(name, 1, nchar(name) - nchar(suf)) else name
}

read_document <- function(path) {
  ext <- tolower(path_suffix(path))
  if (ext == ".docx") return(read_docx(path))
  if (ext == ".pptx") return(read_pptx(path))
  if (ext %in% c(".md", ".markdown", ".txt")) return(read_text_file(path))
  stop_quietly(paste0("Sorry, I can't read '", ext, "' files. Use .docx, .pptx, .md, or .txt."))
}


# =============================================================================
# PART 3. Parsing the reference list
# =============================================================================

DOI_RE <- "(?:https?://(?:dx\\.)?doi\\.org/|\\bdoi:\\s*)?(10\\.\\d{4,9}/[^\\s\"<>]+)"   # ignore case
URL_RE <- "https?://[^\\s\"<>]+"
# "Surname, I. ... (1996)." at the start of a paragraph = looks like a reference.
LOOKS_LIKE_REF_RE <- "^[^\\W\\d_][^:()]{0,300}(?:\\(Eds?\\.\\)\\s*)?\\((?:1[5-9]|20)\\d{2}[a-z]?[^)]*\\)\\."
INITIALS_RE <- "^(?:[A-Z][a-z]?\\.\\s*-?\\s*)+$"

has_doi <- function(text) rx_detect(DOI_RE, text, ignore_case = TRUE)
looks_like_ref <- function(text) !is.null(rx_match_start(LOOKS_LIKE_REF_RE, text))
is_initials <- function(t) !is.null(rx_match_start(INITIALS_RE, t))

# Remove trailing punctuation that is part of the sentence, not the DOI.
clean_doi <- function(doi) {
  doi <- strip_chars(py_strip(doi), ".,;", left = FALSE)
  # Drop a trailing ')' only if it is unbalanced (DOIs can contain parens).
  while (ends_with(doi, ")") && count_char(doi, "(") < count_char(doi, ")")) {
    doi <- substr(doi, 1, nchar(doi) - 1)
  }
  while (ends_with(doi, "]") && count_char(doi, "[") < count_char(doi, "]")) {
    doi <- substr(doi, 1, nchar(doi) - 1)
  }
  doi
}

# Split one reference string into its parts (best effort).
parse_reference <- function(text, index, location) {
  ref <- list(index = index, text = text, location = location)
  m <- rx_search(DOI_RE, text, ignore_case = TRUE)
  ref$doi <- if (!is.null(m)) clean_doi(m$g[1]) else ""
  urls <- vapply(rx_find_all(URL_RE, text), function(u) strip_chars(u$text, ".,;)", left = FALSE),
                 character(1))
  non_doi <- urls[!str_has(urls, "doi.org")]
  ref$url <- if (length(non_doi)) non_doi[1] else ""

  # Year: the first "(1996)" / "(1999a)" / "(2003, September 3)".
  ym <- rx_search("\\(((?:1[5-9]|20)\\d{2})([a-z]?)[^)]*\\)", text)
  ref$year <- if (!is.null(ym)) as.integer(ym$g[1]) else NA_integer_
  ref$suffix <- if (!is.null(ym)) ym$g[2] else ""

  # Authors: the text before "(Ed.)" or "(year".
  am <- rx_search("\\((?:Eds?\\.|(?:1[5-9]|20)\\d{2})", text)
  author_seg <- if (!is.null(am)) py_strip(py_slice(text, 0, am$start)) else ""
  ref$et_al <- rx_detect("et\\s+al", author_seg)
  seg <- rx_replace("et\\s+al\\.?", "", author_seg)
  seg <- gsub(" and ", ",", gsub("&", ",", seg, fixed = TRUE), fixed = TRUE)
  tokens_ <- py_strip(py_split(seg, ","))
  tokens_ <- tokens_[nzchar(tokens_)]
  surnames <- tokens_[!vapply(tokens_, is_initials, logical(1))]
  ref$surnames <- surnames
  ref$corporate <- length(surnames) == 1 && length(py_split_ws(surnames[1])) >= 3 &&
    length(tokens_) <= 1
  # Keys an in-text citation may use for the first author.
  keys <- character(0)
  if (length(surnames)) keys <- c(keys, norm_name(surnames[1]))
  br <- rx_search("\\[([A-Za-z]{2,})\\]", author_seg)          # "[EPA]"
  if (!is.null(br)) keys <- c(keys, norm_name(br$g[1]))
  if (ref$corporate) {
    words <- py_split_ws(rx_replace("\\[.*?\\]", "", surnames[1]))
    firsts <- substr(words, 1, 1)
    keys <- c(keys, norm_name(paste(firsts[grepl("^\\p{Lu}", firsts, perl = TRUE)], collapse = "")))
  }
  ref$keys <- unique(keys)

  # Title: text after "(year). " up to the next sentence end.
  rest <- if (!is.null(ym)) strip_chars(py_slice(text, ym$end), " .", right = FALSE) else ""
  rest_nodoi <- py_strip(rx_replace(DOI_RE, "", rx_replace(URL_RE, "", rest), ignore_case = TRUE))
  ref$rest <- rest_nodoi
  tm <- rx_search("(?<!\\b[A-Z])[.?!](?=\\s+(?!I{1,3}:|IV:|V:)[A-Z(\u201C\"])", rest_nodoi)
  if (!is.null(tm)) {
    keep_q <- if (substr(rest_nodoi, tm$start + 1, tm$start + 1) == "?") 1 else 0
    ref$title <- py_strip(py_slice(rest_nodoi, 0, tm$start + keep_q))
    ref$tail <- py_strip(py_slice(rest_nodoi, tm$end))
  } else {
    ref$title <- rest_nodoi
    ref$tail <- ""
  }
  # Journal: the tail up to the first ", <number>" or " <number>(".
  jm <- rx_match_start("(.+?)(?:,\\s*\\d|\\s+\\d+\\s*\\(|,\\s*e\\w*\\d|\\.\\s*$|$)", ref$tail)
  ref$journal <- if (!is.null(jm)) strip_chars(jm$g[1], " ,.") else ""
  ref$short <- short_label(ref)
  ref
}

# 'Nixon et al. 2009', 'Peterson & Fry 1987', 'EPA 2003'.
short_label <- function(ref) {
  s <- if (length(ref$surnames)) ref$surnames else "?"
  first <- s[1]
  br <- rx_search("\\[([A-Za-z]{2,})\\]", first)
  if (!is.null(br)) first <- br$g[1]
  et_al <- isTRUE(ref$et_al)
  a <- if (length(s) == 1 && !et_al) first else
    if (length(s) == 2 && !et_al) paste0(first, " & ", s[2]) else paste0(first, " et al.")
  year <- if (is.na(ref$year) || ref$year == 0) "n.d." else as.character(ref$year)
  paste0(a, " ", year, ref$suffix)
}

# Find the reference list.
# - If a heading like 'References' exists, take the paragraphs under it
#   (stopping at the next heading) that look like references.
# - Otherwise, take any paragraph that looks like a reference AND has a DOI
#   or URL.
# Returns refs plus skip_ids: the block numbers that belong to the reference
# section, so in-text citation parsing can skip them.
extract_references <- function(blocks, refs_blocks = NULL) {
  source <- if (!is.null(refs_blocks)) refs_blocks else blocks
  own_list <- is.null(refs_blocks)
  refs <- list()
  skip_ids <- integer(0)
  in_refs <- FALSE
  ref_level <- NA_integer_
  collecting <- FALSE
  found_heading <- any(vapply(source, function(b) b$heading && is_ref_heading(b$text), logical(1)))

  for (i in seq_along(source)) {
    b <- source[[i]]
    if (b$heading) {
      if (is_ref_heading(b$text)) {
        in_refs <- TRUE
        ref_level <- b$level
        collecting <- TRUE
        if (own_list) skip_ids <- c(skip_ids, i)
        next
      }
      if (in_refs && !is.na(b$level) && !is.na(ref_level) && b$level <= ref_level) {
        in_refs <- FALSE        # left the reference section
      } else if (in_refs) {
        collecting <- FALSE     # a sub-heading such as "Verification notes"
      }
      if (in_refs && own_list) skip_ids <- c(skip_ids, i)
      next
    }
    if (found_heading) {
      if (in_refs) {
        if (own_list) skip_ids <- c(skip_ids, i)   # back matter: do not parse as claims
        if (collecting && (looks_like_ref(b$text) || has_doi(b$text))) {
          refs[[length(refs) + 1]] <- parse_reference(b$text, length(refs) + 1L, b$location)
        }
      }
    } else {
      t <- b$text
      if (looks_like_ref(t) && (has_doi(t) || rx_detect(URL_RE, t))) {
        refs[[length(refs) + 1]] <- parse_reference(t, length(refs) + 1L, b$location)
        if (own_list) skip_ids <- c(skip_ids, i)
      }
    }
  }
  list(refs = refs, skip_ids = skip_ids)
}


# =============================================================================
# PART 4. Web lookups with caching
# =============================================================================
# Every web answer is stored in a JSON cache file, keyed by the web address,
# so reruns are fast and --no-network works. The cache file has the same
# layout as the Python version's, so either script can reuse the other's.
#
# The cache is an R "environment": a container that functions can change in
# place (unlike ordinary R lists, which are copied).

new_web_cache <- function(path, offline = FALSE, refresh = FALSE) {
  web <- new.env()
  web$path <- path
  web$offline <- offline
  web$refresh <- refresh
  web$data <- list()
  web$last_call <- list()
  web$new_calls <- 0L
  if (file.exists(path)) {
    loaded <- tryCatch(
      jsonlite::fromJSON(paste(readLines(path, encoding = "UTF-8", warn = FALSE), collapse = "\n"),
                         simplifyVector = FALSE),
      error = function(e) NULL
    )
    if (is.null(loaded)) {
      cat("  ! cache file was unreadable; starting a fresh cache\n")
    } else {
      web$data <- loaded
    }
  }
  web
}

web_save <- function(web) {
  dir.create(dirname(web$path), recursive = TRUE, showWarnings = FALSE)
  tmp <- sub("\\.[^.]*$", ".tmp", web$path)
  json <- if (length(web$data) == 0) "{}" else
    jsonlite::toJSON(web$data, auto_unbox = TRUE, null = "null", na = "null", digits = NA)
  writeBin(charToRaw(enc2utf8(as.character(json))), tmp)
  file.rename(tmp, web$path)
}

# One web request. Returns the HTTP status code (0 = network problem) and the
# text of the answer.
http_get <- function(url) {
  req <- httr2::request(url) |>
    httr2::req_user_agent(USER_AGENT) |>
    httr2::req_headers(Accept = "application/json") |>
    httr2::req_timeout(40) |>
    httr2::req_error(is_error = function(resp) FALSE)   # we read the status ourselves
  resp <- tryCatch(httr2::req_perform(req), error = function(e) NULL)
  if (is.null(resp)) return(list(code = 0L, body = ""))
  body <- tryCatch(httr2::resp_body_string(resp, encoding = "UTF-8"), error = function(e) "")
  list(code = as.integer(httr2::resp_status(resp)), body = body)
}

# Return list(status = <int>, body = <parsed JSON or NULL>).
# status 0 = network problem; status -1 = offline and not cached.
web_get_json <- function(web, url) {
  if (!is.null(web$data[[url]]) && !web$refresh) return(web$data[[url]])
  if (web$offline) return(list(status = -1L, body = NULL))

  host <- sub("^[A-Za-z][A-Za-z0-9+.-]*://([^/?#]*).*$", "\\1", url)
  result <- list(status = 0L, body = NULL)
  for (attempt in 0:3) {
    # Rate limit: wait so calls to the same host are >= ~1 s apart.
    last <- get_field(web$last_call, host, 0)
    wait <- MIN_SECONDS_BETWEEN_CALLS - (as.numeric(Sys.time()) - last)
    if (wait > 0) Sys.sleep(wait)
    web$last_call[[host]] <- as.numeric(Sys.time())
    ans <- http_get(url)
    code <- ans$code
    if (code == 200L) {
      parsed <- tryCatch(jsonlite::fromJSON(ans$body, simplifyVector = FALSE),
                         error = function(e) NULL)
      result <- list(status = 200L, body = parsed)
      break
    }
    if (code %in% c(404L, 400L, 410L)) {
      result <- list(status = code, body = NULL)
      break
    }
    # 429 (too many requests), 5xx, or network trouble: back off, retry.
    result <- list(status = code, body = NULL)
    Sys.sleep(3 * (2^attempt))
  }
  web$new_calls <- web$new_calls + 1L
  if (result$status %in% c(200L, 404L, 400L, 410L)) {
    web$data[[url]] <- result
    web_save(web)
  }
  result
}

# Percent-encode a DOI for a web address, leaving "/" alone
# (Python's urllib.parse.quote(doi, safe="/")).
quote_doi <- function(doi) {
  bytes <- as.integer(charToRaw(enc2utf8(doi)))
  safe <- c(48:57, 65:90, 97:122, utf8ToInt("_.-~/"))
  paste(vapply(bytes, function(b) if (b %in% safe) intToUtf8(b) else sprintf("%%%02X", b),
               character(1)), collapse = "")
}

crossref_lookup <- function(web, doi) {
  web_get_json(web, paste0("https://api.crossref.org/works/", quote_doi(doi)))
}
openalex_lookup <- function(web, doi) {
  web_get_json(web, paste0("https://api.openalex.org/works/doi:", quote_doi(doi)))
}
semanticscholar_lookup <- function(web, doi) {
  web_get_json(web, paste0("https://api.semanticscholar.org/graph/v1/paper/DOI:", quote_doi(doi),
                           "?fields=abstract,title,year,isOpenAccess,openAccessPdf"))
}

# OpenAlex stores abstracts as an 'inverted index' {word: [positions]}.
# Put the words back in order.
openalex_abstract <- function(work) {
  inv <- get_field(work, "abstract_inverted_index")
  if (!truthy(inv)) return("")
  words <- rep(names(inv), lengths(inv))
  places <- as.numeric(unlist(inv))
  keep <- !duplicated(places, fromLast = TRUE)     # a later word wins a shared position
  words <- words[keep]
  places <- places[keep]
  paste(words[order(places)], collapse = " ")
}


# =============================================================================
# PART 5. Comparing a reference with its Crossref record
# =============================================================================

# Publication years in the Crossref record: print, online, and "issued".
crossref_years <- function(msg) {
  out <- list()
  for (pair in list(c("published-print", "print"), c("published-online", "online"),
                    c("issued", "issued"))) {
    v <- tryCatch(msg[[pair[1]]][["date-parts"]][[1]][[1]], error = function(e) NULL)
    v <- suppressWarnings(as.integer(v))
    if (length(v) == 1 && !is.na(v)) out[[pair[2]]] <- v
  }
  out
}

# 'Proceedings of the National Academy of Sciences' -> 'pnas'.
acronym <- function(s) {
  small <- c("of", "the", "and", "in", "for", "on", "a", "an", "to")
  words <- py_split_ws(norm_text(s))
  words <- words[!words %in% small]
  paste(substr(words, 1, 1), collapse = "")
}

# Field-by-field comparison. Fills in the ref's check_* fields and status.
compare_reference <- function(ref, msg) {
  problems_major <- character(0)
  problems_minor <- character(0)
  notes <- character(0)
  ref_norm <- norm_text(ref$text)

  # --- First author ---
  authors <- py_or(msg$author, list())
  cr_families <- vapply(authors, function(a) py_or(a$family, py_or(a$name, "")), character(1))
  cr_first <- if (length(cr_families)) cr_families[1] else ""
  ref_first <- if (length(ref$surnames)) ref$surnames[1] else ""
  if (!nzchar(cr_first)) {
    editors <- py_or(msg$editor, list())
    cr_first <- if (length(editors)) py_or(editors[[1]]$family, "") else ""
  }
  first_ok <- nzchar(cr_first) && norm_name(ref_first) == norm_name(cr_first)
  if (!first_ok && nzchar(cr_first) && nzchar(norm_name(cr_first)) &&
      str_has(norm_name(ref_first), norm_name(cr_first))) {
    first_ok <- TRUE
  }
  ref$cr_first_author <- cr_first
  ref$cr_author_count <- length(authors)
  ref$check_author <- if (first_ok) "OK" else if (!nzchar(cr_first)) "no author in Crossref" else "MISMATCH"

  # Other listed authors: are they all present, in the same order?
  cr_norm <- vapply(cr_families, norm_name, character(1), USE.NAMES = FALSE)
  listed <- vapply(ref$surnames, norm_name, character(1), USE.NAMES = FALSE)
  missing <- if (length(listed) > 1) ref$surnames[-1][!listed[-1] %in% cr_norm] else character(0)
  if (length(cr_norm) && length(listed) > 1 && length(missing) == 0) {
    if (!identical(listed, utils::head(cr_norm, length(listed)))) {
      problems_minor <- c(problems_minor, "author order differs from Crossref")
    }
  }
  if (length(missing) && length(cr_norm)) {
    problems_minor <- c(problems_minor,
                        paste0("author(s) not in Crossref list: ", paste(missing, collapse = ", ")))
  }
  if (length(cr_norm) && !ref$et_al && length(listed) != length(cr_norm) && !ref$corporate) {
    problems_minor <- c(problems_minor, sprintf("reference lists %d authors, Crossref has %d",
                                                length(listed), length(cr_norm)))
  }

  # --- Year (print vs online can differ by a year) ---
  years <- crossref_years(msg)
  ref$cr_years <- if (length(years)) paste(names(years), unlist(years), sep = " ", collapse = "; ") else "none"
  yr <- ref$year
  if (length(years) == 0) {
    ref$check_year <- "not in Crossref (unverified)"
  } else if (is.na(yr)) {
    ref$check_year <- "no year in reference"
    problems_minor <- c(problems_minor, "year missing")
  } else if (yr %in% c(years$print, years$issued)) {
    ref$check_year <- "OK"
  } else if (any(abs(yr - unlist(years)) <= 1)) {
    ref$check_year <- paste0("within 1 year (", ref$cr_years, ")")
    if (is.null(years$online) || yr != years$online) {
      problems_minor <- c(problems_minor, paste0("year ", yr, " vs Crossref ", ref$cr_years))
    }
  } else {
    ref$check_year <- "MISMATCH"
    problems_minor <- c(problems_minor, paste0("year ", yr, " vs Crossref ", ref$cr_years))
  }

  # --- Title ---
  cr_title <- paste(unlist(py_or(msg$title, list())), collapse = " ")
  cr_sub <- paste(unlist(py_or(msg$subtitle, list())), collapse = " ")
  cr_full <- if (nzchar(cr_sub)) paste0(cr_title, ": ", cr_sub) else cr_title
  ref$cr_title <- clean_abstract(cr_full)
  rest_norm <- norm_text(ref$rest)
  cands <- c(similarity(ref$title, cr_full), similarity(ref$title, cr_title))
  for (t in c(cr_full, cr_title)) {
    tn <- norm_text(t)
    if (nzchar(tn) && str_has(rest_norm, tn)) {
      cands <- c(cands, 1)
    } else if (nzchar(tn)) {
      cands <- c(cands, seq_ratio(tn, py_slice(rest_norm, 0, nchar(tn) + 5)))
    }
  }
  sim <- if (nzchar(cr_title)) max(cands) else 0
  ref$title_similarity <- round(sim, 2)
  if (sim < 0.6) {
    problems_major <- c(problems_major, sprintf("title does not match (similarity %.2f)", sim))
  } else if (sim < 0.9) {
    problems_minor <- c(problems_minor, sprintf("title differs somewhat (similarity %.2f)", sim))
  }

  # --- Journal / book (container) ---
  containers <- c(unlist(py_or(msg[["container-title"]], list())),
                  unlist(py_or(msg[["short-container-title"]], list())))
  ref$cr_journal <- unlist(py_or(msg[["container-title"]], list("")))[1]
  ctype <- py_or(msg$type, "")
  ref$cr_type <- ctype
  if (length(containers) == 0) {
    ref$check_journal <- if (ctype != "journal-article") "n/a (no container in Crossref)" else "not in Crossref"
  } else {
    ok <- FALSE
    tail_norm <- py_or(norm_text(ref$tail), rest_norm)
    for (c in containers) {
      cn <- norm_text(c)
      if (nzchar(cn) && str_has(ref_norm, cn)) ok <- TRUE
      jn <- norm_text(ref$journal)
      if (nzchar(ref$journal) && nzchar(jn) && str_has(cn, jn)) ok <- TRUE
      if (nzchar(ref$journal) && similarity(ref$journal, c) >= 0.85) ok <- TRUE
      # Fuzzy "is it somewhere in the reference?" check, which tolerates
      # typos in Crossref such as 'ECOLOGYAND' (compare a same-length window).
      if (nzchar(cn) && nchar(cn) >= 8 && !ok) {
        sm <- seq_matcher(ref_norm, cn, autojunk = FALSE)
        blk <- find_longest_match(sm, 0L, nchar(ref_norm), 0L, nchar(cn))
        start <- max(0, blk$i - blk$j)
        window <- py_slice(ref_norm, start, start + nchar(cn) + 3)
        if (seq_ratio(window, cn) >= 0.85) ok <- TRUE
      }
      ac <- acronym(c)
      if (nzchar(ac) && nchar(ac) >= 3 && rx_detect(paste0("\\b", ac, "\\b"), tail_norm)) ok <- TRUE
    }
    ref$check_journal <- if (ok) "OK" else "MISMATCH"
    if (!ok) {
      problems_minor <- c(problems_minor, paste0("journal/book '", ref$journal,
                                                 "' vs Crossref '", ref$cr_journal, "'"))
    }
  }

  # Text after the title, without DOI/URL, is where volume and pages live.
  tail <- fix_dashes(py_or(ref$tail, ref$rest))

  # --- Volume ---
  vol <- py_strip(as.character(py_or(msg$volume, "")))
  ref$cr_volume <- vol
  if (!nzchar(vol)) {
    ref$check_volume <- "not in Crossref"
  } else if (ctype != "journal-article") {
    ref$check_volume <- "n/a (not a journal article)"
  } else if (rx_detect(paste0("(?<![\\w-])", rx_literal(vol), "(?![\\w-])"), tail)) {
    ref$check_volume <- "OK"
  } else {
    ref$check_volume <- "MISMATCH"
    problems_minor <- c(problems_minor, paste0("volume: Crossref says ", vol))
  }

  # --- Pages or article number ---
  page <- fix_dashes(py_strip(as.character(py_or(msg$page, ""))))
  artno <- py_strip(as.character(py_or(msg[["article-number"]], "")))
  ref$cr_pages <- if (nzchar(page)) page else if (nzchar(artno)) paste0("article ", artno) else ""
  artno_in_tail <- function() rx_detect(paste0("(?<![\\w-])", rx_literal(artno), "(?![\\w-])"), tail)
  if (nzchar(page)) {
    dash <- regexpr("-", page, fixed = TRUE)
    first <- if (dash > 0) substr(page, 1, dash - 1) else page
    last <- if (dash > 0) substring(page, dash + 1) else ""
    is_digits <- function(x) rx_detect("^\\d+$", x)
    if (nzchar(last) && nchar(last) < nchar(first) && is_digits(last) && is_digits(first)) {
      last <- paste0(substr(first, 1, nchar(first) - nchar(last)), last)   # "1367-81" -> "1367-1381"
    }
    ref_range <- rx_search(paste0("(?<![\\w-])", rx_literal(first), "\\s*-\\s*(\\w+)"), tail)
    if (nzchar(last) && !is.null(ref_range) && ref_range$g[1] == last) {
      ref$check_pages <- "OK"
    } else if (nzchar(last) && !is.null(ref_range)) {
      ref$check_pages <- "MISMATCH"
      problems_minor <- c(problems_minor, paste0("pages: reference has ", first, "-", ref_range$g[1],
                                                 ", Crossref ", first, "-", last))
    } else if (!nzchar(last) && rx_detect(paste0("(?<![\\w-])", rx_literal(first), "(?![\\w])"), tail)) {
      ref$check_pages <- "first page OK (Crossref gives only the first page; last page unverified)"
    } else if (nzchar(last) && ctype %in% c("book-chapter", "book") && str_has(tail, first)) {
      ref$check_pages <- "OK"
    } else {
      ref$check_pages <- "MISMATCH"
      problems_minor <- c(problems_minor, paste0("pages: Crossref says ", page))
    }
  } else if (nzchar(artno) && artno_in_tail()) {
    ref$check_pages <- "OK"
  } else if (nzchar(artno) && str_has(py_lower(ref$doi), py_lower(artno))) {
    # e.g. AGU papers: Crossref's "article number" is just the DOI suffix,
    # not the citation number printed on the paper (e.g. G03011).
    ref$check_pages <- paste0("unverified: Crossref's article ID is the DOI suffix; ",
                              "the paper's own citation number is not in Crossref")
  } else if (nzchar(artno)) {
    if (artno_in_tail()) {
      ref$check_pages <- "OK"
    } else {
      ref$check_pages <- "MISMATCH"
      problems_minor <- c(problems_minor, paste0("article number: Crossref says ", artno))
    }
  } else {
    ref$check_pages <- "not in Crossref (unverified)"
  }

  # --- Status ---
  if (!first_ok && nzchar(cr_first)) {
    msg_txt <- paste0("first author ", ref_first, " vs Crossref ", cr_first)
    if (sim < 0.85) problems_major <- c(problems_major, msg_txt) else problems_minor <- c(problems_minor, msg_txt)
  }
  ref$status <- if (length(problems_major)) "MAJOR MISMATCH" else
    if (length(problems_minor)) "MINOR MISMATCH" else "OK"
  ref$problems <- paste(c(problems_major, problems_minor, notes), collapse = "; ")
  ref
}

# Look up every reference online (or in the cache) and compare.
check_references <- function(refs, web, want_abstracts = TRUE) {
  for (n in seq_along(refs)) {
    ref <- refs[[n]]
    blank <- list(abstract = "", abstract_source = "", oa_pdf = "", cr_title = "",
                  cr_journal = "", cr_years = "", cr_volume = "", cr_pages = "",
                  cr_first_author = "", cr_type = "", title_similarity = "",
                  check_author = "", check_year = "", check_journal = "",
                  check_volume = "", check_pages = "", problems = "")
    ref[names(blank)] <- blank
    label <- paste0("[", n, "/", length(refs), "] ", ref$short)
    if (!nzchar(ref$doi)) {
      ref$status <- "MANUAL CHECK"
      ref$problems <- paste0("no DOI: check by hand",
                             if (nzchar(ref$url)) paste0(" at ", ref$url) else " (no URL given)")
      cat("  ", label, ": no DOI -> manual check\n", sep = "")
      refs[[n]] <- ref
      next
    }
    cr <- crossref_lookup(web, ref$doi)
    if (cr$status == -1) {
      ref$status <- "NOT CHECKED"
      ref$problems <- "not in cache; rerun without --no-network"
      cat("  ", label, ": not in cache\n", sep = "")
      refs[[n]] <- ref
      next
    }
    if (cr$status != 200 || !truthy(cr$body)) {
      ref$status <- if (cr$status %in% c(404, 400, 410)) "NOT FOUND" else "NOT CHECKED"
      ref$problems <- if (ref$status == "NOT FOUND") {
        paste0("DOI not found on Crossref (HTTP ", cr$status, ")")
      } else {
        paste0("network problem (HTTP ", cr$status, "); rerun later")
      }
      cat("  ", label, ": ", ref$status, "\n", sep = "")
      refs[[n]] <- ref
      next
    }
    msg <- get_field(cr$body, "message", list())
    ref <- compare_reference(ref, msg)
    cat("  ", label, ": ", ref$status,
        if (nzchar(ref$problems)) paste0("  (", ref$problems, ")") else "", "\n", sep = "")

    if (!want_abstracts) {
      refs[[n]] <- ref
      next
    }
    # Abstract: Crossref first ...
    ab <- clean_abstract(get_field(msg, "abstract", ""))
    if (nzchar(ab)) {
      ref$abstract <- ab
      ref$abstract_source <- "Crossref"
    }
    # ... OpenAlex (also gives an open-access PDF link) ...
    oa <- openalex_lookup(web, ref$doi)
    work <- py_or(oa$body, list())
    if (!nzchar(ref$abstract)) {
      ab <- openalex_abstract(work)
      if (nzchar(ab)) {
        ref$abstract <- ab
        ref$abstract_source <- "OpenAlex"
      }
    }
    loc <- py_or(get_field(work, "best_oa_location"), list())
    ref$oa_pdf <- py_or(get_field(loc, "pdf_url"),
                        py_or(get_field(py_or(get_field(work, "open_access"), list()), "oa_url"), ""))
    # ... then Semantic Scholar.
    if (!nzchar(ref$abstract) || !nzchar(ref$oa_pdf)) {
      s2 <- semanticscholar_lookup(web, ref$doi)
      body <- py_or(s2$body, list())
      if (!nzchar(ref$abstract) && truthy(get_field(body, "abstract"))) {
        ref$abstract <- py_strip(body$abstract)
        ref$abstract_source <- "Semantic Scholar"
      }
      if (!nzchar(ref$oa_pdf)) {
        ref$oa_pdf <- py_or(get_field(py_or(get_field(body, "openAccessPdf"), list()), "url"), "")
      }
    }
    if (!nzchar(ref$abstract)) ref$abstract_source <- "none found (read the paper)"
    refs[[n]] <- ref
  }
  refs
}


# =============================================================================
# PART 6. Finding in-text citations
# =============================================================================
# These patterns are the same as the Python ones. "(?P<name>...)" marks a
# piece of the match we want to read back by name.

UP <- "A-Z\u00C0-\u00D6\u00D8-\u00DE"                  # capital letters incl. accented
PARTICLE <- "(?:(?:[Vv]an|[Vv]on|[Dd]e|[Dd]el|[Dd]er|[Dd]en|[Dd]a|[Dd]i|[Dd]u|[Ll]e|[Ll]a|[Tt]er|dos|das)\\s+)"
SURNAME <- paste0("(?<![\\w'\u2019\\-])", PARTICLE, "{0,2}[", UP,
                  "](?:[^\\W\\d_]|['\u2019\\-](?=[^\\W\\d_]))+")
YEAR <- "(?:1[5-9]|20)\\d{2}"
YEARLIST <- paste0("(?P<years>", YEAR, "[a-z]?(?:\\s*,\\s*(?:", YEAR, "[a-z]?|[a-z])(?!\\w))*)")
AUTHORS <- paste0("(?P<a1>", SURNAME, ")(?:\\s+(?P<etal>et\\s+al\\.?)|(?:\\s*&\\s*|\\s+and\\s+)(?P<a2>",
                  SURNAME, "))?")
# "Nixon et al. (2009)"
NARRATIVE_RE <- paste0(AUTHORS, "\\s*\\(\\s*", YEARLIST, "(?:\\s*,\\s*pp?\\.[^)]*)?\\s*\\)")
# "(Koch 2001)", "Koch 2001; Cloern 1996", "Lucas et al. 1999a, b"
PAREN_RE <- paste0(AUTHORS, ",?\\s+", YEARLIST, "(?![\\d\u2013\\-])")
# Full reference style used in reading lists: "Cloern, J.E. (1996)",
# "Smayda, T.J. & Borkman, D.G. (2008)"
INITS <- "(?:[A-Z][a-z]?\\.\\s?-?\\s?)+"
REFSTYLE_RE <- paste0(
  "(?P<rs>", SURNAME, ",\\s*", INITS, "(?:,\\s*", SURNAME, ",\\s*", INITS, ")*",
  "(?:,?\\s*(?:&|and)\\s*", SURNAME, ",\\s*", INITS, ")?(?:,?\\s*et\\s+al\\.)?)\\s*\\((?P<years>",
  YEAR, "[a-z]?)\\)")

# Capitalized words that sit next to a year but are not authors.
STOPWORDS <- tolower(strsplit(paste(
  "January February March April May June July August September October November December",
  "Jan Feb Mar Apr Jun Jul Aug Sep Sept Oct Nov Dec Spring Summer Fall Autumn Winter",
  "Since In By From Through Between And Or The Of To For On At Before After Until During",
  "Circa Around About Year Years Fig Figure Table Tables Group Lecture Slide Slides Reading",
  "Draft Version Revised Updated Accessed Copyright Est Est. Term Semester Week Class",
  "Monday Tuesday Wednesday Thursday Friday Saturday Sunday Present Today Circa FY CY"
), "\\s+")[[1]])

# '1999a, b' -> (1999,'a'), (1999,'b');  '2017, 2022' -> (2017,''), (2022,'').
expand_years <- function(years_text) {
  out <- list()
  last_year <- NA_integer_
  for (part in rx_split("\\s*,\\s*", py_strip(years_text))) {
    m <- rx_search(paste0("^(", YEAR, ")([a-z]?)\\z"), part)
    if (!is.null(m)) {
      last_year <- as.integer(m$g[1])
      out[[length(out) + 1]] <- list(year = last_year, suffix = m$g[2])
    } else if (rx_full("[a-z]", part) && !is.na(last_year) && last_year != 0) {
      out[[length(out) + 1]] <- list(year = last_year, suffix = part)
    }
  }
  out
}

# Return a list of citations found in one block of text.
find_citations <- function(text) {
  found <- list()
  taken <- list()
  overlaps <- function(s, e) any(vapply(taken, function(t) s < t[2] && e > t[1], logical(1)))

  # 1) Reference-style mentions first ("Cloern, J.E. (1996)").
  for (m in rx_find_all(REFSTYLE_RE, text)) {
    rs <- m$g[["rs"]]
    seg <- gsub(" and ", ",", gsub("&", ",", rx_replace("et\\s+al\\.?", "", rs), fixed = TRUE), fixed = TRUE)
    names_ <- py_strip(py_split(seg, ","))
    names_ <- names_[nzchar(names_) & !vapply(names_, is_initials, logical(1))]
    etal <- rx_detect("et\\s+al", rs) || length(names_) >= 3
    for (ys in expand_years(m$g[["years"]])) {
      found[[length(found) + 1]] <- list(raw = m$text, a1 = names_[1],
                                         a2 = if (length(names_) == 2) names_[2] else "",
                                         et_al = etal, year = ys$year, suffix = ys$suffix,
                                         start = m$start)
    }
    taken[[length(taken) + 1]] <- c(m$start, m$end)
  }

  # 2) Narrative "(Author year)" then parenthetical "Author year".
  for (pattern in c(NARRATIVE_RE, PAREN_RE)) {
    for (m in rx_find_all(pattern, text)) {
      if (overlaps(m$start, m$end)) next
      a1 <- m$g[["a1"]]
      a2 <- m$g[["a2"]]
      if (is.na(a2)) a2 <- ""
      a1_words <- py_split_ws(a1)
      last_word <- a1_words[length(a1_words)]
      if (strip_chars(py_lower(last_word), ".") %in% STOPWORDS) next
      if (nzchar(a2)) {
        a2_words <- py_split_ws(a2)
        if (py_lower(a2_words[length(a2_words)]) %in% STOPWORDS) next
      }
      etal <- m$g[["etal"]]
      for (ys in expand_years(m$g[["years"]])) {
        found[[length(found) + 1]] <- list(raw = m$text, a1 = a1, a2 = a2,
                                           et_al = !is.na(etal) && nzchar(etal),
                                           year = ys$year, suffix = ys$suffix, start = m$start)
      }
      taken[[length(taken) + 1]] <- c(m$start, m$end)
    }
  }
  if (length(found) == 0) return(found)
  starts <- vapply(found, function(c) c$start, numeric(1))
  found[order(starts, method = "radix")]     # stable sort, like Python's
}

cite_label <- function(c) {
  a <- paste0(c$a1, if (c$et_al) " et al." else if (nzchar(c$a2)) paste0(" & ", c$a2) else "")
  paste0(a, " ", c$year, c$suffix)
}

# One row per in-text citation, with the sentence around it.
collect_claims <- function(blocks, skip_ids) {
  claims <- list()
  for (i in seq_along(blocks)) {
    b <- blocks[[i]]
    if (i %in% skip_ids || b$heading || isTRUE(b$header_row)) next
    cites <- find_citations(b$text)
    if (length(cites) == 0) next
    sentences <- split_sentences(b$text)
    # Work out where each sentence starts and ends in the original text, so a
    # citation is paired with the sentence at ITS position (the same citation
    # can appear twice in one paragraph).
    spans <- list()
    pos <- 0
    for (s in sentences) {
      start <- py_find(b$text, s, pos)
      if (start < 0) start <- pos
      spans[[length(spans) + 1]] <- list(st = start, en = start + nchar(s), s = s)
      pos <- start + nchar(s)
    }
    for (c in cites) {
      claim <- b$text
      for (sp in spans) {
        if (sp$st <= c$start && c$start < sp$en) {
          claim <- sp$s
          break
        }
      }
      # A full reference written out in a reading list ("Cloern, J.E.
      # (1996). Title...") is best shown whole.
      if (!is.null(rx_match_start(REFSTYLE_RE, c$raw)) && nchar(claim) < 60) claim <- b$text
      claims[[length(claims) + 1]] <- list(
        citation = cite_label(c), raw = c$raw, cite = c, location = b$location,
        claim = claim, context = py_or(b$context, b$text),
        ref_i = NA_integer_, match_status = "", match_note = "",
        passages = character(0), pdf_name = "")
    }
  }
  claims
}


# =============================================================================
# PART 7. Matching in-text citations to references
# =============================================================================
# Surname + year, tolerant of a/b suffixes and +/-1 year. Uses 'et al.' and
# '&' to break ties (e.g. Nixon 1988 vs Nixon et al. 2009).
# Each claim stores the NUMBER of its matched reference (ref_i), and each
# reference collects where it was cited (cited_in). Both lists come back.

form_fits <- function(r, c) {
  n <- length(r$surnames)
  if (nzchar(c$a2)) return(n == 2 && !r$et_al && norm_name(r$surnames[2]) == norm_name(c$a2))
  if (c$et_al) return(n >= 3 || r$et_al)
  (n == 1 && !r$et_al) || r$corporate
}

match_claims <- function(claims, refs) {
  for (k in seq_along(refs)) if (is.null(refs[[k]]$cited_in)) refs[[k]]$cited_in <- character(0)
  for (ci in seq_along(claims)) {
    cl <- claims[[ci]]
    c <- cl$cite
    key <- norm_name(c$a1)
    cands <- which(vapply(refs, function(r) key %in% r$keys, logical(1)))
    # A surname with a particle might be keyed without it (or vice versa).
    if (length(cands) == 0) {
      bare <- norm_name(rx_replace(PARTICLE, "", c$a1))
      cands <- if (nchar(bare) > 2) {
        which(vapply(refs, function(r) any(ends_with(r$keys, bare)), logical(1)))
      } else integer(0)
    }
    year_of <- function(i) refs[[i]]$year
    exact <- cands[vapply(cands, function(i) {
      isTRUE(year_of(i) == c$year) && (!nzchar(c$suffix) || refs[[i]]$suffix == c$suffix)
    }, logical(1))]
    if (length(exact) == 0 && nzchar(c$suffix)) {
      exact <- cands[vapply(cands, function(i) isTRUE(year_of(i) == c$year), logical(1))]
    }
    note <- character(0)
    pool <- exact
    if (length(pool) == 0) {
      pool <- cands[vapply(cands, function(i) {
        !is.na(year_of(i)) && year_of(i) != 0 && abs(year_of(i) - c$year) == 1
      }, logical(1))]
      if (length(pool)) {
        note <- c(note, paste0("year differs: in text ", c$year, ", reference ", refs[[pool[1]]]$year))
      }
    }
    if (length(pool) > 1) {
      fitting <- pool[vapply(pool, function(i) form_fits(refs[[i]], c), logical(1))]
      if (length(fitting) >= 1) pool <- fitting
    }
    if (length(pool) > 1) {
      cl$ref_i <- NA_integer_
      cl$match_status <- "AMBIGUOUS"
      cl$match_note <- paste0("could be: ",
                              paste(vapply(pool, function(i) refs[[i]]$short, character(1)), collapse = "; "),
                              " (add a/b year letters or co-author names)")
      claims[[ci]] <- cl
      next
    }
    if (length(pool) == 0) {
      cl$ref_i <- NA_integer_
      cl$match_status <- "NO REFERENCE"
      others <- py_sort(unique(vapply(cands, function(i) refs[[i]]$short, character(1))))
      cl$match_note <- paste0("no reference with this author and year",
                              if (length(others)) paste0("; same author in list: ",
                                                         paste(others, collapse = ", ")) else "")
      claims[[ci]] <- cl
      next
    }
    r <- refs[[pool[1]]]
    cl$ref_i <- pool[1]
    if (c$suffix != r$suffix) {
      if (nzchar(r$suffix) && !nzchar(c$suffix)) {
        note <- c(note, paste0("reference is ", r$year, r$suffix, ": add the letter in text"))
      } else if (nzchar(c$suffix) && !nzchar(r$suffix)) {
        note <- c(note, paste0("in-text letter '", c$suffix, "' but reference has no letter"))
      }
    }
    if (!r$corporate && !form_fits(r, c)) {
      n <- length(r$surnames)
      want <- if (n == 1 && !r$et_al) r$surnames[1] else
        if (n == 2 && !r$et_al) paste0(r$surnames[1], " & ", r$surnames[2]) else
          paste0(r$surnames[1], " et al.")
      note <- c(note, paste0("citation form: reference suggests '", want, " ", r$year, "'"))
    }
    cl$match_status <- if (length(note) == 0) "MATCHED" else "MATCHED (check note)"
    cl$match_note <- paste(note, collapse = "; ")
    refs[[pool[1]]]$cited_in <- c(refs[[pool[1]]]$cited_in, cl$location)
    claims[[ci]] <- cl
  }
  list(claims = claims, refs = refs)
}


# =============================================================================
# PART 8. PDFs: match to references, then find passages for each claim
# =============================================================================

STOP <- strsplit(paste(
  "a about above after again against all also am an and any are as at be because been",
  "before being below between both but by can could did do does doing down during each few for from",
  "further had has have having here how i if in into is it its itself just more most no nor not now of",
  "off on once only or other our out over own same should so some such than that the their them then",
  "there these they this those through to too under until up very was we were what when where which",
  "while who whom why will with would you your et al e g i ie eg vs using use used can may might shows",
  "show showed shown one two three via per within"), "\\s+")[[1]]

# Words used for passage search: lower case, no accents, no common words,
# crude plural stripping (blooms -> bloom).
tokens <- function(text) {
  words <- regmatches(strip_accents(py_lower(text)),
                      gregexpr(rx("[a-z][a-z0-9\\-]{2,}"), strip_accents(py_lower(text)), perl = TRUE))[[1]]
  words <- words[!words %in% STOP]
  plural <- nchar(words) > 4 & ends_with(words, "s") & !ends_with(words, "ss")
  words[plural] <- substr(words[plural], 1, nchar(words[plural]) - 1)
  words
}

# Extract text from every PDF (sub-folders included). Each PDF is an
# environment so later steps can store results on it.
load_pdfs <- function(pdf_dir) {
  paths <- list.files(pdf_dir, pattern = "\\.pdf$", recursive = TRUE, full.names = TRUE,
                      all.files = TRUE)
  # Sort folder by folder, as Python sorts paths.
  paths <- paths[order(gsub("/", "\001", paths, fixed = TRUE), method = "radix")]
  pdfs <- list()
  for (path in paths) {
    pages <- tryCatch(suppressMessages(pdftools::pdf_text(path)), error = function(e) e)
    if (inherits(pages, "error")) {
      cat("  ! could not read ", basename(path), ": ", conditionMessage(pages), "\n", sep = "")
      next
    }
    pages[is.na(pages)] <- ""
    chars <- sum(nchar(pages))
    pdf <- new.env()
    pdf$path <- path
    pdf$name <- basename(path)
    pdf$pages <- pages
    pdf$has_text <- chars > 500
    pdf$ref_i <- NA_integer_
    pdf$how <- ""
    pdfs[[length(pdfs) + 1]] <- pdf
    cat("  read ", pdf$name, ": ", length(pages), " pages",
        if (chars > 500) "" else " (little/no text: scanned? needs OCR)", "\n", sep = "")
  }
  pdfs
}

# DOI on the first pages, then title similarity, then 'Author_Year' in the
# file name.
match_pdfs <- function(pdfs, refs) {
  for (k in seq_along(refs)) if (is.null(refs[[k]]$pdfs)) refs[[k]]$pdfs <- integer(0)
  for (p_i in seq_along(pdfs)) {
    pdf <- pdfs[[p_i]]
    front <- paste(utils::head(pdf$pages, 2), collapse = " ")
    front_norm <- norm_text(front)
    match <- NA_integer_
    how <- ""
    dois <- unique(py_lower(vapply(rx_find_all(DOI_RE, front, ignore_case = TRUE),
                                   function(m) clean_doi(m$g[1]), character(1))))
    for (i in seq_along(refs)) {
      if (nzchar(refs[[i]]$doi) && py_lower(refs[[i]]$doi) %in% dois) {
        match <- i
        how <- "DOI on first pages"
        break
      }
    }
    if (is.na(match)) {
      best <- NA_integer_
      best_score <- 0
      for (i in seq_along(refs)) {
        title <- py_or(refs[[i]]$cr_title, refs[[i]]$title)
        tn <- norm_text(title)
        if (nchar(tn) < 15) next
        if (str_has(front_norm, tn)) {
          score <- 1
        } else {
          front_cut <- py_slice(front_norm, 0, 4000)
          sm <- seq_matcher(front_cut, tn, autojunk = FALSE)
          blk <- find_longest_match(sm, 0L, min(nchar(front_norm), 4000L), 0L, nchar(tn))
          score <- blk$size / nchar(tn)
        }
        if (score > best_score) {
          best <- i
          best_score <- score
        }
      }
      if (!is.na(best) && best_score >= 0.6) {
        match <- best
        how <- sprintf("title on first pages (%.2f)", best_score)
      }
    }
    if (is.na(match)) {
      fname <- norm_text(path_stem(pdf$path))
      for (i in seq_along(refs)) {
        r <- refs[[i]]
        if (length(r$surnames) && !is.na(r$year) && r$year != 0 &&
            str_has(gsub(" ", "", fname, fixed = TRUE), norm_name(r$surnames[1])) &&
            str_has(fname, as.character(r$year))) {
          match <- i
          how <- "author + year in file name"
          break
        }
      }
    }
    pdf$ref_i <- match
    pdf$how <- how
    if (!is.na(match)) {
      refs[[match]]$pdfs <- c(refs[[match]]$pdfs, p_i)
      cat("  ", pdf$name, " -> ", refs[[match]]$short, " (", how, ")\n", sep = "")
    } else {
      cat("  ", pdf$name, " -> no matching reference\n", sep = "")
    }
  }
  refs
}

# Cut the paper into overlapping 3-sentence windows, remembering pages.
build_windows <- function(pdf) {
  if (!is.null(pdf$windows)) return(pdf$windows)
  sents <- list()
  for (pno in seq_along(pdf$pages)) {
    text <- rx_replace("-\\s*\\n\\s*", "", pdf$pages[pno])     # re-join hyphenated words
    text <- rx_replace("\\s+", " ", text)
    for (s in split_sentences(text)) {
      if (nchar(s) > 25) sents[[length(sents) + 1]] <- list(page = pno, s = s)
    }
  }
  windows <- list()
  for (i in seq_along(sents)) {
    chunk <- sents[i:min(i + 2, length(sents))]
    pages <- sort(unique(vapply(chunk, function(x) x$page, integer(1))))
    txt <- paste(vapply(chunk, function(x) x$s, character(1)), collapse = " ")
    tk <- tokens(txt)
    counts <- if (length(tk)) table(tk) else table(character(0))
    windows[[i]] <- list(i = i - 1L, pages = pages, text = txt,
                         toks = stats::setNames(as.integer(counts), names(counts)))
  }
  # Inverse document frequency: rare words in this paper count more.
  all_toks <- unlist(lapply(windows, function(w) names(w$toks)))
  df <- if (length(all_toks)) table(all_toks) else table(character(0))
  n <- max(length(windows), 1)
  pdf$idf <- stats::setNames(log((n + 1) / (as.numeric(df) + 0.5)), names(df))
  pdf$windows <- windows
  windows
}

best_passages <- function(pdf, query, k = 3) {
  windows <- build_windows(pdf)
  q <- unique(tokens(query))
  if (length(q) == 0 || length(windows) == 0) return(character(0))
  idf <- pdf$idf
  scores <- numeric(0)
  keep <- integer(0)
  for (w_i in seq_along(windows)) {
    w <- windows[[w_i]]
    shared <- intersect(q, names(w$toks))
    if (length(shared) == 0) next
    idf_shared <- idf[shared]
    idf_shared[is.na(idf_shared)] <- 0
    score <- sum(idf_shared * (1 + log(w$toks[shared])))
    score <- score / sqrt(sum(w$toks) + 10)
    scores <- c(scores, score)
    keep <- c(keep, w_i)
  }
  ord <- keep[order(-scores, method = "radix")]
  chosen <- integer(0)
  for (w_i in ord) {
    wi <- windows[[w_i]]$i
    if (all(abs(wi - vapply(chosen, function(c) windows[[c]]$i, integer(1))) >= 3)) {
      chosen <- c(chosen, w_i)   # no overlapping windows
    }
    if (length(chosen) == k) break
  }
  vapply(chosen, function(w_i) {
    w <- windows[[w_i]]
    pg <- if (length(w$pages) == 1) paste0("p. ", w$pages[1]) else
      paste0("pp. ", w$pages[1], "-", w$pages[length(w$pages)])
    paste0("[", pg, "] ", truncate_text(w$text, 900))
  }, character(1))
}

attach_passages <- function(claims, refs, pdfs) {
  for (ci in seq_along(claims)) {
    cl <- claims[[ci]]
    cl$passages <- character(0)
    cl$pdf_name <- ""
    if (is.na(cl$ref_i) || length(refs[[cl$ref_i]]$pdfs) == 0) {
      claims[[ci]] <- cl
      next
    }
    ref_pdfs <- pdfs[refs[[cl$ref_i]]$pdfs]
    with_text <- Filter(function(p) p$has_text, ref_pdfs)
    pdf <- if (length(with_text)) with_text[[1]] else ref_pdfs[[1]]
    cl$pdf_name <- pdf$name
    if (!pdf$has_text) {
      cl$passages <- "(PDF has no extractable text; it may be a scan that needs OCR)"
      claims[[ci]] <- cl
      next
    }
    # Search with the claim minus the citations themselves.
    query <- rx_replace("\\([^)]*\\d{4}[^)]*\\)", " ", cl$claim)
    if (length(tokens(query)) < 4) query <- cl$context
    cl$passages <- best_passages(pdf, query)
    claims[[ci]] <- cl
  }
  claims
}


# =============================================================================
# PART 9. Writing the Excel report and Markdown summary
# =============================================================================

GREEN <- "C6EFCE"; AMBER <- "FFEB9C"; RED <- "FFC7CE"; GREY <- "E7E6E6"; BLUE <- "DDEBF7"
STATUS_COLOR <- c("OK" = GREEN, "MATCHED" = GREEN, "MINOR MISMATCH" = AMBER,
                  "MATCHED (check note)" = AMBER, "MANUAL CHECK" = AMBER, "AMBIGUOUS" = AMBER,
                  "MAJOR MISMATCH" = RED, "NOT FOUND" = RED, "NO REFERENCE" = RED,
                  "NOT CHECKED" = GREY, "NOT CITED" = BLUE)

# Remove characters Excel refuses, and cut text to Excel's cell limit.
excel_safe <- function(v) {
  if (is.character(v) && length(v) == 1 && !is.na(v)) {
    v <- gsub("[\u0001-\u0008\u000B\u000C\u000E-\u001F]", "", v, perl = TRUE)
    return(truncate_text(v, EXCEL_CELL_LIMIT))
  }
  v
}

# Spreadsheet column letters: 1 -> "A", 19 -> "S".
col_letter <- function(n) openxlsx2::int2col(n)

# Write one sheet: blue header row, wrapped text, coloured status cells,
# frozen panes and a filter. `rows` is a list of rows; each row is a list of
# values (text, numbers, or NA for an empty cell).
write_sheet <- function(wb, title, headers, rows, widths, status_cols = integer(0)) {
  wb$add_worksheet(title)
  n_col <- length(headers)
  n_row <- length(rows)
  # Build the table one column at a time. A column of numbers stays numeric;
  # a column that mixes text and numbers (e.g. the Summary's second column)
  # is written as text first and then the numbers are written cell by cell,
  # so Excel sees numbers as numbers, as in the Python version.
  cols <- list()
  mixed_numbers <- list()
  for (j in seq_len(n_col)) {
    vals <- lapply(rows, function(r) excel_safe(r[[j]]))
    is_num <- vapply(vals, function(v) is.numeric(v) && length(v) == 1 && !is.na(v), logical(1))
    is_blank <- vapply(vals, function(v) is.null(v) || length(v) == 0 || (length(v) == 1 && is.na(v)),
                       logical(1))
    if (n_row > 0 && all(is_num | is_blank)) {
      cols[[j]] <- vapply(seq_along(vals), function(i) if (is_num[i]) as.numeric(vals[[i]]) else NA_real_,
                          numeric(1))
    } else {
      cols[[j]] <- vapply(seq_along(vals), function(i) {
        if (is_blank[i] || is_num[i]) NA_character_ else as.character(vals[[i]])
      }, character(1))
      if (any(is_num)) mixed_numbers[[j]] <- which(is_num)
    }
  }
  df <- as.data.frame(stats::setNames(cols, headers), check.names = FALSE, stringsAsFactors = FALSE)
  if (n_row == 0) df <- as.data.frame(stats::setNames(rep(list(character(0)), n_col), headers),
                                      check.names = FALSE)
  wb$add_data(sheet = title, x = df, with_filter = TRUE, na = NULL)
  for (j in seq_along(mixed_numbers)) {
    for (i in mixed_numbers[[j]]) {
      wb$add_data(sheet = title, x = as.numeric(rows[[i]][[j]]),
                  dims = paste0(col_letter(j), i + 1), col_names = FALSE)
    }
  }
  last <- paste0(col_letter(n_col), n_row + 1)
  wb$add_cell_style(sheet = title, dims = paste0("A1:", last), wrap_text = "1", vertical = "top")
  header_dims <- paste0("A1:", col_letter(n_col), "1")
  wb$add_fill(sheet = title, dims = header_dims, color = openxlsx2::wb_color("FF1F4E78"))
  wb$add_font(sheet = title, dims = header_dims, name = "Calibri", size = "11",
              color = openxlsx2::wb_color("FFFFFFFF"), bold = "1")
  wb$set_col_widths(sheet = title, cols = seq_along(widths), widths = widths)
  # Colour each status cell by its value (green OK, amber check, red problem).
  for (sc in status_cols) {
    for (i in seq_len(n_row)) {
      v <- rows[[i]][[sc + 1]]
      colour <- if (is.character(v) && length(v) == 1 && !is.na(v)) STATUS_COLOR[v] else NA
      if (!is.na(colour)) {
        wb$add_fill(sheet = title, dims = paste0(col_letter(sc + 1), i + 1),
                    color = openxlsx2::wb_color(paste0("FF", colour)))
      }
    }
  }
  if (title != "Summary") {
    wb$freeze_pane(sheet = title, first_active_row = 2, first_active_col = 2)
  } else {
    wb$freeze_pane(sheet = title, first_active_row = 2, first_active_col = 1)
  }
  invisible(wb)
}

write_excel <- function(path, refs, claims, pdfs, run_info) {
  wb <- openxlsx2::wb_workbook()
  wb$set_base_font(font_size = 11, font_name = "Calibri")

  # (d) Summary first so it opens on top.
  write_sheet(wb, "Summary", c("Item", "Count / value"), run_info$summary_rows, c(55, 70))

  # (a) References
  ref_rows <- lapply(refs, function(r) {
    list(r$index, r$status, r$problems, r$text, py_or(r$doi, r$url),
         if (length(r$surnames)) r$surnames[1] else "", r$cr_first_author, r$check_author,
         if (is.na(r$year)) NA else r$year, r$cr_years, r$check_year,
         r$title_similarity, r$cr_title,
         r$journal, r$cr_journal, r$check_journal,
         r$cr_volume, r$check_volume,
         r$cr_pages, r$check_pages, r$cr_type,
         length(r$cited_in), r$oa_pdf,
         paste(vapply(r$pdfs, function(p) pdfs[[p]]$name, character(1)), collapse = "; "),
         r$abstract_source, r$abstract)
  })
  write_sheet(wb, "References",
              c("#", "Status", "Problems found", "Reference (as written)", "DOI / URL",
                "First author (ref)", "First author (Crossref)", "Author check",
                "Year (ref)", "Year (Crossref)", "Year check",
                "Title similarity (0-1)", "Title (Crossref)",
                "Journal/book (ref, parsed)", "Journal/book (Crossref)", "Journal check",
                "Volume (Crossref)", "Volume check", "Pages (Crossref)", "Pages check", "Crossref type",
                "Times cited in text", "Open-access PDF link", "Matched local PDF",
                "Abstract source", "Abstract"),
              ref_rows,
              c(5, 16, 40, 60, 30, 14, 14, 10, 8, 22, 14, 10, 45, 25, 25, 10, 9, 10, 12, 18, 14, 8,
                30, 25, 14, 80),
              status_cols = 1)

  # (b) Claims to verify
  claim_rows <- lapply(seq_along(claims), function(n) {
    cl <- claims[[n]]
    r <- if (is.na(cl$ref_i)) NULL else refs[[cl$ref_i]]
    ps <- cl$passages
    list(n, cl$citation, cl$location, cl$claim, truncate_text(cl$context, 1500),
         cl$match_status, cl$match_note,
         if (!is.null(r)) paste0(r$short, ": ", py_or(r$cr_title, r$title)) else "",
         if (!is.null(r)) r$status else "",
         if (!is.null(r)) truncate_text(r$abstract, ABSTRACT_CHARS_IN_CLAIMS) else "",
         if (!is.null(r)) r$abstract_source else "",
         if (!is.null(r)) r$oa_pdf else "",
         cl$pdf_name,
         if (length(ps) > 0) ps[1] else "", if (length(ps) > 1) ps[2] else "",
         if (length(ps) > 2) ps[3] else "",
         "", "", "")
  })
  write_sheet(wb, "Claims_to_verify",
              c("#", "Citation (as parsed)", "Location in document", "Claim (sentence citing it)",
                "Full paragraph / table row", "Match", "Match note", "Matched reference",
                "Reference status", "Abstract (truncated)", "Abstract source", "Open-access PDF link",
                "Local PDF", "Best PDF passage 1", "Best PDF passage 2", "Best PDF passage 3",
                "Interpretation correct? (Y/N/Partly)", "Notes", "Action"),
              claim_rows,
              c(5, 22, 32, 55, 40, 14, 30, 40, 14, 70, 12, 25, 20, 50, 50, 50, 16, 35, 30),
              status_cols = c(5, 8))
  # Highlight the reviewer's columns (yours) and offer a Y/N/Partly drop-down.
  if (length(claim_rows)) {
    wb$add_fill(sheet = "Claims_to_verify", dims = paste0("Q2:S", length(claim_rows) + 1),
                color = openxlsx2::wb_color("FFFFF2CC"))
    wb$add_data_validation(sheet = "Claims_to_verify", dims = paste0("Q2:Q", length(claim_rows) + 1),
                           type = "list", value = '"Y,N,Partly"', allow_blank = TRUE)
  }

  # (c) Unmatched
  un_rows <- list()
  for (cl in claims) {
    if (cl$match_status %in% c("NO REFERENCE", "AMBIGUOUS", "MATCHED (check note)")) {
      un_rows[[length(un_rows) + 1]] <- list(cl$match_status, cl$citation, cl$location, cl$claim, cl$match_note)
    }
  }
  for (r in refs) {
    if (length(r$cited_in) == 0) {
      un_rows[[length(un_rows) + 1]] <- list(
        "NOT CITED", r$short, "reference list", r$text,
        "listed but never cited in the text (may be fine: background/optional reading)")
    }
  }
  write_sheet(wb, "Unmatched", c("Type", "Citation / reference", "Location", "Claim or reference text", "Note"),
              un_rows, c(22, 28, 40, 70, 50), status_cols = 0)

  dir.create(dirname(path), recursive = TRUE, showWarnings = FALSE)
  wb$save(path)
}

write_markdown <- function(path, refs, claims, run_info) {
  L <- c(paste0("# Citation check: ", run_info$doc_name), "",
         paste0("Run ", run_info$when, " | network: ", run_info$network, " | PDFs: ", run_info$pdf_note),
         paste0("Full report: `", run_info$xlsx_name, "`"), "",
         "## Counts", "")
  for (row in run_info$summary_rows) {
    k <- row[[1]]
    if (nzchar(k) && !starts_with(k, "---")) L <- c(L, paste0("- ", k, ": ", row[[2]]))
  }
  L <- c(L, "")

  section <- function(L, title, rows) {
    c(L, paste0("## ", title), "", if (length(rows) == 0) "None.", rows, "")
  }
  need <- Filter(function(r) r$status != "OK", refs)
  L <- section(L, "References needing attention",
               vapply(need, function(r) paste0("- **", r$short, "** (", r$status, "): ", r$problems),
                      character(1)))
  nr <- Filter(function(cl) cl$match_status == "NO REFERENCE", claims)
  L <- section(L, "In-text citations with no reference",
               vapply(nr, function(cl) paste0("- ", cl$citation, " (", cl$location, "): ", cl$match_note),
                      character(1)))
  amb <- Filter(function(cl) cl$match_status %in% c("AMBIGUOUS", "MATCHED (check note)"), claims)
  L <- section(L, "Ambiguous or noted matches",
               vapply(amb, function(cl) paste0("- ", cl$citation, " (", cl$location, "): ", cl$match_note),
                      character(1)))
  unc <- Filter(function(r) length(r$cited_in) == 0, refs)
  L <- section(L, "References never cited in text (may be fine)",
               vapply(unc, function(r) paste0("- ", r$short), character(1)))
  L <- c(L, paste0("Reminder: this tool checks that citations exist and are described correctly. ",
                   "Whether each claim fairly represents the paper is for a human to judge in the ",
                   "Claims_to_verify sheet."))
  writeBin(charToRaw(enc2utf8(paste0(paste(L, collapse = "\n"), "\n"))), path)
}


# =============================================================================
# PART 10. The command line
# =============================================================================

HELP_TEXT <- "Check the citations in a lecture outline, slide deck, chapter draft or manuscript.

It confirms that every reference is a real paper and is described correctly
(authors, year, title, journal, volume, pages, DOI), then lines up every
in-text citation with the paper's abstract (and, if you have the PDFs, the
best-matching passages) so you can confirm you are interpreting it correctly.

Output: an Excel workbook and a short Markdown summary, saved in a folder
called citation_reports/ next to your document. Nothing in your document is
changed.
"

OPTIONS_TEXT <- "options:
  -h, --help     show this help message and exit
  --doc DOC      the document to check (.docx, .pptx, .md, .txt)
  --pdfs PDFS    folder of PDFs to search (sub-folders included). Optional.
  --no-pdfs      skip the PDF search entirely
  --refs REFS    take the reference list from THIS file instead (useful for
                 slide decks that cite the outline's references)
  --out OUT      folder for the report (default: citation_reports/ next to
                 your document)
  --no-network   do not go online; reuse answers saved in the cache from
                 earlier runs
  --refresh      ignore saved answers and look everything up again
"

EPILOG <- "examples:
  # Check a Word document, online:
  Rscript check_citations.R --doc \"My Chapter Draft.docx\"

  # Also search a folder of PDFs for the passages behind each claim:
  Rscript check_citations.R --doc lecture_outline.docx --pdfs readings/

  # Slides usually have no reference list: borrow the outline's list
  Rscript check_citations.R --doc lecture_slides.pptx --refs lecture_outline.docx

  # Rerun quickly without internet, using saved answers from earlier runs:
  Rscript check_citations.R --doc lecture_outline.docx --no-network

optional setting:
  Crossref answers faster if you include a contact email. To do that, set
  CITATION_CHECK_MAILTO before running, e.g.
    export CITATION_CHECK_MAILTO=you@university.edu
  Nothing is sent unless you set it.
"

USAGE <- paste0("usage: check_citations.R [-h] --doc DOC [--pdfs PDFS] [--no-pdfs] [--refs REFS]\n",
                "                         [--out OUT] [--no-network] [--refresh]\n")

# Read the command-line options (the words after "Rscript check_citations.R").
parse_args <- function(argv) {
  args <- list(doc = NULL, pdfs = NULL, refs = NULL, out = NULL,
               no_pdfs = FALSE, no_network = FALSE, refresh = FALSE)
  arg_error <- function(msg) {
    cat(USAGE, "check_citations.R: error: ", msg, "\n", sep = "", file = stderr())
    quit(save = "no", status = 2)
  }
  with_value <- c("--doc" = "doc", "--pdfs" = "pdfs", "--refs" = "refs", "--out" = "out")
  flags <- c("--no-pdfs" = "no_pdfs", "--no-network" = "no_network", "--refresh" = "refresh")
  i <- 1
  while (i <= length(argv)) {
    a <- argv[i]
    if (a %in% c("-h", "--help")) {
      cat(USAGE, "\n", HELP_TEXT, "\n", OPTIONS_TEXT, "\n", EPILOG, sep = "")
      quit(save = "no", status = 0)
    }
    key <- sub("=.*$", "", a)
    if (key %in% names(with_value)) {
      if (grepl("=", a, fixed = TRUE)) {
        value <- sub("^[^=]*=", "", a)
      } else {
        if (i == length(argv) || starts_with(argv[i + 1], "--")) {
          arg_error(paste0("argument ", key, ": expected one argument"))
        }
        i <- i + 1
        value <- argv[i]
      }
      args[[with_value[[key]]]] <- value
    } else if (a %in% names(flags)) {
      args[[flags[[a]]]] <- TRUE
    } else {
      arg_error(paste0("unrecognized arguments: ", a))
    }
    i <- i + 1
  }
  if (is.null(args$doc)) arg_error("the following arguments are required: --doc")
  args
}

# Stop early, with one clear message, if any package is missing.
check_packages <- function(need_pdftools = FALSE) {
  needed <- c("xml2", "httr2", "jsonlite", "openxlsx2", "stringi")
  if (need_pdftools) needed <- c(needed, "pdftools")
  missing <- needed[!vapply(needed, requireNamespace, logical(1), quietly = TRUE)]
  if (length(missing)) {
    stop_quietly(paste0(
      "This script needs these R packages, which are not installed: ",
      paste(missing, collapse = ", "), ".\n",
      "Install them once (in R or RStudio) with:\n",
      "  install.packages(c(", paste0("\"", missing, "\"", collapse = ", "), "))",
      if ("pdftools" %in% missing) "\nor run again with --no-pdfs." else ""))
  }
}

# Absolute, tidy version of a path typed on the command line (like Python's
# Path(...).expanduser().resolve()). It works for folders that do not exist yet.
resolve_path <- function(p) {
  p <- path.expand(p)
  if (!grepl("^/", p)) p <- file.path(getwd(), p)
  parts <- character(0)
  rest <- p
  while (!file.exists(rest) && dirname(rest) != rest) {
    parts <- c(basename(rest), parts)
    rest <- dirname(rest)
  }
  base <- normalizePath(rest, mustWork = FALSE)
  out <- base
  for (part in parts) {
    if (part == "..") out <- dirname(out) else if (part != ".") out <- file.path(out, part)
  }
  out
}

main <- function() {
  args <- parse_args(commandArgs(trailingOnly = TRUE))

  doc <- resolve_path(args$doc)
  pdf_dir <- if (args$no_pdfs || is.null(args$pdfs)) NULL else resolve_path(args$pdfs)
  check_packages(need_pdftools = !is.null(pdf_dir) && dir.exists(pdf_dir))

  if (!file.exists(doc)) stop_quietly(paste0("Can't find the document: ", doc))
  cat("Document: ", doc, "\n", sep = "")

  out_dir <- if (!is.null(args$out)) resolve_path(args$out) else file.path(dirname(doc), "citation_reports")
  cache_file <- file.path(out_dir, "cache", "api_cache.json")

  # 1. Read the document and find the reference list.
  blocks <- read_document(doc)
  if (!is.null(args$refs)) {
    refs_path <- resolve_path(args$refs)
    if (!file.exists(refs_path)) stop_quietly(paste0("Can't find the --refs file: ", refs_path))
    refs <- extract_references(blocks, refs_blocks = read_document(refs_path))$refs
    skip_ids <- extract_references(blocks)$skip_ids
  } else {
    found <- extract_references(blocks)
    refs <- found$refs
    skip_ids <- found$skip_ids
  }
  cat("Found ", length(refs), " references.\n", sep = "")
  if (length(refs) == 0) cat("  ! No reference list found. Add a 'References' heading, or use --refs.\n")

  # 2. Look up references.
  web <- new_web_cache(cache_file, offline = args$no_network, refresh = args$refresh)
  cat("Checking references", if (args$no_network) " (offline, cache only)" else " (online; ~2-4 s each)",
      "...\n", sep = "")
  refs <- check_references(refs, web)

  # 3-4. In-text citations and matching.
  claims <- collect_claims(blocks, skip_ids)
  matched <- match_claims(claims, refs)
  claims <- matched$claims
  refs <- matched$refs
  cat("Found ", length(claims), " in-text citations.\n", sep = "")

  # 5. PDFs.
  pdfs <- list()
  for (k in seq_along(refs)) refs[[k]]$pdfs <- integer(0)
  if (!is.null(pdf_dir) && dir.exists(pdf_dir)) {
    cat("Reading PDFs in ", pdf_dir, " ...\n", sep = "")
    pdfs <- load_pdfs(pdf_dir)
    refs <- match_pdfs(pdfs, refs)
  } else if (!is.null(args$pdfs) && !args$no_pdfs) {
    cat("  ! PDF folder not found: ", pdf_dir, "\n", sep = "")
  }
  claims <- attach_passages(claims, refs, pdfs)

  # 6. Reports.
  today <- format(Sys.Date(), "%Y-%m-%d")
  stem <- rx_replace("[^\\w\\-]+", "_", path_stem(doc))
  xlsx <- file.path(out_dir, paste0("citation_report_", stem, "_", today, ".xlsx"))
  md <- file.path(out_dir, paste0("citation_report_", stem, "_", today, ".md"))
  # Never overwrite an existing report (it may hold your notes): add _run2, _run3...
  n <- 2
  while (file.exists(xlsx)) {
    xlsx <- file.path(out_dir, paste0("citation_report_", stem, "_", today, "_run", n, ".xlsx"))
    md <- file.path(out_dir, paste0("citation_report_", stem, "_", today, "_run", n, ".md"))
    n <- n + 1
  }

  ref_status <- vapply(refs, function(r) r$status, character(1))
  match_status <- vapply(claims, function(cl) cl$match_status, character(1))
  st <- function(s) sum(ref_status == s)
  ms <- function(s) sum(match_status == s)
  uncited <- sum(vapply(refs, function(r) length(r$cited_in) == 0, logical(1)))
  matched_pdfs <- sum(vapply(pdfs, function(p) !is.na(p$ref_i), logical(1)))
  summary_rows <- list(
    list("Document", doc),
    list("Run date", today),
    list("Network", if (args$no_network) "offline (cache only)" else "online"),
    list("--- References ---", ""),
    list("References in list", length(refs)),
    list("OK", st("OK")),
    list("MINOR MISMATCH", st("MINOR MISMATCH")),
    list("MAJOR MISMATCH (probably the wrong paper)", st("MAJOR MISMATCH")),
    list("NOT FOUND (DOI does not resolve)", st("NOT FOUND")),
    list("MANUAL CHECK (no DOI)", st("MANUAL CHECK")),
    list("NOT CHECKED (network problem or not cached)", st("NOT CHECKED")),
    list("References with an abstract found", sum(vapply(refs, function(r) nzchar(r$abstract), logical(1)))),
    list("References with an open-access PDF link", sum(vapply(refs, function(r) nzchar(r$oa_pdf), logical(1)))),
    list("--- In-text citations ---", ""),
    list("In-text citations parsed (one per citation, per location)", length(claims)),
    list("Matched to a reference", ms("MATCHED") + ms("MATCHED (check note)")),
    list("  of which have a note (year off by 1, a/b letter, et al./& form)", ms("MATCHED (check note)")),
    list("No matching reference", ms("NO REFERENCE")),
    list("Ambiguous (more than one possible reference)", ms("AMBIGUOUS")),
    list("References never cited in text", uncited),
    list("--- PDFs ---", ""),
    list("PDF folder", if (!is.null(pdf_dir) && dir.exists(pdf_dir)) pdf_dir else "none"),
    list("PDFs read", length(pdfs)),
    list("PDFs matched to a reference", matched_pdfs)
  )
  run_info <- list(summary_rows = summary_rows, doc_name = basename(doc), when = today,
                   network = if (args$no_network) "offline" else "online",
                   pdf_note = paste0(length(pdfs), " read, ", matched_pdfs, " matched"),
                   xlsx_name = basename(xlsx))
  write_excel(xlsx, refs, claims, pdfs, run_info)
  write_markdown(md, refs, claims, run_info)
  if (web$new_calls > 0) web_save(web)

  cat("\n")
  for (row in summary_rows) {
    k <- row[[1]]
    if (!starts_with(k, "---") && k != "Document") cat("  ", k, ": ", row[[2]], "\n", sep = "")
  }
  cat("\nReport:  ", xlsx, "\nSummary: ", md, "\n", sep = "")
}

main()
