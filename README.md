# Research skills for Claude Code

Small, practical add-ons ("skills") for [Claude Code](https://claude.com/claude-code) that take
recurring, fiddly parts of academic work off your plate: setting up a project, checking citations,
planning a talk. They are written for faculty, postdocs and graduate students. **You do not need
to know how to code to use them.**

Made by Kris Lewis, PhD (marine ecosystem modeler and faculty member) as part of
**Research, Reclaimed**. Free to use under the MIT license.

---

## What is a skill?

A skill is a folder holding a plain-text instruction file (`SKILL.md`), and sometimes a small script.
Once a skill is installed, Claude Code follows its steps whenever you ask for that task, or when you type
its name with a slash (for example `/citation-check`). Think of it as a written lab protocol that
Claude follows the same way every time, instead of improvising.

You can open any `SKILL.md` in a text editor and read exactly what it will do. Nothing is hidden.

---

## The skills

| Skill | In one line |
|---|---|
| [`new-project`](#new-project) | Set up every new paper, analysis or course the same way, with a notebook README and private GitHub backup |
| [`citation-check`](#citation-check) | Confirm every reference in a document is real and correctly described, and line up each claim with what the paper says |
| [`talk-story-planner`](#talk-story-planner) | Find the story for a seminar or job talk before you make a single slide |
| [`templates/CLAUDE.md`](#claudemd-template) | A starter set of standing rules for researchers (never invent citations, raw data is read-only, ...) |

Coming next: a purchase-card (P-Card) reconciliation skill and a prospective-student email triage skill.

---

### new-project

**Who it's for:** anyone who starts more than one project a year and has ever opened an old folder
and thought "where was I?"

**When you'd use it:** You've just agreed to write a paper with a collaborator, or a student is
starting a thesis chapter. You type `/new-project` and answer four questions (what kind of project,
its name, new or existing folder, anything sensitive in it).

**What it does:**
- Makes the same folder layout every time: raw data kept apart from derived data, one script per step.
- Writes a **README that works as the project notebook**, with a "Pick up here" section at the top.
  Each session ends by noting what was done and what's next, so you, a student or Claude can pick
  the project up months later in seconds.
- Writes a `.gitignore`: a list of files git must never save, such as data, passwords, PDFs and
  confidential files.
- Makes a first save point in git and a **private** backup on GitHub.

**What you still do:** answer the questions, check the list of files before the first save, and
sign in to GitHub once.

**What you need:** Claude Code, git, and a free GitHub account with the
[GitHub command-line tool](https://cli.github.com) (the skill walks you through it).

**Limits:** it never makes a repository public, and it won't move or delete anything in an existing folder.

---

### citation-check

**Who it's for:** anyone who cites papers: lecture outlines, slide decks, thesis chapters,
manuscripts, grant proposals. It's especially useful when AI tools helped with the reading or
drafting, because AI tools sometimes produce references that look real but aren't.

**When you'd use it:** It's the night before a lecture and your outline has 14 citations, some
copied from old slides. Or a student hands you a chapter and you want to know whether the
references hold up before you read it closely. You type `/citation-check` and give it the file.

**What it does:**
1. Looks up every DOI in the reference list on Crossref (the official DOI registry). It checks the
   authors, year, title, journal, volume and pages against the publisher's record, field by field.
2. Finds every in-text citation, such as "(Koch 2001)" or "Nixon et al. (2009)", and matches it to
   the reference list. It flags citations with no reference and references never cited.
3. Builds an Excel sheet with one row per cited claim. Each row holds your sentence, the paper's
   abstract, and (if you give it a folder of the PDFs) the three passages in the paper that best
   match, with page numbers. You record Y / N / Partly for "does the paper actually say this?".

**What you still do:** judge each claim. The tool lines up the evidence; whether a sentence is a
fair reading of the paper is your call. References without a DOI (reports, some books) are left
for you to check by hand.

**What you need:** Claude Code, and either **R** or **Python** (both versions are included; use
whichever you know). Internet access for the lookups. Your document must be Word, PowerPoint,
Markdown or plain text.

**Limits:** it reads author-year citations, not numbered ones like [1]. Abstracts often leave out the exact
number you cite, so then you read the paper. Scanned PDFs need OCR (text recognition) first. It never edits your document.

**Try it:** `skills/citation-check/example/sample_lecture_notes.md` has two planted mistakes. The
check finds both in about a minute.

---

### talk-story-planner

**Who it's for:** anyone giving a seminar, job talk, conference talk or guest lecture who has
plenty of material but no clear story yet.

**When you'd use it:** You've been invited to give a departmental seminar in three weeks. You have
papers, figures and an old slide deck, but the old structure feels dated. You type
`/talk-story-planner`.

**What it does:**
1. Interviews you **one question at a time**: who's in the room, the one idea they should remember
   a week later, what you want to happen after the talk, and which results are finished and which
   are in progress.
2. Proposes **2-3 genuinely different story options** (different opening, different main message),
   says what each would cost you in new figures or analysis, and recommends one.
3. Waits for you to choose, then writes a lean, timed outline. Each part is tied to a specific
   paper or figure in your folder, and gaps are marked `[TODO]`.

**What you still do:** choose the story and make the slides. The skill deliberately stops before slides.

**Rules it follows:** it cites only papers you've put in the talk's `papers/` folder, marks
anything else `[VERIFY]`, and never invents numbers. If you start going in circles, it narrows the
choices instead of adding more.

**What you need:** Claude Code and a folder with your papers (PDFs).

---

### CLAUDE.md template

**Who it's for:** anyone setting up Claude Code for research for the first time.

**What it is:** Claude Code reads a file called `CLAUDE.md` at the start of every session. It holds
your standing instructions, like a lab's rules posted by the door. This template is a starting
point with rules researchers usually want:
- Never invent citations, DOIs or data. Mark unknowns as `TODO`.
- Keep values from data, from literature and from assumptions labeled separately.
- Raw data is read-only.
- Ask before deleting or overwriting anything.
- Say how confident you are, and report errors instead of working around them quietly.

It also has blanks for your field, coding language and writing style.

**How to use it:** copy `templates/CLAUDE.md` to `~/.claude/CLAUDE.md` (Mac) or
`%USERPROFILE%\.claude\CLAUDE.md` (Windows), then fill in the parts in `<angle brackets>` and delete
anything you disagree with.

---

## Install

### Mac
1. Download this repository: green **Code** button, then **Download ZIP**. Unzip it.
2. In Finder, choose **Go > Go to Folder...**, type `~/.claude/skills` and press Return. If there is
   no `skills` folder, open `~/.claude` and create one.
3. Copy the skill folders you want from the download's `skills/` folder into it. You should end up
   with, for example, `~/.claude/skills/citation-check/SKILL.md`.
4. Start Claude Code and type `/` followed by the skill's name.

### Windows
1. Download this repository: green **Code** button, then **Download ZIP**. Right-click the ZIP and
   choose **Extract All**.
2. In File Explorer, click the address bar, type `%USERPROFILE%\.claude` and press Enter. Open the
   `skills` folder (create it if it isn't there).
3. Copy the skill folders you want into it, so you have, for example, `.claude\skills\citation-check\SKILL.md`.
4. Start Claude Code and type `/` followed by the skill's name.

The skills were built and tested on a Mac. If a step fails on Windows, paste the error into Claude
and ask it to fix the skill for your computer.

### Or let Claude install it
Unzip the download, start Claude Code, and say: *"Install the citation-check skill from the folder I
just unzipped in Downloads. Show me each step before you do it."*

---

## Changing a skill to fit how you work

Nothing in a skill is locked. To change one, describe the change to Claude in plain words, for example:

- "Edit my new-project skill so manuscripts also get a `data-availability/` folder."
- "Edit my talk-story-planner skill: my talks are always 15 minutes, so skip the length question."

Good habits:
1. Ask Claude to **show you the edit before making it**.
2. Ask it to **keep a copy of the old version** first (`SKILL-backup.md`).
3. Start a new Claude session afterwards so it picks up the change.

---

## Privacy

- The skills run on your own computer. Nothing is sent to the author.
- `citation-check` sends only DOIs to public scholarly databases (Crossref, OpenAlex, Semantic
  Scholar). It does not send your document. It adds a contact email only if you set one yourself.
- If your work involves student records, human-subject data or unpublished collaborator data,
  check your institution's rules on AI tools before using any AI assistant with it.

## License

MIT: use, change and share freely, including commercially, as long as you keep the copyright
notice. See [LICENSE](LICENSE).

Questions or improvements: open an issue on this repository.
