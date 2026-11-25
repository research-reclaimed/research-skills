---
name: new-project
description: Start a new research project (analysis, manuscript, thesis chapter, grant or course) the same way every time - folder layout, a README that works as the project notebook with a "Pick up here" section, a short CLAUDE.md, a .gitignore that keeps data and private files out of git, a first commit, and a private GitHub repository. Also puts an existing folder under git. Use when the user says they are starting a new project, paper, chapter, analysis, grant or course, asks to "set up a repo", "put this under git", "make a GitHub repo", or runs /new-project (optionally with a name, e.g. /new-project eelgrass-temp-2027).
---

# Start a project the same way every time

Every project gets one folder, a README that works as the project notebook, a short CLAUDE.md, a
.gitignore that keeps data and private material out of git, a first commit, and a private GitHub
repository. Explain each step in plain language; the user may be new to git. Define a term the
first time you use it (for example: "a repository, or repo, is a folder whose history git keeps").

## Step 1. Ask (one question set, up to 4 questions)
1. **Type:** analysis or code / manuscript or thesis chapter / grant, course or admin.
2. **Name:** short, lowercase, hyphens (e.g. `eelgrass-temp-2027`). Also the GitHub repository name.
3. **New or existing folder?** If existing, get the path.
4. **Anything sensitive?** Human-subject data, student records, unpublished collaborator data, data
   under a use agreement, personnel, client or financial files. If yes, use an allow-list
   .gitignore (Step 4) and say why.

Where it lives: default `~/Projects/<name>`. If the user picks a folder inside Dropbox, OneDrive,
Google Drive or iCloud, say once that git's FAQ warns cloud syncing can damage a repository
(https://git-scm.com/docs/gitfaq) and suggest a plain folder instead. The user decides; many people
use a synced folder without trouble if they avoid editing the same repo on two computers at once.

## Step 2. Check before creating
- New folder: confirm the path does not already exist.
- Existing folder: list its top level, flag files over 50 MB (`find . -size +50M`), Office lock files
  (`~$*`) and anything that looks sensitive. Run `git -C <path> rev-parse` to see whether it is
  already a repository. Do not move, rename or delete anything.
- Never open `.env` files, keys, passwords or tokens.

## Step 3. Create the layout (only what does not exist yet)

**Analysis or code**
```
data/raw/        original data; never edited; not in git (the README says where it came from)
data/derived/    made by scripts; not in git
R/  or  python/  01_clean, 02_model, 03_figures ... one script per step
output/figures/  made by scripts
docs/            notes, methods
```
For R projects, mention `here::here()` for file paths and `renv` for recording package versions.
Do not run `renv::init()` without asking.

**Manuscript or thesis chapter**
```
manuscript/      dated drafts, e.g. <Lastname>-chapter2-<DDMonYYYY>.docx (or .qmd)
figures/         figure scripts and exported figures
literature/      PDFs (not in git: copyright)
reviews/         advisor or reviewer comments and responses
submission/      what was actually submitted, by date
notes/           meeting notes, decisions
```
Git can store Word files but cannot show changes inside them or merge two people's edits. Edit one
copy at a time, and commit at milestones ("Sent draft to advisor 03Mar").

**Grant, course or admin**: `planning/`, `drafts/`, `build/` (scripts that make documents),
`submitted/` (or `content/` and `site/` for a course website), `private/` (not in git).

## Step 4. Write three files

**README.md** (the project notebook; the most important file)
```
# <Project title>

## Pick up here
<Date>: what was done last, what is next, what is waiting on whom.

## What this is
One paragraph: the question, collaborators, deadline.

## Layout
What each folder holds.

## How to rebuild
Exact commands to rerun the analysis or rebuild the documents.

## Data and sources
Where each input came from (URL, DOI, person, date). Raw data is read-only.

## Decisions log
Dated, one line each: what was decided and why.
```

**CLAUDE.md** (short; only what a new Claude session cannot work out for itself)
```
# <Project title>
- Start from README.md "Pick up here".
- <Project rules, e.g. "R, tidyverse, here::here() for paths">
- Never commit data files, passwords or large files.
```

**.gitignore**
```
.DS_Store
~$*
.Rhistory
.Rproj.user/
__pycache__/
.env
*.key
*.pem
private/
```
Add by type: analysis `data/raw/`, `data/derived/`, `renv/library/`; manuscript `literature/*.pdf`.
For a sensitive project, use an allow-list instead: ignore everything (`/*`), then allow folders one
at a time (`!/README.md`, `!/R/` ...), with a comment saying why. An allow-list means a new folder
stays out of git until someone deliberately adds it.

## Step 5. First commit (show before saving)
1. `git init -b main` (skip if already a repository). Put an empty `.gitkeep` file in each empty
   folder that is not ignored, because git does not save empty folders.
2. `git add -A`, then `git status --short`. Show the user the list of files about to be saved.
   Stop and ask if anything looks like data, a password, a file over 50 MB or sensitive material.
3. Commit with the message `Start <name>: README, CLAUDE.md, .gitignore, folder layout`.

## Step 6. Private GitHub repository (ask first)
- Check `gh auth status`. If the GitHub command-line tool is missing, point to
  https://cli.github.com. If not signed in, ask the user to run `gh auth login` themselves
  (GitHub.com, HTTPS, log in with a web browser).
- Ask which GitHub account or organization should own it. Confirm the name, then:
  `gh repo create <owner>/<name> --private --source . --remote origin --push`
- Always private. Never make a repository public, push to someone else's repository, or
  force-push unless the user says so. (In R, `usethis::use_github()` makes a PUBLIC repository
  unless given `private = TRUE`.)

## Step 7. Finish
- Give the repository link and the page for adding a collaborator (advisor, co-author):
  `https://github.com/<owner>/<name>/settings/access`
- Fill in "Pick up here" with today's date and the next step.
- Tell the user the end-of-session routine once: update "Pick up here", commit with a clear
  message, push. A new session then starts by reading the README.
