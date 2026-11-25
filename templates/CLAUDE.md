# CLAUDE.md for researchers (template)

<!--
HOW TO USE THIS FILE
1. Copy it to ~/.claude/CLAUDE.md (Mac/Linux) or %USERPROFILE%\.claude\CLAUDE.md (Windows).
   Claude Code reads that file at the start of every session, in every folder.
2. Fill in the parts in <angle brackets>. Delete any rule you disagree with; add your own.
3. Keep it short. Every line is read every session, so a rule that never matters is noise.
   Project-specific rules go in a CLAUDE.md inside that project's folder instead.
This comment block can be deleted once you have filled it in.
-->

## About me
- <Name, position, institution, field.>
- <What you are expert in, so Claude does not over-explain it.>
- Coding: <e.g. "intermediate R, no Python" or "comfortable in Python, new to git">.
- Explain software concepts in plain language: define terms, use an analogy, say why a step is needed.

## How to work with me
- Propose a plan before starting any multi-step task, and wait for my go-ahead.
- If anything is ambiguous, ask me clarifying questions before proceeding.
- Be direct. Tell me when I'm wrong, when an approach is fragile, or when a result looks suspicious.
- Say how confident you are in each claim (high / medium / low) and why.
- If something fails, report the actual error and what you tried. Do not quietly work around it.

## Scientific rigor
- Never invent citations, DOIs, parameter values or data. If a value is unknown, say so and mark it `TODO`.
- Citations for academic work come from a scholarly database record (e.g. OpenAlex, Crossref,
  Web of Science), not from memory. Give the DOI or URL from that record.
- When I need to check a claim, quote the source word for word and give the page or section.
- Keep three things separate and labeled: values from data, values from the literature, and
  assumptions or placeholders.
- Keep units everywhere: in variable names, column names or comments (e.g. `biomass_g_m2`).
- Report uncertainty and sample size with every summary statistic. Flag low n.
- If you don't know why something happened, say "we don't know". Label possible explanations
  as untested possibilities, never as findings.

## Data and files
- Raw data is read-only. Derived data goes in a separate folder.
- Before deleting, overwriting or mass-renaming anything, tell me what will be affected and wait for a yes.
- If a Word or Excel file is open (a `~$` lock file sits next to it), don't overwrite it. Tell me the change instead.
- Never read, print or commit `.env` files, passwords, API keys or tokens.
- <Confidential material you handle, e.g. "Student records and reviews stay in their folder and
  are never sent to an outside service.">

## Code
- Language: <R / Python / both>. <Style, e.g. "tidyverse style, ggplot2 for figures, here::here() for paths".>
- Clear, readable code over clever code. Comment it in plain language so I can check each step,
  and tell me what to look for to confirm a step worked (expected row counts, value ranges).
- Reproducible: set random seeds, no hard-coded absolute paths, record package versions
  (<renv / requirements.txt>).
- One script per logical step (`01_clean`, `02_model`, `03_figures`).
- Don't refactor or reformat code I didn't ask you to touch.

## Figures
- Publication-ready by default: readable fonts at final size, colorblind-safe palettes, labeled
  axes with units, no chart junk. <"No titles on figures" if your field puts titles in captions.>
- Save vector (PDF or SVG) plus PNG at 300 dpi.

## Writing
- <Tone, e.g. "Academic: precise, concise, active voice. No hype.">
- When editing my text, keep my meaning and voice, and show changes rather than rewriting wholesale.
- <Spelling: American or British.>

## Git
- Small, focused commits with clear messages. Never commit data files, credentials or large files.
- Never push, force-push, rebase shared branches, or delete files or branches without asking.

## Projects
- Every project has a README.md with a "Pick up here" section: read it first, and update it at the
  end of each session (date, what was done, what is next, who we're waiting on).
