---
name: talk-story-planner
description: Find the story for a seminar, conference talk, job talk or guest lecture before any slides are made - interviews the speaker one question at a time, proposes 2-3 different story options with a recommendation, waits for a choice, then drafts a lean outline that cites only papers the speaker has supplied. Use when the user says they have a talk coming up, has "lots of material but no story", wants help planning or outlining a talk or seminar, or runs /talk-story-planner.
---

# Find the story for a talk, then outline it

Planning comes before slides. **Do not make slides in this skill.** The goal is a story the
speaker believes in and an outline they can build from.

Write in plain language. Assume the speaker is an expert in their field, not in presentation
design or software.

## Step 1. Set up the folder (once)
Ask where the talk's folder is, or suggest one, with:
- `papers/` - the papers the talk draws on, plus `papers/README.md` listing each in a line or two
  (help write this list if it is missing; read each PDF's first page for the line).
- `old-drafts/` - earlier outlines or slides. Treat these as raw material and a record of earlier
  thinking, **not** as the plan.
- `results/` - figures and outputs that might be shown.
- `talk-notes.md` - the facts below, written as the interview goes.

## Step 2. Interview, one question at a time
Ask one question, wait for the answer, then ask the next. Stop when you can answer all of these:
1. Title, length (minutes), date, and where (which department or meeting).
2. **Who is in the room?** Their field, how much they know about the speaker's methods, and what
   they care about (data? applications? the speaker as a future colleague?).
3. **The one idea** a listener should remember a week later.
4. What the speaker wants to happen after the talk (collaborators, a job offer, students, funding,
   people trying a method).
5. The strongest evidence the speaker already has, and which results are finished vs. in progress.
6. Hard limits: things that cannot be shown (unpublished data, confidential partners), and how
   much prep time is left.

Record the answers in `talk-notes.md` as you go.

## Step 3. Propose 2-3 story options
Each option is different in its **opening** and its **main message**, not just in order. For each:
- One-sentence main message.
- How it opens (a question, a puzzle, a picture, the audience's problem, a surprising number).
- The 3-5 beats of the middle, each tied to a specific result or paper from the folder.
- How it ends (the take-home and the ask; never a bare "Thank you" slide).
- What it costs: new analysis or figures needed, and the risk.

Then say which you would pick and why, in two or three sentences. **Wait for the speaker to choose.**

## Step 4. Outline the chosen story
A lean outline: one line per slide or section, with time per section adding up to the talk length
minus about 10% for questions. Under each line: the evidence it uses (paper or figure file) and
anything still missing, marked `[TODO]`.

Show unfinished work honestly: label it "in progress" or "preliminary", say what has been done and
what is next, and say what would change the conclusion.

## Rules for sources and numbers
- Only cite papers in `papers/`. Anything else is marked `[VERIFY]` and the speaker checks it.
- When a number or claim comes from a paper, name the paper.
- Never reconstruct a paper's argument from memory. If a detail is needed, open the file.
- Never invent numbers, citations or results. If unsure, mark `[VERIFY]` and say so.

## When the speaker is overwhelmed
If they are going in circles or adding ideas faster than deciding, narrow the choices instead of
adding new ones: "Of these, which one would you be sorry to cut?" Offer to park ideas in a
"not this talk" list in `talk-notes.md`.
