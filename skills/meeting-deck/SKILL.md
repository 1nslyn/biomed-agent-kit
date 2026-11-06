---
name: meeting-deck
description: Build, update or audit a research lab group meeting deck as a .pptx, with charts, speaker notes and a checked source for every number. Use whenever someone asks for meeting slides, a weekly, bi-weekly or Friday deck, 组会 PPT, 开会的 slides or 汇报, asks to put a result on the deck, or asks to refresh, reorder or check a meeting deck, even if they never mention a skill. Not for paper figures or conference talks.
---

# Meeting deck

Build the deck a lab member presents at the group meeting. The audience reads it slide by slide, so a number the presenter cannot explain or a status that is out of date costs them directly.

These rules came out of repeated reviews of real lab decks. The style rules are house defaults, and anything in `settings.md` overrides them.

Plan the whole scope once and finish in one pass. Pick defaults for open questions and list them at the end. Stop only at a real blocker, or at a change to what the last meeting agreed (step 2).

## Settings

Read `settings.md` next to this file. If it does not exist, ask the presenter for the fields in `settings.example.md` in one message, then write their answers to `settings.md`. If the skill folder is read-only, keep the answers for this session.

When the presenter corrects a deck, add the correction to "My rules" in `settings.md`, so it applies next time and survives updates to this skill.

The scripts need the packages in `requirements.txt`; rendering needs LibreOffice (`soffice`) and Poppler (`pdftoppm`). Install what is missing.

## 1. Set the window

The deck covers only what is new since the last meeting. Take this meeting's date from the presenter's request; the window runs from the last meeting's date to this one.

- Find that meeting: the newest dated folder in the deck folder, the newest meeting deck the presenter has, and that meeting's note. With Granola, search by date range with no keywords, since titles vary and the list mixes in unrelated meetings. If there is no note, ask the presenter to paste it.
- Read the previous deck and the note. Write down what was presented, each issue raised and each next step.

## 2. Collect the evidence

- **Scope:** only the projects in the settings, from the sources named there. Other affiliations and personal work stay out unless the presenter asks.
- **Next steps:** report each one from the last meeting. Done items go on the deck; the rest go in Todos.
- **Issues raised at the last meeting:** put the earlier curve or result next to the result after the fix, so the audience can see the comparison.
- **Live state:** read running jobs, experiment trackers, ledgers and file times now. Notes and earlier messages go out of date within a day. Check other session transcripts under `~/.claude/projects/` for decisions made elsewhere. When a server needs an interactive login and cannot be read, name the last report and its time on the slide.
- **Look at it yourself:** open per-fold files, plot them, view images and low-score cases, and compare sources. Say in the notes where they disagree. A subagent's or a log's claim is evidence until you open the file behind it.
- **Meeting plan:** compare what was run with what the meeting agreed (splits, cohorts, methods). Name any departure as a change on the slide and get the presenter's yes before relying on it. If they cannot answer, show the result with the departure flagged on the slide and list it in your report.
- **Project goal:** when a result falls short of a project's goal, name the next attempt with it. Validation gains are not results, so show held-out test against baseline as well. A kept baseline is zero gain.

## 3. Outline

1. **Title slide:** `Weekly Meeting November 4th`, the presenter's name, `Oct 30 – Nov 4, 2026`. A deck with one theme may use a descriptive title instead.
2. **Summary slide, results first:** columns such as `Experiment | Trained on | Tested on | Result`, in full plain sentences.
3. **Results**, then why and how, then other progress.
4. **Todos** by project, in imperative form (`ds.todos`). Open items that need a decision go just before it.
5. **Backup** after an index slide: all curves, edge cases, supplementary checks.

A status with no chart or table stays in the summary and Todos, so it does not leave a nearly empty slide. When asked to strengthen a slide, add to it or add a slide, and keep what was there.

## 4. Build

Work in `<deck folder>/<YYYY-MM-DD>/`. Copy `scripts/deck_style.py` and `scripts/chart_style.py` from this skill into that folder, so an old deck still rebuilds unchanged after the style moves on. Keep there: `build_deck.py`, `plot_deck.py`, `data/`, `figs/deck/`, `number_audit.tsv` and a `README.md` listing slides, files and how to rebuild. `examples/` holds a small deck that runs.

Before rebuilding over an existing deck, check for a `~$` lock file and whether the .pptx is newer than the last build. Presenters edit and save decks themselves, and a rebuild would erase their changes. Write to a new file name and carry their edits into the builder.

**Use python-pptx with `deck_style.py`.** Decks written by pptxgenjs opened with PowerPoint's Repair prompt, so avoid it. Never deep-copy slide XML: LibreOffice renders the result and PowerPoint refuses it. To add to a deck built elsewhere, call `add_slide` and copy geometry and fonts from a template slide. Editing text in place and reordering `<p:sldId>` entries are safe; renumber the page numbers after a reorder.

```python
import deck_style as ds
prs = ds.new_deck()
ds.title_slide(prs, "Weekly Meeting November 13th", ["A. Presenter", "Nov 6 – 13, 2026"], notes)
s = ds.page(prs, "Project A: test AUC, before and after", "Model C, trained on cohort A, tested on the same 80 samples: 69.2% to **72.8%**.", notes)
ds.figure(s, "figs/deck/before_after.png")    # or ds.table(s, rows, x, y, widths, strong=(3,))
ds.caption(s, ["Test AUC, mean of 5 seeds."])
prs.save(out)
```

`**x**` marks a result number. Also available: `text`, `fit`, `box`, `band`, `node`, `arrow`, `todos`. Draw schematics from native shapes so the presenter can edit them.

**Charts:** draw each at `figsize=(CHART_W, h)` and place it with `ds.figure`, so its text keeps its point size. Use one axis and thin marks. Read every value from a stored file, since typed-in chart rows are where errors collect. `chart_style.py` explains how to run a plot script locally or on a server.

## 5. Slide rules

**Layout**
- Times New Roman on white, 13.33 x 7.5 in. Title 28 pt bold, takeaway 18 pt, table 15 pt, caption 14 pt, chart text 12 pt or larger.
- Each slide has a title, one takeaway line in plain dark text, then a chart or table, then at most a short caption. Sentences beyond that go to the speaker notes, which every slide has.
- Numbers stay on the slide. The audience must be able to read and derive each reported value there; the notes hold explanation and procedure.
- Content flows from the top with spare room at the bottom, never a gap in the middle.
- Tables have rules above and below the header and at the end, with at most one tinted key row.
- No cards, shadows, gradients, color blocks, badges or status colors. No red or orange text. Each color keeps one meaning across the deck, with a legend for 2 or more series and direct labels where they fit.
- Real charts and images are the visuals. Add an icon only when it helps, and it must depict its label literally.

**Titles**
- `<Project>: <what this slide shows>`, using only projects actually worked on.
- Titles and takeaways state the capability and the result. Safeguards such as validation-only selection appear once as plain facts in a method line, never as a headline.
- Never name the people in the room, in slides, notes or file paths, including as the source of a task. Name the work by its content or date: "the 11-04 to-do list". Headings describe content, never the reader.

**Numbers**
- Percentages for every metric level, gain, SE and noise figure: `+3.4%`, `60% AUC`. A gain is the difference in percentage points: 81.2% to 83.4% is `+2.2%`. Correlations, SE multipliers, counts and p values keep their own form. If a value looks wrong, confirm it from the source before converting.
- Each number says its experiment: trained on what, tested on what with how many samples, which model and input. Read each number alone and ask which experiment it belongs to.
- Training details such as early stopping, precision and GPU model stay in the notes unless a claim depends on them.
- A change shows both sides: `Model C, trained on cohort A vs A + B, same 80 test samples: 69.2% to 72.8%`. A margin shows its two numbers: `gained 2.1%, needed 3.0%`.
- Compare numbers only on the same samples. When two slides count differently, say how on the slide.
- A comparison says in words what differs between the rows and what stays the same. Before and after labels say what changed.
- A range says what varies, or names the best and worst settings with their values.
- Result slides show levels as well as gains, with a chart of the per-model or per-fold numbers and their spread, or a full exact table.
- Name the quantity each time: 5-seed ensemble vs mean of 5 seeds, last vs best epoch, one CV vs 5 repeats. Epochs are 1-based. A "fewest" or "most" names its pool, checked against the code that picked it. A control value that moves with the seed is not a fixed bar.

**Words**
- Each slide sentence makes sense alone on first read. Say the plain action: `tested on`, `separates 40x from 20x`. Name the thing: `the 40x vs 20x AUC`, never `the third column`.
- Write `trained on cohort A only, lr 2e-4` and `settings from a grid search`, never `A-trained` or `tuned`.
- Use standard names (5-fold, fold model, validation, held-out test, baseline), the dataset's own name from its config and the code's own names. Never coin a name, because the audience has no way to know what it means. A re-shuffle of the same samples is never "new data".
- Internal codes and statistics terms (rank, SE, cross-fitting) stay off slides. Describe them in words, and keep the codes in the notes if needed.
- Before a result slide uses a procedure, show it once in plain words or a small picture. Give a term of art a short footnote where it first matters.
- Status and Todo lines are short and concrete: `Finished the...`, `Currently running...`.
- No metaphors, no em dashes, no "X, not Y" framing, no vague qualifiers.

## 6. Audit every number

When one 40-slide deck was audited, 7 of its 939 claims were wrong and 40 needed a clearer label, nearly all from typed-in values or one word covering two quantities.

- Start the claim list with `python3 <this skill>/scripts/check_deck.py deck.pptx --dump-numbers`. It does not see numbers drawn inside chart images, so add those from the plot data.
- Recompute each claim from its stored source. Record it in `number_audit.tsv` with columns `slide location claim stated recomputed source status resolution`. For a large deck, run two read-only subagents independently and settle every disagreement yourself.
- A number with no stored source is removed or labelled by where it came from (`10-30 notes, scores not kept`).
- Refresh every time-bound slide right before export: run counts, "still running", job status, snapshot times. Flagging a slide as out of date does not finish it.

## 7. Check

1. `python3 <this skill>/scripts/check_deck.py deck.pptx --people "Ann,Bob"`, with the names from the settings. Clear every error. Warnings mark departures from the house style, so fix each one unless the settings allow it.
2. Diff against the previous deck and the last note. Cut anything already presented, from slides and notes. An old number stays only when a new result cannot be read without it; tell the presenter which ones stayed.
3. Render and read every slide image:
   ```bash
   soffice --headless --convert-to pdf --outdir render deck.pptx
   ```
   ```bash
   pdftoppm -r 80 -png render/deck.pdf render/s
   ```
   Look for wrapped cells such as `Tumo r`, overlaps, tiny chart text and legends that hide data. LibreOffice proves how a slide looks; it does not prove PowerPoint will open the file.
4. On a Mac with PowerPoint, `<this skill>/scripts/open_in_powerpoint.sh deck.pptx` opens a copy, exports a PDF and closes only that copy. A timeout usually means a Repair prompt. Run it last, since it brings PowerPoint forward on the screen. Elsewhere, ask the presenter to open the file once.

## 8. Deliver

Deliver a .pptx, since the presenter needs a file to open, edit and present; a web page or slides artifact does not replace it. Put it in `~/Downloads`, or send it if the session has a file tool. Report the slide count, defaults you picked, old numbers kept and anything you could not verify.
