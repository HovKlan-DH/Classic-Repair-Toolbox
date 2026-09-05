# The Workbooks tab

← [Back to Workbooks](Workbooks)

The Workbooks tab is the full browser for your repair jobs. It appears when **"Enable Workbooks
tab"** is ticked in Configuration.

It is laid out in three columns, with a header spanning the right-hand two:

```
┌───────────────┬───────────────────────────────────────────────────────────────┐
│ Find a repair │ #3 · Dead C64, no video                          [🔓 Open]    │
│ [__________]  │ Reported as working until a thunderstorm.                     │
│               │                     [Edit workbook] [Delete workbook]         │
│ 3 workbooks   │                     [Export to PDF] [Export to ZIP]           │
│               │ ▸ 3 worklogs · 2.0 h · 160 · 1 open                           │
│ ┌───────────┐ ├──────────────────────────────┬────────────────────────────────┤
│ │#3 Dead C64│ │  ┌────────────┐ ┌──────────┐ │  Worklogs on "Video"           │
│ │  🔓 Open  │ │  │  Power     │ │  Video   │ │  ┌──────────────────────────┐  │
│ └───────────┘ │  │  [#1]      │ │  [#2]▓   │ │  │ ②  VIC socket cracked  ✕ │  │
│ ┌───────────┐ │  └────────────┘ └──────────┘ │  │ Pin 8 lifted from the …  │  │
│ │#2 Tape …  │ │                              │  │ [Issue] [🔒 Closed]      │  │
│ └───────────┘ │                              │  │ 1.5 h · 120 · 2 photos   │  │
└───────────────┴──────────────────────────────┴────────────────────────────────┘
     ①                          ②                            ③
```

① the **workbook list** · ② the **board pane** · ③ the **worklog list**

Both dividers can be dragged, and both remember where you put them.

---

## ① The workbook list

One card per workbook, **newest first**. Above them, a count — "3 workbooks" — which doubles as the
result count when a search is active.

Each card shows the workbook's id and description, its status pill, its start date and how many
worklogs it holds. In "Show all workbooks" scope it also names the board it belongs to.

**Clicking a card activates that workbook** app-wide — the worklog bar, "Show worklogs" and "Add
worklog" all follow it, and the choice is saved for that board. The tab does *not* switch away: the
header and the board pane update in place, which is the point of doing it here.

If the board has no workbooks yet, the panel says so. If a *search* is what emptied it, it says
"No workbooks or worklogs match your search" instead — a board with no data and a search with no
results are different situations, and saying "none recorded" for the second reads as data loss.

### Scope: showing other boards' workbooks

**Configuration → Workbooks → "Scope in Workbooks tab"** decides what this list holds:

| Setting | Behaviour |
| --- | --- |
| **Show only workbooks for selected board** (default) | Just this board's jobs |
| **Show all workbooks** | Every workbook on every board; each card names its board |

In "Show all workbooks" scope you can click a card belonging to a **different** board. The app
switches to that board first, then activates the workbook — the same jump the worklog bar's picker
performs.

The board pane can only ever draw the board that is actually loaded, so with nothing explicitly
clicked the header and pane stay on the current board's own workbook.

---

## The header

Spanning the board pane and the worklog list, describing the **selected workbook**:

**Line 1** — `#3 · Dead C64, no video` and the status pill.

**Line 2** — the workbook's **Note**. The whole line is hidden when there is no note, rather than
leaving a blank gap.

**Right-aligned, two rows of buttons:**

| Button | Does |
| --- | --- |
| **Edit workbook** | Reopens the create dialog, pre-filled, to change the description and note. Nothing else about the workbook is touched |
| **Delete workbook** | Permanently deletes it, after a confirmation — see [Deleting things](Workbooks-Deleting-things) |
| **Export to PDF** | The customer-facing document — see [Exporting](Workbooks-Exporting-a-workbook) |
| **Export to ZIP** | That PDF plus every original photo and file |

All four are hidden when no workbook is selected.

Below them sits the collapsible **[summary strip](Workbooks-The-summary-strip)** — totals for the whole
workbook.

---

## ② The board pane

Every schematic in the board that has at least one worklog in the **selected** workbook, drawn as a
preview with a thin outline. Schematics with no worklogs are not shown.

Each worklog appears on its preview exactly as it appears on the real Schematics tab:

- **"Show marked area" on** — a dashed, category-coloured rectangle plus its **"#N" pill** anchored
  to that area.
- **"Show marked area" off** — no rectangle, and the pill **parked in the preview's own top-right
  corner**, stacking neatly with any others.

**Every pill is clickable** and opens the same [worklog editor](Workbooks-The-worklog-editor) that the
Schematics tab's pills open — the component checklist included.

**Clicking anywhere else on a preview selects that schematic**: it takes a highlighted border, and
the worklog list on the right switches to that schematic's worklogs. The pane defaults to the first
schematic with worklogs, and keeps your selection across a refresh.

If the workbook has no worklogs anywhere, the pane says *"No worklogs recorded yet for any
schematics in this board"*.

---

## ③ The worklog list

One detail card per worklog on the **selected schematic**. Each card is a single bordered panel with
four rows:

1. **`#N` badge and title** — the badge is filled in the worklog's category colour.
2. **Description**.
3. **A category chip and a status pill** — outlined, in their own colours. They are informational
   here, not selectable: this list has no "chosen" concept the way the editor does.
4. **A stats row** — total hours and cost across the worklog's Work-done rows, plus how many
   comments, links, photos and files it carries. This exists so a workbook's worth of pills tells
   you how much is behind each one without opening them.

**Clicking a card opens the worklog editor** — the same window its pill opens. Hovering outlines the
card in the same accent a selected schematic preview uses.

**Each card carries a "Delete worklog" button in its top-right corner.** It is a real button, so a
click on it never also opens the editor underneath. See [Deleting things](Workbooks-Deleting-things).

---

## Finding a repair

The **"Find a previous repair"** box above the workbook list filters the entire tab as you type —
the workbook list, the board pane and the worklog list all narrow together, and matched text is
highlighted wherever it is drawn.

The query understands quoted phrases, implicit AND and `-`exclusions. See
[Searching](Workbooks-Searching) for the full grammar and the list of fields it looks at.

Searching **never** changes which workbook is active. It narrows what you see; it does not redirect
where new work is written. The query is cleared automatically when you switch boards, since a board
change is a change of subject.

---

## What keeps it up to date

The tab rebuilds itself whenever anything changes — a board change, a workbook created, edited,
activated or deleted, a worklog saved or deleted, an attachment filed — through the same single
refresh path the worklog bar and the Schematics overlay use. It also rebuilds when you switch back
to it.

---

**Next:** [The worklog editor](Workbooks-The-worklog-editor) · [Searching](Workbooks-Searching) · [The summary strip](Workbooks-The-summary-strip)
