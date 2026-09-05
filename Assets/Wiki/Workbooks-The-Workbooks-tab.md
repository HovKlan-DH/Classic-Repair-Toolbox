# The Workbooks tab

Where you browse your repair jobs.

[Wiki Home](Home) · [Workbooks](Workbooks)

---

```
┌───────────────┬──────────────────────────────────────────────┐
│ Find a repair │ #3 · Dead C64, no video           [Open]     │
│ [__________]  │ Reported as dead after a thunderstorm        │
│               │      [Edit workbook] [Delete workbook]       │
│ 3 workbooks   │      [Export to PDF] [Export to ZIP]         │
│               │ ▸ 3 worklogs · 2.0 h · 160 · 1 open          │
│ ┌───────────┐ ├──────────────────┬───────────────────────────┤
│ │#3 Dead C64│ │ ┌──────┐ ┌─────┐ │ Worklogs on "Video"       │
│ │  Open     │ │ │Power │ │Video│ │ ┌───────────────────────┐ │
│ └───────────┘ │ │ [#1] │ │[#2] │ │ │ ② VIC socket      ✕  │ │
│ ┌───────────┐ │ └──────┘ └─────┘ │ │ Pin 8 lifted          │ │
│ │#2 Tape …  │ │                  │ │ [Issue] [Closed]      │ │
│ └───────────┘ │                  │ │ 1.5 h · 120 · 2 photos│ │
└───────────────┴──────────────────┴───────────────────────────┘
      ①                  ②                      ③
```

## ① Workbook list

One card per job, newest first. **Click a card** to switch to that workbook - the bar, the board and "Add worklog" all follow it.

Which workbooks are listed depends on the scope setting in "Configuration": this board only (default), or all boards. With all boards, each card also names its board, and clicking one from another board switches the application to it.

## ② Board pane

Every schematic that has worklogs in the selected workbook, with the markers drawn on it.

* **Click a marker** - opens that worklog in the editor
* **Click anywhere else on a schematic** - selects it, and the list on the right switches to its worklogs

## ③ Worklog list

One card per worklog on the selected schematic: number and title, description, category and state, and a line of totals (hours, cost, and how many comments/links/photos/files it holds).

**Click a card** to open the worklog. **The ✕ in its corner deletes it** - see below.

## The header

The selected workbook's number, title, status and note, plus four buttons:

| Button | Does |
| --- | --- |
| Edit workbook | Change the description and note |
| Delete workbook | Deletes the whole job - see [Export and your data](Workbooks-Export-and-data) |
| Export to PDF | The customer document |
| Export to ZIP | That PDF plus the original photos and files |

## The totals strip

Under the header:

```
▸  7 worklogs · 12.5 h · 430 · 4 open
```

Click the arrow to expand it into a breakdown by category, by state, by attachment, and by component.

Hours and cost come from the Work done lines in every worklog. There is no currency symbol - it is your own number back again.

## Find a previous repair

The search box filters everything - the workbook list, the board pane and the worklog list - and highlights what matched.

| You type | It finds |
| --- | --- |
| `cpu` | Anything containing "cpu" |
| `cpu socket` | Records containing **both** words |
| `"cracked socket"` | That exact run of text |
| `-psu` | Excludes anything containing "psu" |

It searches everything you have typed: titles, descriptions, comments, work done, component names, link and file names.

It does **not** search numbers (hours, cost, dates, id numbers) or the words Open/Closed - those two would match nearly everything, and both already have a pill you can see.

Searching never changes which workbook you are writing into. The box clears itself when you change board.

---

**Next:** [Export and your data](Workbooks-Export-and-data) — PDF/ZIP, where files live, deleting
