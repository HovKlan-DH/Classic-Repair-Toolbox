[Wiki Home](Home)

Record what you find on a board and what you did about it — and share the finished repair as a PDF.

---

* [Getting started](Workbooks-Getting-started) - turn it on and record your first repair
* [Daily use](Workbooks-Daily-use) - the worklog bar, the editor, marking areas on a schematic
* [Browsing and search](Workbooks-Browsing-and-search) - find an older repair, and read the totals
* [Export and your data](Workbooks-Export-and-data) - PDF/ZIP export, where files are stored, deleting

## The two words

**Workbook** = one repair on one board.\
**Worklog** = one thing you noted down inside that repair.

```
Workbook  "Dead C64, no video - attic find"       #3, Open
├── Worklog #1  "VIC socket pin cracked"          Issue, Closed
├── Worklog #2  "Check the 9V AC rail"            Note, Open
└── Worklog #3  "Scratched case lid"              Cosmetic, Open
```

A workbook belongs to the board that was selected when you created it, and cannot be moved to another board afterwards.

## Categories and states

Every worklog has a category and a state:

| Category | Use for |
| --- | --- |
| Note | An observation |
| Cosmetic | Condition, not function |
| Issue | An actual fault |

State is **Open** or **Closed**. The category decides the colour the worklog is drawn in on the board.

## Open or Closed workbooks

You never set this yourself. A workbook is **Closed** when it has at least one worklog and all of them are Closed - otherwise it is Open.

So closing the last worklog closes the workbook, and adding a new worklog reopens it.

While a workbook is Open, it is shown as **started** on the date it began. Once it is Closed, it is shown as **ended** on the day the last outstanding worklog was closed. Reopening and closing it again updates that date; the start date never changes.

This reads the same everywhere the workbook appears - the bar above the tabs, its card in the workbook list, and the header of an exported PDF.

A workbook closed by an older version of the application has no ending date recorded, so it keeps showing its start date.

## Where it is stored

On your own machine only. Never uploaded, never synced.
