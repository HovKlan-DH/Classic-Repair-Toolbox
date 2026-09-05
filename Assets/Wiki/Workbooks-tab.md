[Wiki Home](Home)

Record what you find on a board and what you did about it — and hand the finished job to a customer as a PDF.

---

* [Getting started](Workbooks-Getting-started) - turn it on and record your first repair
* [Daily use](Workbooks-Daily-use) - the worklog bar, the editor, marking areas on a schematic
* [Browsing and search](Workbooks-Browsing-and-search) - find an older repair, and read the totals
* [Export and your data](Workbooks-Export-and-data) - PDF/ZIP export, where files are stored, deleting

## The two words

**Workbook** = one repair job on one board.\
**Worklog** = one thing you noted down inside that job.

```
Workbook  "Dead C64, no video - J. Hansen"        #3, Open
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

## Where it is stored

On your own machine only. Never uploaded, never synced.
