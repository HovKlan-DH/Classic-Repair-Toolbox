# The summary strip

← [Back to Workbooks](Workbooks)

A collapsible strip under the header on the [Workbooks tab](Workbooks-The-Workbooks-tab), showing what the
selected workbook adds up to.

---

## The headline

Always visible, on one line:

```
▸  7 worklogs · 12.5 h · 430 · 4 open
```

| Part | Meaning |
| --- | --- |
| **7 worklogs** | How many worklogs the workbook holds |
| **12.5 h** | Total hours across every Work-done row in every worklog |
| **430** | Total cost, same way. **No currency symbol** — see below |
| **4 open** | How many worklogs are still Open |

Every **number** is bold and the words are not, so the figures read at a glance.

### Why there is no currency symbol

The app never asks which currency you work in, so printing one would be a guess. The number you
typed is the number you get back — in the strip, in the worklog cards, in the editor's own total and
in the exported PDF alike.

---

## The breakdown

Click the chevron to expand. Collapsed is the default, and your choice is remembered (per user, not
per board).

**By category** and **by state**, drawn as pills carrying their count:

```
[ 3 Note ]  [ 1 Cosmetic ]  [ 3 Issue ]
[ 4 Open ]  [ 3 Closed ]
```

The two rows are kept separate deliberately — run together, five pills read as one undifferentiated
list.

**Every value is shown, including the zeroes.** `0 Issue` is information: it says this workbook
records no faults. It also keeps the row from changing width as the job progresses.

**By attachment:**

```
12 comments · 4 links · 9 photos · 2 files · 15 work done
```

**By component:**

```
11 components in scope · 6 completed
```

Components are counted **distinctly** across the workbook. The same chip legitimately appears in
several worklogs — checked in one, replaced in another — and adding those up would report more
components than the board has.

The component line is **hidden entirely** when the workbook scopes none, rather than showing a
permanent zero.

The whole strip is hidden when no workbook is selected.

---

## Where else these numbers appear

The **exported PDF** prints this same summary as its opening section, from the same calculation. An
exported document cannot report different totals from the screen it was produced from.

The per-worklog stats row on each [worklog card](Workbooks-The-Workbooks-tab#-the-worklog-list) is the same
idea one level down — that worklog's own hours, cost and attachment counts.

---

## A note on pill icons

A pill carrying a **count** drops its icon; a pill without a count keeps it.

That is deliberate. On a worklog card the padlock is the only thing separating Open from Closed at a
glance. In a counted pill, a glyph sitting between a number and its label reads as a third piece of
information rather than as decoration.

---

**Next:** [Exporting a workbook](Workbooks-Exporting-a-workbook) · [The Workbooks tab](Workbooks-The-Workbooks-tab)
