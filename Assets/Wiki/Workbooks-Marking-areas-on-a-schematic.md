# Marking areas on a schematic

← [Back to Workbooks](Workbooks)

A worklog can point at a place on the board. This page covers drawing that area, what gets drawn on
the schematic, and the alternative for worklogs that are not about a place.

---

## Drawing an area

1. Make sure the right **workbook is active** and the right **schematic is open** on the Schematics
   tab.
2. Click **"Add worklog"** on the worklog bar.
3. A hint appears above the tabs: *"Now mark an area on the schematics image, to select the
   components in scope of your worklog."*
4. **Drag a rectangle** around the area — a chip, a section of circuit, a connector.
5. On release, the [worklog editor](Workbooks-The-worklog-editor) opens on that area, with every component
   the rectangle touched already ticked into scope.

Save, and the worklog exists. **Cancel, and nothing is written at all** — the rectangle goes with
the mode.

Use **Cancel entry** on the bar to leave the mode without drawing.

The area is stored in the schematic image's own pixel coordinates, so it stays put however you zoom
or pan, and it is drawn in the right place in the exported PDF regardless of page size.

> Switching the active workbook while area-marking mode is open **cancels the mode**. The mode
> captured which workbook it would write into when it started, and silently writing into a different
> one would be worse than making you start again.

---

## What appears on the board

With **"Show worklogs"** ticked on the worklog bar, each worklog in the active workbook is drawn on
its schematic as:

- a **dashed rectangle** in the worklog's category colour, around its marked area,
- a **"#N" pill** anchored to that area, carrying the worklog's number, its category colour and a
  padlock in its state colour.

So a **Closed Issue** shows as a green padlock on a red badge — two different pieces of information,
in two different channels.

The same pills appear on the **thumbnails** in the schematic gallery, so you can see which schematics
have work recorded against them without opening each one.

**Click any pill** — on the board, on a thumbnail, or on a Workbooks-tab preview — to open that
worklog in the editor.

---

## Parked pills: worklogs with no area

Not every worklog is about a place on the board. "Customer reports intermittent fault", "case
scratched", "measured 4.1V at the bench" — these have nothing useful to point at.

Untick **"Show marked area on schematics image"** in the editor and the worklog gets **no rectangle**.
Its pill is **parked in the top-right corner** of the schematic panel instead, stacking with any
other parked pills rather than overlapping them, and wrapping into a grid when there are many.

A parked worklog is not hidden. It is still listed, still clickable, still exported — it simply is
not claiming to mark a spot it does not have.

Worklogs created from an [oscilloscope capture](Workbooks-Attaching-oscilloscope-captures) start parked, for
exactly this reason: they were born at the bench with a probe in hand, not by dragging a rectangle.

### Turning an area back on

Tick "Show marked area" on a parked worklog and it gets a real, usable square in the board's
**bottom-right** corner — deliberately the opposite corner from the parked pills, so a fresh area is
never mistaken for one while you drag it into place.

A worklog that **already has** a drawn area keeps it through any number of tick/untick cycles. The
app only ever adds an area; it never moves one you placed.

---

## The three surfaces that draw pills

The same worklog is drawn in three places, and they agree with each other:

| Surface | Where |
| --- | --- |
| The schematic image | Schematics tab, main view |
| The thumbnails | Schematics tab, gallery strip |
| The board previews | Workbooks tab, board pane |

All three follow the same rule — anchored when the area is shown, parked in the top-right corner
when it is not — and all three redraw whenever a worklog changes, wherever you changed it from.

The **exported PDF** follows the same rule again, so a printed document matches the screen it came
from.

---

**Next:** [Components in scope](Workbooks-Components-in-scope) · [Exporting a workbook](Workbooks-Exporting-a-workbook)
