[Wiki Home](Home) · [Workbooks](Workbooks-tab)

The bar, the editor, and the markers on the board.

---

## The worklog bar

The strip above the tabs, visible on every tab.

| Control | Does |
| --- | --- |
| Workbook drop-down | Which workbook you are working in |
| Status pill | Open or Closed |
| **Show worklogs** | Draws the markers on the schematic |
| **Add worklog** | Start marking an area (see below) |
| **Create new workbook** | New job on the selected board |

The drop-down lists workbooks from **all** boards. Picking one from another board switches the application to that board first.

## Adding a worklog

**Add worklog** → drag a rectangle on the schematic → the editor opens.

If you want to back out, click **Cancel entry** in the bar.

## The editor

Opens when you create a worklog, or when you click a `#N` marker, a pill or a card.

Top of the window:

* **Title** - required
* **Description**
* **Category** - Note / Cosmetic / Issue
* **State** - Open / Closed
* **Show marked area on schematics image** - see below

Below that, seven lists you fill in as the job runs: Links, Work done, Comments, Components in scope, Components completed, Photos, Files.

**Work done** takes hours and cost per line and totals them. Cost is a plain number - the application never asks about currency, so use whatever you bill in.

**Photos** can be dragged up and down; that order is the order they appear in the PDF.

Ctrl+Enter saves.

### What Cancel does

| | Cancel |
| --- | --- |
| New worklog | Nothing is written at all |
| Existing worklog | Only Title/Description/Category/State are discarded |

On an existing worklog, anything you add to the lists (a comment, a photo) is saved the moment you add it. Cancel does not take those back.

## Markers on the board

With **Show worklogs** ticked, each worklog is drawn on its schematic as a dashed rectangle in its category colour, with a `#N` marker. The same markers appear on the thumbnails.

Click any marker to open that worklog.

### Worklogs without an area

Untick **Show marked area on schematics image**, and the worklog gets no rectangle - its marker parks in the top-right corner of the panel instead.

Use it for things that are not about a place on the board: "customer reports intermittent fault", "case scratched".

Tick it again and the worklog gets a square in the bottom-right corner of the board, which you drag where you want. A worklog that already has an area keeps it.

## Components in scope

Two checklists in the editor:

* **Components in scope** - which components the worklog is about. Everything your rectangle touched starts ticked.
* **Components completed** - which of those you have finished. Useful for "replace every electrolytic" jobs.

Both have **All** and **None** buttons.

If the section is missing entirely, the board data has not finished loading, or a region filter is hiding the components - reopen the worklog once the board is loaded. Your saved components are not lost in the meantime.

## Filing an oscilloscope capture

After a capture, the "Saved image as [...]" banner in the component popup has an **Attach image to worklog** button.

It asks which worklog. Worklogs that already have the component you were probing in scope are listed first, and the first one is preselected - so probing U8 while working a fault on U8 is one click.

**Create new worklog** is always last in the list. It opens the editor with the image already attached, and the new worklog is filed against the schematic you have open, with no marked area.

---

**Next:** [Browsing and search](Workbooks-Browsing-and-search) — find an older repair, and read the totals
