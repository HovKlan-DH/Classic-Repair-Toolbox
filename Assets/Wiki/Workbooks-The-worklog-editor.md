# The worklog editor

← [Back to Workbooks](Workbooks)

The worklog editor is the one window where a worklog is written, whether you are creating it or
changing it later. Everything a worklog can hold is here.

## Opening it

| From | How |
| --- | --- |
| **Creating** | "Add worklog" on the worklog bar, then drag an area on the schematic |
| **Editing** | Click a worklog's **"#N" pill** on the Schematics tab |
| | Click a **pill** on a board preview on the Workbooks tab |
| | Click a **worklog card** in the Workbooks tab's worklog list |
| | Click a pill on a schematic **thumbnail** |

All of these open the *same* window with the same contents — a card is simply the same worklog drawn
larger than its pill.

The window remembers its size and position, and which monitor you last had it on.

---

## The top fields

- **Title** — the one-line headline. **Save stays disabled until this has text.**
- **Description** — the longer comment. Web links typed in here become clickable.
- **Category** — Note / Cosmetic / Issue, as pills. The chosen one is filled with its colour.
- **State** — Open / Closed, as pills, with a padlock.
- **"Show marked area on schematics image"** — see below.

Changing the category or state writes an automatic comment into the worklog's own Comments list, so
the worklog records its own history: `Worklog changed to "Issue"`, `Worklog closed`.

**Ctrl+Enter** saves.

---

## Show marked area

Ticked, the worklog is drawn on the board as a dashed rectangle with its pill anchored to it.
Unticked, there is no rectangle and the pill **parks in the top-right corner** of the schematic
panel instead.

A parked worklog is not hidden — you can always reach it by its pill. Unticking is for worklogs that
are not really *about* a place on the board: "customer reports intermittent fault", "case scratched",
a measurement taken at the bench.

**Ticking it on a worklog that never had an area gives it a real, draggable square** in the board's
bottom-right corner — the opposite corner from the parked pills, so a freshly placed area cannot be
confused with one. A worklog that already has a drawn area **keeps it** through any number of
tick/untick cycles; the app never moves an area you placed yourself.

See [Marking areas on a schematic](Workbooks-Marking-areas-on-a-schematic).

---

## Worklog location

A read-only preview of the schematic with this worklog's area shown on it, so you can see where you
are without switching back to the Schematics tab.

---

## The lists

Seven collapsible sections. **Which ones you fold away is remembered per worklog**, so a worklog you
only ever use for photos opens with the rest tucked out of the way.

### Links of interest

A headline plus a URL. Forum threads, datasheets, part suppliers.

A link row is a **declared destination**, so it is always clickable — and if you leave the scheme
off (`example.com`), `https` is filled in for you when the row is opened or exported.

### Work done

A dated note with **hours spent** and **cost**. The section shows a running total across every row,
and those totals feed the worklog's stats row, the workbook summary strip and the exported PDF.

The cost is a plain number with no currency symbol anywhere in the app — use whatever currency you
bill in.

Sortable newest-first or oldest-first; the choice is remembered across worklogs and restarts.

### Comments

A running log. Every comment is dated.

This list also carries the **automatic comments** the app writes — `Worklog created`, `Worklog
opened`, `Worklog closed`, `Worklog changed to "Cosmetic"` — mixed in with your own, so the list
reads as the worklog's history.

Sortable the same way as Work done.

### Components in scope · Components completed

Two checklists. See [Components in scope](Workbooks-Components-in-scope) — they have enough behaviour to
deserve their own page.

### Photos

Images with a comment each. Add them with **Add photo**; click one to open it full size.

**Drag a photo up or down to reorder it.** The order you set is the order they appear in the
exported PDF, so you can put the "before" shot before the "after" one.

### Files

Anything that is not an image: datasheets, scope captures, invoices, board scans. Same shape as
Photos — a file plus a comment, reorderable — and they travel in the **ZIP** export.

---

## Saving

- **Update worklog** / **Add worklog** — the button is labelled for what it will do.
- **Cancel** — see below, the behaviour differs between a new worklog and a saved one.

### A new worklog is held entirely in memory until you save

Nothing at all reaches disk until you press **Add worklog**. Cancel — or closing the window — leaves
no trace: no half-made worklog in the workbook, and any photos or files you attached are cleaned up.

### A saved worklog writes its lists through immediately

For a worklog that already exists, changes to the **lists** (a comment added, a photo attached, a
reorder) are written the moment you make them. Cancel does *not* take those back — it only discards
unsaved changes to the top fields (title, description, category, state).

This is deliberate: it means a long session of attaching photos cannot be lost by a mis-click on
Cancel.

> **Photos and files are the exception in both cases** — the bytes are copied as soon as you attach
> them, because a photo has to exist somewhere before it can be shown. For a new worklog those bytes
> are cleaned up if you cancel.

If a save cannot reach disk — a locked file, a full disk — the editor **says so** rather than
closing as though it worked.

---

**Next:** [Components in scope](Workbooks-Components-in-scope) · [Marking areas on a schematic](Workbooks-Marking-areas-on-a-schematic)
