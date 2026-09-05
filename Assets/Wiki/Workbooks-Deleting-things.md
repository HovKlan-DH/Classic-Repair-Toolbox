# Deleting things

← [Back to Workbooks](Workbooks)

Both workbooks and individual worklogs can be deleted. Both are **permanent** — there is no undo and
no recycle bin inside the app.

---

## Deleting a worklog

Each worklog card in the [Workbooks tab](Workbooks-The-Workbooks-tab#-the-worklog-list)'s right-hand list
carries a **"Delete worklog"** button in its top-right corner.

A confirmation dialog names the worklog by the same `#N · Title` its card and its board pill show, so
you can be sure which one you are about to lose. (An untitled one is named `#N · (untitled)`.)

**What is removed:**

- the worklog's row in `entries.json`,
- its entire `worklog_{id}` attachment folder — **every photo and file attached to it**.

**What happens afterwards:**

- The remaining worklogs **keep their numbers**. Nothing is renumbered. See
  [below](#why-numbers-are-never-reused).
- The workbook's status is recalculated. Deleting the last still-Open worklog **closes** the
  workbook; deleting the only worklog **reopens** it.
- The board overlay, the thumbnail pills and the Workbooks tab all refresh, so the worklog is gone
  from every surface at once.

> **Enter and Escape both cancel** in the confirmation dialog, even when the Delete button has
> keyboard focus. Deleting is a deliberate click, never a reflexive keypress.

---

## Deleting a workbook

**"Delete workbook"** in the header, beside "Edit workbook".

A confirmation names the workbook — several cards can be on screen at once, so it says which.

**What is removed: the workbook's entire folder.** Every worklog in it, every photo, every attached
file. One folder is the whole workbook, so deleting the folder deletes the job completely.

**What happens afterwards:** the tab lands on the board's next remaining workbook automatically. If
the deleted workbook was the active one, its saved id no longer names anything, and the app falls
back to the newest remaining — the same fallback that covers a workbook you deleted by hand outside
the app.

---

## Before you delete

There is no undo. If there is any chance you will want the job back:

- **[Export it to ZIP](Workbooks-Exporting-a-workbook) first.** That gives you the full document plus every
  original photo and file in one archive.
- Or copy the workbook's folder out of the
  [`Workbooks` folder](Workbooks-Where-your-data-is-stored#the-layout).

Remember that a **closed** workbook costs you nothing to keep. It stays in the list, does not
interfere with anything, and is exactly the record you will want when the same board comes back.
Closing is almost always the right answer; deleting is for mistakes.

---

## Why numbers are never reused

Delete workbook #2 of two, and the next one you create is **#3**, not #2. The gap stays. The same
applies to worklog numbers inside a workbook.

This is deliberate, and it protects two things:

**Documents already sent.** A workbook id names an exported file —
`Workbook_2_Commodore_C64_20260904.pdf` — which has very likely already been emailed to a customer.
If #2 were handed to a different repair, that document would silently start describing the wrong
job.

**Files on disk.** Ids name real folders (`2/`, `worklog_2/`). A reused number would let a new record
inherit a deleted one's photos and attached files.

The app records the highest id it has ever **handed out**, rather than looking at what survives —
in `counters.json` for workbooks, and in each workbook's own `index.json` for its worklogs. A deleted
number is spent.

**A gap in the numbering is not a bug.** It is the correct record that something existed there.

For the same reason, deleting a worklog does not renumber the ones after it: their ids are what the
board pills, the cards and any already-exported PDF show.

---

## Deleting by hand

Deleting a workbook's folder from the `Workbooks` folder works — the app scans the folders and simply
will not find it. The active-workbook fallback handles it cleanly.

**Do not renumber folders to close a gap.** You would be reintroducing exactly the collision the
counters exist to prevent.

---

**Next:** [Where your data is stored](Workbooks-Where-your-data-is-stored) · [Troubleshooting and FAQ](Workbooks-Troubleshooting-and-FAQ)
