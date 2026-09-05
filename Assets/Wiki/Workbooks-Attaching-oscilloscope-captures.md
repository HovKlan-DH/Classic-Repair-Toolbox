# Attaching oscilloscope captures

← [Back to Workbooks](Workbooks)

When you capture a waveform from the oscilloscope, you can file it straight into a worklog without
leaving the bench.

---

## Where the button is

Take a capture from the component popup as usual. The **"Saved image as […]"** banner that appears
afterwards carries an extra button:

> **Attach image to worklog**

That is the whole entry point. The button is hidden when the Workbooks feature is off, or when the
board has no workbook to file into.

**Why a button on the banner and not a mode you turn on first:** at the bench you are sweeping pin
after pin, and being prompted after every capture would be intolerable. A button you can ignore
costs nothing.

**The capture is saved to the oscilloscope image folder first, before any of this.** Cancelling the
dialog — or a failed attach — costs you the filing, never the measurement.

---

## The dialog

One modal, four things:

| Field | |
| --- | --- |
| **Image** | A preview of the capture you just took |
| **Workbook** | Named, not chosen — see below |
| **Worklog** | The dropdown you actually answer |
| **Comment** | Optional, stored against the attached photo |

**Ctrl+Enter** confirms.

### The workbook is named, not asked

Which workbook is active is already settled app-wide, so the dialog states it rather than asking
again. It is named explicitly because this dialog opens from the component popup, which can be
sitting over a schematic while you have been looking at the scope — so you can see where the capture
is about to go.

### The worklog list is ranked

The dropdown puts the likely answer first, under two visible headings:

```
  Worklogs with [U8] in scope        ← faint, not selectable
    #2 · Ripple on the 5V rail
    #5 · U8 output stuck low
  All other worklogs                 ← faint, not selectable
    #1 · Cracked trace at CN2
    #3 · Case scratched
  Create new worklog
```

**Worklogs that already have the component you are probing in scope come first**, then everything
else. Within each band, worklogs are in **ascending id order** — the same `#1, #2, #3` counting order
the board pills show, so the list looks ordered rather than arbitrary.

The first real worklog is **preselected**, so probing U8 while working a fault on U8 is a single
click.

**Closed worklogs are kept, not hidden.** Re-measuring a repair you already finished is exactly what
you do when a board comes back, and hiding the worklog that describes that repair would send the
measurement to a new one instead.

The headings only appear when there is actually a match to separate out.

### Create new worklog

Always offered, and always **last**, so it never displaces a real worklog from the preselected slot.

Picking it changes the button to read **"Create worklog"** and opens the full
[worklog editor](Workbooks-The-worklog-editor) on a new worklog with the capture **already attached** — you
are not made to save an empty worklog and come back for the photo. Probing before anything is
written down is how diagnosis starts.

That new worklog is:

- filed against **the schematic currently open** on the Schematics tab, so it appears on the board
  like any other,
- created with **no marked area**, so its pill parks in the corner — it was born at the bench with a
  probe in hand, not by dragging a rectangle. Tick "Show marked area" later if you want to place it.

---

## After attaching

The capture becomes an ordinary **photo** on the worklog: it appears in the Photos list, can be given
a comment, reordered, viewed full size, and it goes into the exported PDF and ZIP like any other.

The worklog bar, the Schematics overlay and the Workbooks tab all refresh, so the worklog's counts
are up to date immediately.

If the full editor happens to be open on that same worklog at the time, the attach still lands
correctly — it re-reads the worklog from disk rather than trusting a copy that may have moved on.

---

**Next:** [The worklog editor](Workbooks-The-worklog-editor) · [Exporting a workbook](Workbooks-Exporting-a-workbook)
