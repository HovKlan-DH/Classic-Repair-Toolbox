# The worklog bar

← [Back to Workbooks](Workbooks)

The worklog bar is the strip directly above the tab headers. It is **always visible, on every tab**,
so the repair you are working on stays in front of you whether you are looking at a schematic, the
oscilloscope or the parts overview.

It appears only when **"Enable Workbooks tab"** is ticked in Configuration.

```
┌────────────────────────────────────────────────────────────────────────────────────┐
│ Workbook  [ #3 · Dead C64, no video ▾ ]  [🔓 Open]  3 worklog entries · started …   │
│                                    Show worklogs ☑   [Add worklog]  [Create new …] │
└────────────────────────────────────────────────────────────────────────────────────┘
```

---

## The workbook picker

The dropdown names the **active workbook** — the one every worklog control acts on. See
[the active workbook](Workbooks-Concepts-and-vocabulary#the-active-workbook) for how it is chosen.

**The picker lists every workbook on every board**, not just the current one. This is deliberate: it
makes the bar a navigator. Picking a workbook that belongs to a different board **switches the app
to that board first**, then activates it — so you can jump straight back to a job on another machine
without hunting through the hardware dropdowns.

Picking a workbook here is exactly the same action as clicking its card on the Workbooks tab. Your
choice is saved per board and survives a restart.

Beside the picker:

- a **status pill** — Open or Closed, with a padlock,
- a line of detail — how many worklogs the workbook holds and the date it was started
  (`3 worklog entries · started 2026-September-04`), or `No worklog entries yet · started …` for a
  brand-new one.

If the board has no workbooks at all, the bar says so and only "Create new workbook" is available.

---

## Show worklogs

A checkbox. When ticked, the Schematics tab draws every worklog in the active workbook onto the
board:

- a dashed, category-coloured rectangle for each worklog that has a **marked area** shown,
- a **"#N" pill** for each worklog — anchored to its area, or parked in the top-right corner when
  the worklog has no shown area,
- matching pills on the **thumbnails** in the schematic gallery.

Click any pill to open that worklog in the editor.

The setting is remembered between sessions. It defaults to on.

Turning it off costs nothing — the app skips the overlay work entirely rather than drawing it
invisibly — so leave it off if you find the markers distracting while tracing.

See [Marking areas on a schematic](Workbooks-Marking-areas-on-a-schematic).

---

## Add worklog

Starts **area-marking mode** on the Schematics tab.

1. Click **Add worklog**. A hint appears in the tab-header row telling you what to do.
2. Drag a rectangle on the schematic around the area you want to describe.
3. On release, the full [worklog editor](Workbooks-The-worklog-editor) opens on that area, with every
   component your rectangle touched already ticked into scope.
4. Save, and the worklog is created. Cancel, and **nothing whatsoever is written**.

While the mode is active the bar shows a **Cancel entry** button to back out without drawing.

The button is unavailable when the board has no workbook — create one first.

> **Note:** the worklog is created against the schematic you drew on, and against the workbook that
> was active when you started. Switching the active workbook mid-draw cancels the mode, precisely so
> a worklog cannot be written into a workbook you have since navigated away from.

---

## Create new workbook

Opens the create dialog for the **currently selected board**. See
[Getting started](Workbooks-Getting-started#3-create-a-workbook).

The workbook it creates is made active immediately, so the next thing you draw goes into it rather
than into whatever was active before.

---

## What refreshes the bar

Everything about the bar — the picker, the pill, the counts, the "Show worklogs" state — is rebuilt
whenever anything about the worklog data changes: a board change, a workbook created, edited,
activated or deleted, a worklog saved or deleted, an attachment filed.

The same refresh redraws the Schematics tab's overlay rectangles and thumbnail pills, so what is on
the board can never lag behind what is in the workbook.

---

**Next:** [The Workbooks tab](Workbooks-The-Workbooks-tab) · [Marking areas on a schematic](Workbooks-Marking-areas-on-a-schematic)
