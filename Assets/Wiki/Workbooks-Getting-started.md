# Getting started

← [Back to Workbooks](Workbooks)

This page takes you from a fresh install to a workbook with a worklog in it.

---

## 1. Turn the feature on

Go to **Configuration → Workbooks** and tick **"Enable Workbooks tab"**.

This one checkbox controls the whole feature. With it ticked you get:

- the **worklog bar** — a strip directly above the tab headers, visible on every tab,
- the **Workbooks tab** itself.

With it unticked, both disappear. Nothing is deleted — your workbooks stay on disk and come back
when you tick it again.

Underneath the checkbox is a **Scope in Workbooks tab** choice:

| Option | Effect |
| --- | --- |
| **Show only workbooks for selected board** (default) | The Workbooks tab lists just this board's jobs |
| **Show all workbooks** | The tab lists every workbook on every board, each card naming its board |

Start with the default. "Show all workbooks" is useful once you have jobs across several machines
and want to find one without first switching to its board — see
[The Workbooks tab](Workbooks-The-Workbooks-tab#scope-showing-other-boards-workbooks).

---

## 2. Select the board you are repairing

Pick the hardware and board in the two dropdowns at the top left of the main window, exactly as you
would to look at its schematics.

**Workbooks belong to a board.** Whichever board is selected when you create a workbook is the board
that workbook is attached to, permanently. There is no way to move a workbook to a different board
afterwards, so make sure the right one is selected.

---

## 3. Create a workbook

Click **"Create new workbook"** on the worklog bar. (The same button appears on the Workbooks tab.)

The dialog asks for:

- **Description** — required. The one-line title for the job. "Dead C64, no video — J. Hansen".
- **Note** — optional. Any extra context: the customer's details, what they reported, what you were
  told over the phone.

The dialog shows the id the workbook will get (`#1` on your first). Press **Create workbook**, or
**Ctrl+Enter**.

The new workbook is immediately made **active**, so everything you do next goes into it.

> **Tip:** one workbook per *job*, not per fault. A board that comes in with three separate problems
> is one workbook with three worklogs in it. If the same board comes back six months later, that is
> a new workbook.

---

## 4. Add your first worklog

Go to the **Schematics** tab and open the schematic where the thing you want to record lives.

Click **"Add worklog"** on the worklog bar. The bar tells you what to do next, and the app goes into
area-marking mode.

**Drag a rectangle** on the schematic around the area you are describing — the chip, the section of
circuit, the corner of the board. As soon as you release the mouse, the full **worklog editor**
opens on that area.

Fill in at least a **Title** — the Save button stays disabled until you do. Everything else is
optional and can be added later.

While you are there, note two things:

- **Category and State** at the top — Note/Cosmetic/Issue, Open/Closed.
- **"Mark components in scope"** — the app has already worked out which components your rectangle
  touches and ticked them all. Untick any that are not actually part of this worklog.

Press **Add worklog** to save. Press **Cancel** and nothing is written at all — a new worklog is
held entirely in memory until you save it.

Your worklog now appears:

- as a coloured rectangle and a **"#1"** pill on the schematic (if "Show worklogs" is ticked on the
  bar),
- as a pill on the schematic's **thumbnail** in the gallery,
- as a card on the **Workbooks** tab.

---

## 5. Work the job

As the repair progresses, click any pill or card to reopen the worklog editor and add:

- **Work done** — a dated note with hours and cost, which the app totals for you,
- **Comments** — an ongoing log,
- **Photos** — before/after shots, with a comment each,
- **Files** — datasheets, scope captures, invoices,
- **Links** — forum threads, part suppliers.

Flip a worklog to **Closed** when it is dealt with. When the last one closes, the whole workbook
closes itself.

See [The worklog editor](Workbooks-The-worklog-editor) for everything the editor can do.

---

## 6. Hand it over

On the **Workbooks** tab, with the workbook selected, use:

- **Export to PDF** — the customer-facing document: every worklog, grouped by schematic, with the
  marked areas drawn on the board images, the photos, the totals.
- **Export to ZIP** — that same PDF *plus* every original photo and attached file, one folder per
  worklog.

See [Exporting a workbook](Workbooks-Exporting-a-workbook).

---

## A worked example

> A C64 comes in dead. You select `Commodore / C64 / 250407` and create the workbook
> **"Dead C64 — no picture, no sound (customer: J. Hansen)"**, note *"Reported as working until a
> thunderstorm."*
>
> On the *Power* schematic you drag an area around the 7805 and record worklog **#1 — "5V rail at
> 4.1V under load"**, category *Issue*, state *Open*, components in scope `U1`. You attach a scope
> capture of the rail.
>
> On the *Video* schematic you drag an area around the VIC-II and record **#2 — "VIC socket pin 8
> cracked"**, category *Issue*, components in scope `U19`. Two photos.
>
> You also spot a scratch on the case and record **#3 — "Case lid scratched, customer informed"**,
> category *Cosmetic* — no marked area needed, so its pill parks in the corner.
>
> You replace the regulator (work done: 0.5 h, 40.00) and close #1. You reflow the socket (1.5 h,
> 120.00) and close #2. You close #3 once the customer confirms they don't care.
>
> With all three closed, the workbook closes itself. You **Export to ZIP** and email
> `Workbook_1_Commodore_C64_20260904.zip` — the PDF plus the original photos and the scope capture.

---

**Next:** [Concepts and vocabulary](Workbooks-Concepts-and-vocabulary) · [The worklog bar](Workbooks-The-worklog-bar)
