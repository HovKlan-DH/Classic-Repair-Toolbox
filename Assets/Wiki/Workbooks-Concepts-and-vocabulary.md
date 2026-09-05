# Concepts and vocabulary

← [Back to Workbooks](Workbooks)

The Workbooks feature has a small vocabulary. Getting it straight makes everything else easier.

---

## Workbook

**One repair job, on one board.** "Dead C64, no video — customer J. Hansen" is a workbook.

A workbook belongs to exactly one board — a specific hardware/board combination such as
`Commodore|C64` — chosen implicitly: whichever board is selected when you create it. It carries:

| Field | Meaning |
| --- | --- |
| **Id** | `#3` — assigned automatically, never reused. See [numbering](#id-numbering) below |
| **Description** | The one-line title. The create dialog labels this field "Description"; it is required |
| **Note** | Optional free text — extra context, a customer's details, anything you like |
| **Status** | Open or Closed. **Computed, never set by hand** — see [status](#workbook-status) below |
| **Start date** | When you created it |
| **Board** | Which board the job is on |

Everything belonging to a workbook — its worklogs, their photos, their attached files — lives inside
one folder on disk. See [Where your data is stored](Workbooks-Where-your-data-is-stored).

---

## Worklog

**One thing you noted down inside a workbook.** A bad capacitor, a lifted trace, a measurement to
come back to, a scratch on the case.

The app calls these **worklogs** everywhere you can see one. (Internally the code calls them
"entries" — you may see that word in log files, but not in the interface.)

A worklog carries:

| Field | Meaning |
| --- | --- |
| **Id** | `#4` — numbered within its workbook, never reused |
| **Title** | The one-line headline |
| **Description** | The longer comment |
| **Category** | Note, Cosmetic or Issue |
| **State** | Open or Closed |
| **Schematic** | Which schematic image it is filed against |
| **Marked area** | Optionally, a rectangle on that schematic |
| **Components in scope** | Board labels — `U19`, `C64` — this worklog covers |
| **Components completed** | Which of those you have ticked off |
| Seven lists | Links, Comments, Work done, Photos, Files, and the two component lists above |

---

## Categories

Three, fixed:

| Category | For |
| --- | --- |
| **Note** | An observation. Something to remember, not necessarily a fault |
| **Cosmetic** | Condition, not function — a scratch, a yellowed case, a missing screw |
| **Issue** | An actual fault |

The category decides the **colour** a worklog is drawn in — its pill on the board, its badge, its
marked-area outline, and its chip in the exported PDF.

Changing a worklog's category writes an automatic comment into its own Comments list
(`Worklog changed to "Issue"`), so the worklog records its own history.

---

## States

Two, fixed: **Open** and **Closed**.

A state is about the *worklog* — is this particular thing dealt with? It is drawn as a pill with a
padlock: open padlock for Open, closed padlock for Closed.

Flipping a state also writes an automatic comment (`Worklog opened` / `Worklog closed`).

---

## Workbook status

A workbook's **Open/Closed status is computed, not chosen.** There is no button to close a workbook.

> A workbook is **Closed** when it has at least one worklog and *every* worklog in it is Closed.
> It is **Open** in every other case — including when it has no worklogs at all.

This is recalculated every time a worklog is added, changed or deleted. So:

- Closing the last outstanding worklog **closes the workbook**.
- Adding a new worklog to a closed workbook **reopens it**.
- Deleting the only worklog **reopens it** (a workbook with no worklogs is Open).

A closed workbook is not hidden or archived — it stays in the list, can still be selected, edited and
exported, and can be reopened simply by adding work to it.

---

## The active workbook

A board can have many workbooks. Exactly one of them is **active** at a time, and that is the one
every worklog control acts on: the worklog bar, "Show worklogs", "Add worklog", and the Workbooks
tab's board pane.

The rule is:

1. If you have clicked a workbook card (which saves your choice, per board), that workbook is active
   — as long as it still exists.
2. Otherwise, the board's **newest** workbook (highest id), open or closed.

Your choice is remembered per board and survives restarts. If the workbook you chose is later
deleted, the app falls back to the newest remaining one rather than showing nothing.

**Activating is deliberately status-blind** — a closed workbook can be the active one. If it were
not, a workbook would vanish from the interface the moment you finished it, which looks like data
loss.

Searching does **not** change which workbook is active. Typing in the search box narrows what you
*see*; it never redirects where "Add worklog" writes.

---

## Id numbering

Ids are assigned in order and **never reused**, even after a delete.

Delete workbook #2 of two, and the next workbook you create is **#3**. The gap is deliberate — it is
the correct record that #2 existed.

This matters more than it looks:

- A workbook id names an exported PDF (`Workbook_2_Commodore_C64_20260904.pdf`) that has very likely
  already been emailed to a customer. Handing #2 to a different repair would silently make that
  document describe the wrong job.
- Ids name real folders on disk. A reused number would let a new record inherit a deleted one's
  photos and files.

The same rule applies to worklog ids inside a workbook. Worklog numbering is **per workbook**, so
every workbook starts at #1.

Deleting a worklog does **not** renumber the ones that remain. Their ids are what the board pills,
the cards and any exported PDF already show.

---

## What is local and what is not

Workbooks are **entirely local**. They are:

- never uploaded anywhere,
- never part of the hardware data the app syncs from `classic-repair-toolbox.dk`,
- stored beside your settings and log file, not inside the synced `Data` folder.

See [Where your data is stored](Workbooks-Where-your-data-is-stored) for the exact location and how to back
it up or move it.

---

**Next:** [Getting started](Workbooks-Getting-started) · [The Workbooks tab](Workbooks-The-Workbooks-tab)
