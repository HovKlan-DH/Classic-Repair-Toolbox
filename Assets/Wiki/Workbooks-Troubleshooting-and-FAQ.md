# Troubleshooting and FAQ

← [Back to Workbooks](Workbooks)

---

## The feature is missing

**I cannot see the Workbooks tab or the worklog bar.**
Tick **"Enable Workbooks tab"** in **Configuration → Workbooks**. Nothing is deleted while it is off
— your workbooks come back when you tick it again.

**The Workbooks tab is empty even though I have workbooks.**
Check the **Scope** setting under that checkbox. In "Show only workbooks for selected board" you see
only the current board's jobs — switch board, or choose "Show all workbooks".

Also check the **search box** is empty. When a search is what emptied the list it says *"No workbooks
or worklogs match your search"* rather than "none recorded".

---

## Workbooks and boards

**Can I move a workbook to a different board?**
No. A workbook is attached to the board that was selected when you created it, permanently. If you
picked the wrong board, export it if you want the record, then recreate it on the right one.

**Why does my board show no workbooks when I know I made one?**
It was probably created against a different board — an easy mistake when several revisions of the
same machine are in the dropdown. Switch Scope to **"Show all workbooks"**; each card then names its
board, and clicking one switches the app to that board.

**How do I get back to a job on a different machine?**
Use the workbook picker on the worklog bar. It lists **every** workbook on every board, and picking
one switches the app to its board first.

---

## Status and state

**How do I close a workbook?**
You do not — close its **worklogs**. A workbook closes itself once every worklog in it is Closed, and
reopens if you add or reopen one. See
[workbook status](Workbooks-Concepts-and-vocabulary#workbook-status).

**My workbook reopened by itself.**
You added a worklog to it, or reopened one, or deleted its only worklog (a workbook with no worklogs
is Open). All three are the rule working as intended.

**Can I still use a closed workbook?**
Yes. It stays in the list, can be activated, edited and exported, and adding work to it reopens it.
Nothing is archived or hidden.

---

## Worklogs on the board

**My worklog does not appear on the schematic.**
Check, in order:

1. **"Show worklogs"** is ticked on the worklog bar.
2. The **right workbook is active** — worklogs from other workbooks are not drawn.
3. You are on the **schematic the worklog was filed against**.
4. The worklog has **"Show marked area"** ticked. If not, it is not missing — its pill is **parked in
   the top-right corner** of the panel.

**A worklog's pill is in the corner instead of on the board.**
That worklog has "Show marked area" unticked, which is the normal state for one created from an
oscilloscope capture. Tick it in the editor and the app gives it a real, draggable square in the
bottom-right of the board to move where you want.

**I ticked "Show marked area" and got a square in the wrong place.**
That is the starting position for a worklog that never had an area — deliberately the opposite corner
from the parked pills, so it cannot be confused with one. Drag it where it belongs.

**I unticked and re-ticked, and my carefully placed area is still there.**
Correct — the app never moves an area you placed yourself. It only ever *adds* one when there was
none.

**The marker did not update after I edited a worklog.**
It should — every worklog change redraws the board overlay and the thumbnail pills, wherever you made
the change. If you are seeing stale markers, that is worth reporting.

---

## The editor

**The Save button is greyed out.**
The **Title** is empty. It is the one required field.

**I pressed Cancel and my photos are still there.**
On a worklog that already exists, list changes (comments, photos, files, reordering) are written the
moment you make them — Cancel only discards unsaved changes to the title, description, category and
state. This is deliberate, so a long session of attaching photos cannot be lost to a mis-click.

On a **new** worklog nothing is written until you save, and cancelling leaves no trace at all.

**The "Components in scope" section is not showing.**
The app could not work out what your area covers — usually because the board data has not finished
loading, or a region filter is hiding that schematic's components. Let the board load and reopen the
worklog.

Importantly, saving in that state **does not wipe** the components you had already put in scope.

**Why are there components in scope I did not pick?**
Everything your rectangle touched is ticked by default, on the assumption that you drew it around
them. A generous rectangle around a chip usually catches a couple of neighbouring passives — untick
them, or use the **None** button and pick.

---

## Searching

**Searching for "open" returns almost nothing useful.**
Open/Closed are deliberately **not searched** — every record carries one of two values, and "open" is
a word this domain uses constantly ("open circuit", "opened the case"). Filter by the pills instead;
they are always visible.

**Searching for a number does not find anything.**
Numbers are not searched — ids, hours, cost, dates. A search for `2` would otherwise match nearly
everything through fields you never see as text.

**A single letter matches everything.**
Matching is substring, so short terms are broad. `p c u` finding "CPU" is the same rule working for
you; a single letter is the same rule working against you. Use quotes for a phrase.

**My search vanished when I changed board.**
Deliberate — a board change is a change of subject, and carrying the filter across would land you on
an empty list with the reason in a box you are not looking at.

**Did searching change which workbook I am writing into?**
No, never. Search narrows what you see; it does not redirect where "Add worklog" writes. If the
active workbook is filtered out, the tab shows a different one temporarily and reverts when you clear
the box.

---

## Export

**Where did the ZIP button go?**
It is its own button beside "Export to PDF" in the header, on the second row of actions.

**I typed `repair.pdf` in the ZIP dialog and got `repair.zip`.**
Correct — the format comes from the button you pressed, not from what you type. An unrelated suffix
is kept: `board rev 2.5` becomes `board rev 2.5.zip`.

**The export did not open afterwards.**
By design. The app only opens local files from inside its own data folder, and an export is saved
wherever you chose.

**Why is the workbook's description not in the file name?**
It is a sentence you typed, often carrying a customer's own details, on a file that is about to be
emailed. The id, hardware, board and date identify it without that.

**A photo is missing from the export.**
Its file is not on disk — most likely the `worklog_{id}` folder was edited by hand. Missing files are
skipped rather than failing the export.

---

## Data and numbering

**There is a gap in my workbook numbers.**
Deliberate. Ids are never reused after a delete, because an id names an already-exported PDF and a
folder on disk. See [why numbers are never reused](Workbooks-Deleting-things#why-numbers-are-never-reused).

**How do I back up?**
Copy the `Workbooks` folder — that is the entire backup. **Configuration → "Open data/workbooks/log/
settings folder"** takes you there. See
[Where your data is stored](Workbooks-Where-your-data-is-stored).

**Are my workbooks uploaded anywhere?**
No. They are entirely local, never synced, and not part of the hardware data the app downloads.

**Do workbooks survive an app update?**
Yes. They live in a folder that updates do not touch.

**Can I keep workbooks on a shared drive?**
Launch with the `--workbooks-root=` switch. See
[Using a different folder](Workbooks-Where-your-data-is-stored#using-a-different-folder).

**My attachments disappeared after updating from an old version.**
Attachment folders used to be named `entry-{id}-files` and are now `worklog_{id}`. Old folders are
not renamed automatically. Renaming them by hand restores the attachments.

---

## Still stuck?

The app writes a log file beside the settings file — reachable from the same **Configuration →
"Open data/workbooks/log/settings folder"** button. Worklog operations are logged, including failed
writes, which is the first place to look if something did not save.

---

← [Back to Workbooks](Workbooks)
