[Wiki Home](Home) · [Workbooks](Workbooks-tab)

Turn it on, and record your first repair.

---

From nothing to a recorded repair.

## 1. Turn it on

"Configuration" tab → tick **Enable Workbooks tab**.

You now have a "Workbooks" tab, and a bar above the tabs.

Under the checkbox you can also choose whether the Workbooks tab shows all workbooks, or only those for the selected board. Start with the default (selected board only).

## 2. Select the board

Pick hardware and board in the drop-downs, as you normally would.

The workbook belongs to whichever board is selected when you create it, and it cannot be moved later - so get this right first.

## 3. Create the workbook

Click **Create new workbook** in the bar.

* **Description** - required. The repair, in one line: `Dead C64, no video - top attic find`
* **Note** - optional. Anything else: how the fault showed itself, where the board came from, etc.

Click **Create workbook** (or Ctrl+Enter).

One workbook per repair - not per fault. Three faults on the same board is one workbook with three worklogs.

## 4. Record a worklog

1. Go to the "Schematics" tab and open the schematic you want.
2. Click **Add worklog** in the bar.
3. Drag a rectangle on the schematic around the area - a chip, a section, a connector.
4. The editor opens. Type a **Title** - Save stays greyed out until you do.
5. Click **Add worklog**.

The components your rectangle touched are already ticked under "Mark components in scope". Untick any that got caught by accident.

Cancel instead, and nothing at all is written.

## 5. Work through it

Click the `#1` marker on the board to reopen the worklog, and add as you go:

* **Work done** - a note with time and cost, which is totalled for you (time is typed as decimal hours and shown back everywhere as hours and minutes)
* **Photos** - before/after, with a comment each
* **Files** - datasheets, receipts, scope captures
* **Comments** and **Links**

Set the worklog to **Closed** when it is done. When the last one closes, so does the workbook.

## 6. Export it

"Workbooks" tab → **Export to PDF** for the write-up on its own, or **Export to ZIP** to include the original photos and files.

See [Export and your data](Workbooks-Export-and-data).

---

**Next:** [Daily use](Workbooks-Daily-use) — the bar, the editor, and markers on the board
