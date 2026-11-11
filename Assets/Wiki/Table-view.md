[Wiki Home](Home)

Work on a board's data as a spreadsheet: every row of a part of the board in one table, with copy, paste, search and
replace.

---

On the "Contribute" tab, the switch **UI | Table** at the top shows the entry picked in the left column either as its
normal page (**UI**) or as a table (**Table**). The table is the same data, row by row, and every change you make in it
goes into your draft of the board - it saves itself a moment later, just as the normal pages do. The "Maintainer" tab
shows the same table when a maintainer looks at a submission or a board.

## Moving around and typing

| To do this | Do this |
| --- | --- |
| Change a cell | Double-click it, press **F2**, or just start typing |
| Finish the cell | **Enter** (moves down) or **Tab** (moves right; **Shift+Tab** moves left) |
| Undo the typing in a cell you are still in | **Ctrl+Z** inside the cell |
| Go to the first or last column of the row | **Home** / **End** |
| Go to the first or last row | **Ctrl+Home** / **Ctrl+End** |

Every cell wraps its text, so a long note is shown over several lines in its row.

## Selecting cells

* Click a cell to select it.
* **Shift+click** another cell to select everything between the two.
* **Ctrl+click** to add a cell to the selection, or take it away.

You always work with cells, not whole rows or columns: to take a "column", select its cells from the first to the last
row; to take a "row", select its cells from left to right.

## Copy and paste

* **Ctrl+C** copies the selected cells (or, with nothing selected, the cell with the dashed frame). They are copied as a
  spreadsheet block, so you can paste them into Excel or LibreOffice Calc too.
* **Ctrl+V** pastes. A block copied from here or from a spreadsheet lands with its top-left corner in the selected cell,
  and fills the cells to the right and down from there:
  * rows shown in red (deleted rows) are skipped;
  * when the block runs past the last row, new rows are added at the end - but not while the table is narrowed by a
    pill, a search or components picked on the left;
  * columns past the table's last column, and a heading row in what you copied, are left out.

  A line above the table says what went in, and what did not and why.
* Copy **one** value and select **several** cells before pasting: every selected cell gets that value.

Each paste is one step, so **Ctrl+Z** takes it all back at once.

On macOS use **Cmd** where this page says **Ctrl**.

## Find and replace

* **Ctrl+F** puts the cursor in the search box above the table. It shows only the rows it finds, and marks the text
  found. Several words must all be found; put words in "quotes" to find them together; put **-** before a word to leave
  out the rows that have it.
* **Ctrl+H**, or the **Replace** button, opens the replace bar under the toolbar. **Find** narrows the table and marks
  every place it finds, exactly as typed (capitals ignored); what you type under **Replace with** is shown in their
  place, in another color, so you see the result before anything changes. Under **Columns** you can tick which columns
  it works in. **Replace all** says how many places it will change. **Cancel** closes the bar.

Replace works only in the rows the table shows on the part you are looking at.

## Rows

| To do this | Do this |
| --- | --- |
| Add a row | Select a row and press **Insert row** - the new row goes below it |
| Delete rows | Select a cell in each and press **Delete row** - a row that is published stays in red where it was, until you submit |
| Move a row | Drag the grip at the far left of the row, or press **Alt+Up** / **Alt+Down** |

Rows cannot be moved while the table is narrowed by a pill or a search.

## What the colors and pills mean

| You see | It means |
| --- | --- |
| A green row (marked **+**) | A row that is not in the published data - you added it |
| An orange cell (row marked **~**) | You changed this cell. Point at it to see the published value |
| A red, struck-through row (marked **-**) | A published row your draft no longer has |
| A red triangle in a cell's corner | An error - it must be fixed before you can submit. Point at the cell to see why |
| An amber triangle in a cell's corner | A warning - worth a look, but it does not stop anything |

The pills above the table - **Added**, **Modified**, **Deleted**, **Errors**, **Warnings** - count those rows in the
whole draft. Click a pill to show only those rows, and click it again to show every row. Several pills can be picked at
once.

Point at a cell naming a file to see the file in a small card beside it, and at a changed color to see the color
before and after.

## Columns

* Drag the edge of a column's heading to make it wider or narrower.
* Double-click the edge of a heading to fit the column to the longest text in it.

## Undo and redo

**Ctrl+Z** undoes your last change, and **Ctrl+Y** (or **Ctrl+Shift+Z**) brings it back. On the "Contribute" tab this
is the same undo as the draft's own **Undo** and **Redo** buttons, so a change made on a normal page can be undone in
the table too, and the other way round.
