[Wiki Home](Home) · [Workbooks](Workbooks-tab)

The bar, the editor, and the markers on the board.

---

## The worklog bar

The strip above the tabs, visible on every tab.

| Control | Does |
| --- | --- |
| **Device** drop-down | **Show workbooks for all devices**, or one of your devices - each with its status and how many workbooks it has. Picking a device goes to its repair, and while it is picked you work only in its workbooks - see [Devices](Workbooks-Devices) |
| **Workbook** drop-down | Which workbook you are working in - `#3 \| No video`, with Open or Closed at the end of the line |
| **Show worklogs** | Draws the markers on the schematic |
| **Add worklog** | Start marking an area (see below) |
| **Cancel** | Shown while you are marking an area - backs out |
| **Add workbook** | New workbook on the selected board |
| **Add device** | Add a computer you own - see [Devices](Workbooks-Devices) |

The **Workbook** drop-down lists every workbook, on every board - or, with a device picked, that device's workbooks. One on another board names that board, `#3 | No video (C64/250407)`, and picking it switches the application to that board first.

The device picked is the same choice as the device selected under "My devices" on the "Workbooks" tab - picking it in one place picks it in the other. CRT remembers it. **While a device is picked, you work only in its workbooks**: with none of them on the board shown, the **Workbook** drop-down says so and **Add worklog** is hidden until **Add workbook** makes one. Working in a workbook of another device picks that device; a workbook on no device - or switching to a board with none of the device's workbooks - puts the drop-down back on **Show workbooks for all devices**.

Both drop-downs are as wide as their text. When the window is too narrow for both, the shorter one keeps its text whole and the longer one is cut short.

After **Clear** under the workbooks on the "Workbooks" tab, the drop-down reads "No workbook picked" and **Add worklog** is hidden until you pick one.

On a board with no workbook yet, the drop-down reads "Select a workbook" when you have workbooks on other boards to pick from, and **Show worklogs** and **Add worklog** stay hidden until you pick or create one. Until you have made your first workbook, there is no **Workbook** drop-down at all, and until you have added your first device there is no **Device** drop-down - only the buttons to add them.

## Adding a worklog

**Add worklog** -> drag a rectangle on the schematic -> the editor opens.

If you want to back out, click **Cancel** in the bar, or just press `Esc`.

## The editor

Opens when you create a worklog, or when you click a `#N` marker, a pill or a card.

Top of the window:

* The title box ("Worklog title") - required
* **Category** - Note / Cosmetic / Issue
* **State** - Open / Closed
* **Description**
* **Show marked area on schematics image** - see below

On the right, "Worklog location" shows the schematic with the worklog's area marked on it.

Below that, seven lists you fill in as the repair goes on: Links of interest, Work done, Comments, Components in scope, Components completed, Photos, Files.

* Click a list's heading to fold it away. Each worklog remembers which lists you folded.
* **Work done** and **Comments** have two sort buttons beside their heading: newest first or oldest first.
* Click a Work done or Comments row to edit it. Links, Photos and Files rows have their own edit and delete buttons.
* Click a link row to open the link, a photo to view it full size, or a file name to open the file.
* In the photo viewer, the mouse wheel zooms in and out where the pointer is. A drag with the right or middle mouse button moves the zoomed photo - as on the "Schematics" tab - and a double-click shows it whole again.

**Work done** takes time and cost per line and totals them. The Time spent field is the one place in the application where time is typed as decimal hours, because that is what a quarter of an hour is easiest to enter as. As you type it is read back underneath in hours and minutes - `1,25` shows as **1** hour and **15** minutes - so a mistyped decimal shows itself straight away rather than once it is in the totals (`1,4` is an hour and **24** minutes, not an hour and forty).

**Everywhere the application shows you a time, it says it in hours and minutes** - the work done lines themselves, the totals beside the section heading, the worklog cards, the totals strip and the exported PDF all read `45 minutes` or `1 hour and 15 minutes`, never `0,75 h`.

**A figure with nothing to report is left out rather than shown as a zero**, for the time and the cost alike. A work done line with time but no cost shows only the time; one with neither shows just its note.

The cost field names the currency you have chosen ("Cost (DKK)"), and every total is shown with that currency code beside it. Pick the country you live in under **Configuration -> Workbooks -> Currency used for worklog costs**; it defaults to United States (USD).

Each work done line keeps the currency it was entered in, and shows it. The totals simply add the numbers up and show them in the currency chosen now - so if you change the currency part-way through a repair, the totals mix the two.

**Photos** and **Files** can be dragged up and down; that order is the order they appear in the PDF.

Ctrl+Enter saves.

### Automatic comments

The application writes a few comments into the Comments list by itself, so a worklog keeps its own history: "Worklog created" when it is made, "Worklog opened" or "Worklog closed" when you change its state, and `Worklog changed to "Issue"` (or whichever category) when you change its category.

### Links in your text

A web link starting with `http://`, `https://` or `www.` can be clicked when it is in a workbook's note, a worklog's description, a work done line, a comment, or the comment on a photo or file. Titles are never links. The same links are clickable in the exported PDF.

### Formatting your text

Every box you write a longer text in - a workbook's note, a worklog's description, a work done line, a comment, the comment on a photo or file, and a device's description - has a small row of buttons at its top: **bold**, *italic*, underline, ~~strikethrough~~ and a bullet list. You see the formatting in the box as you type.

| Keys | Does |
| --- | --- |
| Ctrl+B, Ctrl+I, Ctrl+U | Bold, italic, underline - on the selected text, or, with nothing selected, on what you type next |
| Ctrl+Shift+X | Strikethrough |
| Ctrl+Shift+8 | Bullets on or off for the lines selected |
| Tab / Shift+Tab | On a bullet: one level deeper / one level back |
| Enter | On a bullet: the next bullet. On an empty bullet: ends the list |
| Ctrl+Z / Ctrl+Y | Undo / redo, the formatting included |

The text is stored as you would write it by hand: `**bold**`, `*italic*`, `__underline__`, `~~strikethrough~~`, and a line starting with `- ` (two spaces more for each level) for a bullet. The formatting is shown on the cards, the device page and in the exported PDF, and the search looks in the words, not the marks.

### What Cancel does

| | Cancel |
| --- | --- |
| New worklog | Nothing is written at all |
| Existing worklog | Takes back only what has not been saved yet - see below |

On an existing worklog, a lot is saved the moment you do it:

* Picking a category or a state is saved at once, together with its automatic comment.
* Adding, editing, deleting or dragging anything in Links of interest, Work done, Comments, Photos and Files is saved at once.
* Each of those saves also writes the title, description, **Show marked area on schematics image** and the component ticks as they are at that moment.

So Cancel only takes back changes to the title, description, **Show marked area on schematics image** and the component ticks made since the last of those saves.

**On a NEW worklog nothing is written until you click "Add worklog"** - including its comments and photos. That is deliberate, so backing out leaves nothing behind, but it is the opposite of how the same actions behave on an existing worklog. So if you close a new worklog that has anything in it - by Cancel, Escape, or the window's ✕ - you are asked to confirm first. Choose **Keep editing** to go back and save it.

## Markers on the board

With **Show worklogs** ticked, each worklog is drawn on its schematic as a dashed rectangle in its category color, with a `#N` marker. The thumbnails show the `#N` markers too.

Click a marker on the schematic to open that worklog.

### Moving and resizing a marked area

With **Show worklogs** ticked, point at a worklog's rectangle on the "Schematics" tab:

* Drag a corner or a side to resize it.
* Drag inside it to move it.

It is saved the moment you let go - no need to open the editor.

Making an area smaller removes the components it no longer covers from "Components in scope". Making it larger never adds any - tick those in the editor.

Where a component sits inside the rectangle, pressing there selects the component instead, so grab the area by its edge or by an empty part of it.

### Worklogs without an area

Untick **Show marked area on schematics image**, and the worklog gets no rectangle - its marker parks in the top-right corner of the panel instead.

Use it for things that are not about a place on the board: "fault is intermittent, only when warm", "case scratched".

Tick it again and the worklog gets a square in the bottom-right corner of the board, which you then move and resize into place as described above. A worklog that already has an area keeps it.

## Components in scope

Two checklists in the editor:

* **Components in scope** - which components the worklog is about. Everything your rectangle touched starts ticked.
* **Components completed** - which of those you have finished. Useful for a "replace every electrolytic" pass.

Both have **All** and **None** buttons.

If the section is missing entirely, the board data has not finished loading, or a region filter is hiding the components - reopen the worklog once the board is loaded. Your saved components are not lost in the meantime.

## Filing an oscilloscope capture

After a capture, the "Saved image as [...]" banner in the component popup has an **Attach image to worklog** button. It only appears when the "Workbooks" tab is enabled and the board has a workbook, and the image goes into the workbook shown in the bar.

The dialog shows the image and the workbook, asks which worklog, and takes a comment. Worklogs that already have the component you were probing in scope are listed first, under "Worklogs with [U8] in scope", then "All other worklogs". The first one is preselected - so probing U8 while working a fault on U8 is one click on **Attach to existing worklog**.

**Create new worklog** is always last in the list; with it chosen the button reads **Create worklog**. It opens the editor with the image already attached, and the new worklog is filed against the schematic you have open, with no marked area.

The image is **moved**, not copied: once it is filed, it lives in the worklog's own folder and is no longer in the oscilloscope image folder, so you do not end up with two of every capture. If you cancel the new worklog instead, nothing is filed and the capture stays where it was.

---

**Next:** [Browsing and search](Workbooks-Browsing-and-search) - find an older repair, and read the totals
