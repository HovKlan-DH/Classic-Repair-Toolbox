[Wiki Home](Home)

Fix a value or add a datasheet from inside the application.

---

Spotted a wrong value, a missing part number, or do you have a datasheet, photo or scope
reading that others would benefit from? You can add it directly from the **Contribute**
tab - no GitHub account, no spreadsheets, no technical knowledge needed.

Your edit is saved into your own local **draft** for this board and takes effect on your own
machine right away — untick "View boards as officially published" in the Configuration tab at any
point to check what your edit looks like. Nothing changes for anyone else until you submit your
draft and it has been reviewed and accepted, after which it reaches everyone the next time the
application syncs its data.

## 1. Pick your board

At the top of the main window, select the **hardware** and the **board revision** you want to
improve. The "Contribute" tab always shows the board that is currently selected.

## 2. Open the "Contribute" tab

You will see all the components of the board, grouped in columns by category (ICs, capacitors,
resistors and so on). Hover a label to see its full name.

The line **"Board Excel data last revisioned"** tells you how fresh the current data is.

## 3. Click the component you want to change

Click a board label (for example `U1` or `C15`) and the **Component editor**
opens in full screen. Everything in it is already filled in with the data the app has today —
you are simply correcting or extending it.

## 4. Make your changes

The editor is divided into collapsible sections. Click a heading to open it. The number in
brackets tells you how many entries it already holds.

| Section | What belongs here |
| --- | --- |
| **Component** | The basics: friendly name, technical name or value, part number, category, short description |
| **Component images** | Photos and oscilloscope readings for this component, with the scope settings used |
| **Component local files** | Files about this component — datasheets, instructions, manuals |
| **Component links** | Web links about this component |
| **Board local files** | Files covering the whole board — service and troubleshooting manuals |
| **Board links** | Web links about the whole board — repair logs, YouTube videos |

Then:

- **To change something** — just type in the box. Correct a wrong value, fill in a blank one.
- **To add something** — click the **Add new …** button at the top of the section. A blank row
  appears; fill it in.
- **To delete something** — click the **Remove** button on that row.

### Attaching a file or an image

In any row that has a **File** box, click the box and a normal file browser opens. Pick the
file from anywhere on your computer — a photo, a datasheet, a screenshot. The file name appears
in the box, images show a small preview, and a copy of the file is stored inside your local draft.

The **File location** dropdown next to it says which folder the file should end up in once
published. Leave it as it is unless you have a reason to change it.

A system you added yourself is in that list too. Its `Scope baseline` folder is where scope
images from a known good board go.

> **Can't find your component in the list?** Pick the closest one and simply explain in the
> note (next step) what should be added, or use **"Add new component"** on the Contribute tab
> instead to start one from scratch.

## 5. Write a note (optional)

At the bottom, you can describe **what you changed, what was wrong or missing**, and anything
worth remembering when you come back to this draft or submit it later. Worth mentioning **which
exact board revision you have** — boards vary, and that detail matters at review time.

This is entirely optional — it is a note to yourself and, later, to whoever reviews your
submission, not a requirement to save.

## 6. Save

Click **Save to draft**. Once it has been saved, the editor closes and you are taken to the
**Drafts** tab, where your draft now is - and where you go on to review it and submit it.

If something is missing or goes wrong, the editor stays open and its message says what, so you
can put it right and save again.

**If the table on the Drafts tab has unsaved edits for the same board**, nothing is saved: a
message asks you to go to the Drafts tab, save or discard the table's edits there, and then come
back and save this again. Both would write the same draft, and saving here first would make the
table's edits impossible to save. The editor stays open with everything you typed. The same goes
for saving new component labels on the Schematics tab.

Use **Cancel** to close the editor without saving whatever you have not saved yet. The **▲** and
**▼** buttons jump to the top and bottom of a long form.

## Adding a board that is not in the list at all

Everything above assumes the board you want to improve is already there. If it is not, the
**Add a new system** button on the Contribute tab creates one: give it a manufacturer, hardware and
board name, and it appears in the lists immediately as your own local draft. From that point it is an
ordinary board in every respect — you add board images, label the components and attach files with
exactly the same tools described above.

Before it is created you are asked to accept the role of **maintainer** for it. If you later submit
the system for the community to use, you will review the changes others submit for it and publish
the ones that are right.

The full walkthrough, including importing a KiCad project so that clicking a component lights up its
real copper traces, is on [Add new board with KiCad data](Add-new-board-with-KiCad-data).

## Worth knowing

Your edit takes effect in your own local view of the board as soon as you save it — the "Drafts"
tab lists every board you have local changes on, and you can discard a draft there at any time.
If one of the draft's files is open in another program - its workbook in Excel, say - Discard
cannot remove all of it and says so at the top of the tab. Close the file there and press Discard
again.
Nothing changes for anyone else until you submit your draft and it is reviewed and accepted, after
which it reaches everyone the next time the application syncs its data.

**Where a new component appears.** However you add one - the Contribute window, labelling it on a
schematic, or the table below - a new component goes into its category, in label order (C1, C2,
C10, not C1, C10, C2). A component in a category the board does not use yet goes at the end, and
another region of a component that is already there goes right beside it. Editing an existing
component never moves it.

**A draft you never submit is a perfectly good thing to have.** A board you built yourself keeps
working on your own machine for as long as you want it — it is not a staging area you are expected
to empty.

## When the official data changes underneath your draft

Board data is improved by other people too, so the official version of a board you are working on
can be updated while your draft sits on top of it. When that happens the "Drafts" tab marks that
board **"Updated officially"**, and the board itself shows a one-line notice.

**Your edits are not lost and nothing is overwritten.** Your draft and the synced data live in
separate folders; your changes are still applied on top of whatever the official data now says.

Click **"What changed"** to see how your own edits line up with the official data as it stands now.
Two things there are worth acting on:

* **an edit whose component no longer exists officially** — your change has stopped having any
  effect, because the row it was changing is gone;
* **something you added that has since been added officially too** — the official version is the
  one being shown, not yours.

Everything else is listed for completeness.

> [!NOTE]
> CRT cannot show you the official data *as it was* when you started, because that copy is replaced
> when the data syncs. What it shows is how your edits stand against the data as it is now.

When you have had a look and are happy, click **"I have looked at this"** and the notice goes away.
That writes down which version you have seen and **changes none of your edits** — it is not a merge,
and nothing of yours is discarded. If the official data changes again later, you will be told again.

## Editing a draft as a table

If you would rather work the way you would in Excel - for example because you are copying a lot of
values over from somewhere else - click **"Edit in table format"** on the board's row in the
**"Drafts"** tab. The board's data opens right below that row, with one tab for each sheet of its
Excel file (Board schematics, Components, Component images and so on) and every column and row in
it. The other drafts are hidden while the table is open, and the button changes to
**"Close table"**.

Everything that differs from the official data is coloured:

* **green** - a row you added;
* **orange** - a value you changed. Only the changed cell is coloured, and hovering over it shows
  the official value;
* **red, and struck through** - a row you deleted. It is shown where it used to be, so you can see
  what was around it;
* **violet** - a row worth a second look, marked `!`. Either another row above it describes the
  same thing (two rows for the same component and region, say), or it is missing a value it needs
  and will be left out when saving. Hover over the row to see which.

The colour key above the table counts each kind for the sheet you are looking at - "2 Added",
"1 Modified" and so on, with a kind that has none shown faded - and ticking **"Show changes only"** hides every row you have not touched,
so you can check exactly what you did. Violet rows stay visible when it is ticked.

Each row also starts with a small sign that says the same thing without relying on colour: `+`
added, `~` changed, `-` deleted, `!` worth a second look. The number on each sheet's tab is how
many rows were added, changed or deleted on that sheet; violet rows are not counted there, because
they are not changes. For a board you created yourself nothing is marked as added, changed or
deleted, since all of it is your own - but violet rows are still shown.

**Editing works much as it does in Excel.** Click a cell to select it - it gets a dashed red
frame - then just start typing, double-click it, or press F2. **Tab** moves to the next cell to the right (and on to the next row
at the end of one), **Shift+Tab** to the left, and **Enter** down. **Insert row above** and
**Insert row below** add an empty row next to the selected one, and **Delete row** removes the
selected row. Changed your mind? Ctrl+Z, below.

**Deleting a component deletes everything that belongs to it.** Its rows on the *Component
images*, *Component local files* and *Component links* sheets go with it - you will see them in red
on those sheets - and its highlights on the schematics go when you save. A line under the table
says what else went. If the component has another row with the same board label for a different
region, only the images for the region you deleted go, since the rest still belong to the other
one. One Ctrl+Z brings all of it back.

**Ctrl+Z undoes, Ctrl+Y redoes** (Ctrl+Shift+Z works too, and Cmd on a Mac). That covers every
change in the table - typing into a cell, pasting, and inserting, deleting, restoring or moving
rows - one step at a time, and takes you to the sheet and row where the change was. It reaches
back to your last save. While you are still typing inside a cell, Ctrl+Z undoes just that typing,
as in Excel. For something saved earlier, hover over an orange cell to see its official value, and
a red row still shows everything it held - type or copy the values back in.

**Rows can be moved.** Drag a row by the dotted handle at its far left: while you drag, the row
turns into a red dashed empty slot that moves with the mouse, and the other rows make room, so you
can see exactly where it will land - let go of the mouse button there. Dragging above or below the
rows you can see carries it further, one row per move. Alt+Up / Alt+Down move the selected row
one place from the keyboard. A red (deleted) row cannot be moved, and while **"Show changes only"**
is ticked no row can. The order matters: it is the order the
components appear in on the left of the main window. A moved row is not coloured, because its
content has not changed - see the note below about order on its own.

**A new component is put in its place for you.** When you save, a component row you inserted goes
into its category, in label order - C1, C2, C10 - just as components added any other way do (see
below). If you moved it yourself before saving, it stays where you put it. After saving, the cursor
is on the row wherever it went.

**One component, several regions.** A component that differs between regions has one row per region
with the same board label - for example U1 with region PAL and U1 with region NTSC. Each is its own
row: adding the NTSC one shows it as added, and it is saved right beside its PAL twin. Giving an
existing component a region makes it a new row in the same way as changing its label would - the
old one shows as deleted and the new one as added.

**Copy and paste one cell at a time.** Ctrl+C and Ctrl+V work on a selected cell and inside a cell
you are editing, and cells copied from Excel paste in fine. Pasting a whole block of several cells at
once is not supported yet - copy them one by one.

**Nothing is written until you click "Save changes".** If you close the table, submit, add schematic
images or KiCad data, or quit CRT while the table has edits that are not saved, you are asked
whether to save or discard them first.

The table and the board's Excel file are the same data, so you can switch between the two freely.
While the table is open, CRT keeps an eye on the Excel file:

* **When the file is saved in Excel** (or anywhere else outside the table), the table notices within
  a couple of seconds. If you have nothing unsaved in the table, it simply reloads and says so - and
  the draft's "rows changed" and the board on the other tabs follow. If
  you do, a yellow bar says the table is out of date and your unsaved edits can no longer be saved
  into the draft - CRT will not save the table over those changes, so **"Save changes"** is greyed
  out. Copy anything you need, then press **"Reload"** on the bar. Closing the table then asks only
  whether to discard its unsaved edits. It never reloads while you are typing in a cell or dragging
  a row.
* **When the file is open in Excel**, a yellow bar says so for as long as it is, from the moment you
  open the table: edit it in one place at a time. The table follows what you save in Excel, but
  cannot save its own edits until the file is closed there, so **"Save changes"** is greyed out
  until then - and closing the table offers only to discard its edits or cancel. (If Excel crashed,
  the bar can stay after it is gone - saving works again then.)

CRT cannot see edits you have made in Excel until Excel saves them.

The rectangles that mark components on the schematics are not part of the table - they are not in
the Excel file either - and saving the table leaves them untouched.

> [!NOTE]
> Changing only the ORDER of rows is saved into your draft, but it does not count as a change on its
> own: the "rows changed" count stays the same, and a draft whose only change is a new order cannot be
> submitted by itself. The new order is sent along with any other change you submit.

## Sending your work in

When a draft is ready, open the **"Drafts"** tab and click **"Submit"** on that board's row.

**You do not need an account, and CRT will not ask you to make one.** The only thing the dialog
asks for is an email address, and it is used for one purpose: telling you whether your
contribution was accepted, and why if it was not. You also write a short summary of what you
changed, which is what the maintainer reads first.

Before anything is sent, the dialog shows you exactly what is about to go: how many schematics,
components and highlights, and how many files are referenced. Nothing leaves your machine until
you click Submit.

**Files you already downloaded are not uploaded again.** CRT asks the server which files it is
missing and sends only those, so correcting a typo on a board with hundreds of images uploads
almost nothing. A large new board with its own images will take longer, and the dialog shows
progress while it works. You can cancel at any point; a cancelled submission leaves nothing behind
on the server and does not touch your draft.

**Your draft stays exactly where it is after you submit.** It is not cleared, emptied or locked,
and you can carry on using the board as normal while the contribution waits to be looked at. That
is deliberate: review takes time, and you should not lose the use of your own work while it
happens.

**Once your contribution is published, the draft tidies itself away.** When the application has
downloaded the updated data and your draft holds nothing the published board does not - no other
rows, no other KiCad calibration and no other files - the draft is removed on its own, the next time
the application starts or when you close "My submissions". If you kept working in the draft after
submitting, or it is open in the table editor with unsaved edits, it stays, and nothing of yours is
lost. **A whole new system is tidied away the same way**, once CRT lists the published system in
its hardware and board lists. Until then your draft is the only place you can see it, so it stays.

If the submission cannot be sent - no internet connection, or a file the data refers to has been
moved or deleted since you added it - you are told which file and what went wrong. Nothing is lost
and your draft is unchanged.

**A few things are refused before anything is uploaded**, each with a message saying which file
and why:

* **Only the kinds of file boards actually use can be sent** - pictures (PNG, JPG, GIF, BMP,
  WebP), PDF documents, plain text and web pages (HTML). A file of any other kind, or a hidden file
  whose name starts with a dot, is refused.
* **A file must really be what its name says.** A ".png" has to contain a picture and a ".pdf" a
  PDF document; a text file has to be plain text, saved as UTF-8.
* **A board can only change its own files and the shared folders.** You can use a file that belongs
  to another board, exactly as it is, but you cannot change it from here - change it through that
  board's own draft instead.
* **Every file sent must be used by the board.** CRT only ever sends the files your board refers
  to, so this only matters if a file was added by hand.
* **There is a daily limit per internet connection**, generous enough that it only stops something
  going badly wrong. If you reach it, the message says when you can try again, and your draft is
  kept.

## Checking how a contribution is getting on

The **"My submissions"** button on the Drafts tab lists what you have sent, with the state of each
one and anything the maintainer has said. **"Check for updates"** asks the server for the latest.

**An accepted contribution is published in two steps.** First it is published to the **BETA
source**, where a maintainer gives the board a final check; "My submissions" then says *Published to
BETA source*. After that it is published to the ordinary **source** that everyone downloads from, and
the row says *Published to source*. You get an email at each step. CRT looks for the second step
each time it starts, for a month after the first; after that, "Check for updates" still asks.

**A maintainer may correct small things before publishing** - a typo, a wrong part number - rather
than sending the whole contribution back to you. When that happens, "My submissions" and the email
both say that a maintainer changed some of the details, so what is published is not exactly what you
sent. Your own draft is not changed, so it still holds what you sent and is not removed
automatically - once your data has updated, look at the board, then discard the draft on the
Drafts tab (or keep working from it).

A few things worth knowing about that list:

* **It is kept on this computer.** Because contributing needs no account, there is nothing that
  could make the list follow you to another machine - so a contribution sent from a different
  computer, or sent before a reinstall, will not be listed. **You are still told the outcome by
  email either way**, and that is the part that does not depend on this list.
* **"Remove" does not withdraw anything.** It only takes the row off the list on this computer.
  The contribution has already been sent, it will still be reviewed, and you will still be emailed
  about it. What you lose is the ability to check on it from inside CRT.
* Contributions that have already been decided are not re-checked, so pressing "Check for updates"
  when everything is settled will simply tell you there is nothing left to ask about.
* With no internet connection the list still opens and shows what it last knew, rather than going
  blank.

### If you contributed a whole new system

For a brand new hardware and board, the project owner may set you up as a **maintainer** of that
system once it is published, so you can look after it from then on - reviewing what others send in
for it, and publishing what is right. That happens after the fact and only for new systems; it is
never something you need before contributing.

## That's it

Thank you — every correction makes the data better for the next person repairing the same board.
