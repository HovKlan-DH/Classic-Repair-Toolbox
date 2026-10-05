[Wiki Home](Home)

Fix a value or add a datasheet from inside the application.

---

Spotted a wrong value, a missing part number, or do you have a datasheet, photo or scope
reading that others would benefit from? You can add it directly from the **Contribute**
tab - no account, no spreadsheets, no technical knowledge needed.

Your edit is saved into your own local **draft** for this board and takes effect on your own
machine right away - tick "View boards as officially coming from online source (hide my local draft
changes)" in the Configuration tab at any point to see the board without your edit
([View boards from online source](View-boards-from-online-source)). Nothing changes for anyone else
until you submit your draft from the [Drafts tab](Drafts-tab) and a maintainer has accepted it. It is
then published to the BETA source first, and after a final check to the stable source that everyone
downloads from.

## 1. Pick your board

At the top of the main window, select the **hardware** and the **board revision** you want to
improve. The "Contribute" tab always shows the board that is currently selected.

## 2. Open the "Contribute" tab

You will see all the components of the board, grouped in columns by category (ICs, capacitors,
resistors and so on). Hover a label to see its full name.

The line **"Board Excel data last revisioned"** tells you how fresh the current data is.

## 3. Click the component you want to change

Click a board label (for example `U1` or `C15`) and the **Component editor**
opens in full screen. Everything in it is already filled in with the data the app has today -
you are simply correcting or extending it.

## 4. Make your changes

The editor is divided into collapsible sections. Click a heading to open it. The number in
brackets tells you how many entries it already holds.

| Section | What belongs here |
| --- | --- |
| **Component** | The basics: board label, category, friendly name, region, technical name or value, short description, part number - and the **Delete this component** button |
| **Component images** | Photos and oscilloscope readings for this component, with the scope settings used |
| **Component local files** | Files about this component - datasheets, instructions, manuals |
| **Component links** | Web links about this component |
| **Board local files** | Files covering the whole board - service and troubleshooting manuals |
| **Board links** | Web links about the whole board - repair logs, YouTube videos |

Then:

- **To change something** - just type in the box. Correct a wrong value, fill in a blank one.
- **To add something** - click the add button at the top of the section:
  **Add new component image**, **Add new component file**, **Add new component link**,
  **Add new board file** or **Add new board link**. A blank row appears; fill it in.
- **To delete something** - click the **Remove** button on that row (**Remove image** on a
  component image).
- **To delete the whole component** - open the **Component** section and click
  **Delete this component**. Its images, files and links go with it.

### Attaching a file or an image

In any row that has a **File** box, click the box (or **Browse**) and a normal file browser opens. Pick the
file from anywhere on your computer - a photo, a datasheet, a screenshot. The file name appears
in the box, images show a small preview, and a copy of the file is stored inside your local draft.

A **Component images** row needs either an image file or a note. A row with only a note is fine -
many boards have a "Pinout" row that just gives a compatible part number. A row with neither
is marked red, and **Save to draft** will not save until you fill in one of them or remove the row.

The **File location** dropdown next to it says which folder the file should end up in once
published. It only lists the folders a file for this board may go in:

- the board's own folder and the folders inside it
- `Shared files` under the board's manufacturer - files all of that manufacturer's boards can use
- `Generic shared files` - files every board can use

A file you pick from elsewhere on your computer starts with **no folder chosen, so pick one**.
Until you do, **Save to draft** will not save, and the **File location** box turns red. A file
you pick from CRT's own data folder keeps the folder it is already in.

A system you added yourself is in that list too. Its `Scope baseline` folder is where scope
images from a known good board go.

> **Can't find your component in the list?** Use **Add new component** on the Contribute tab to
> start one from scratch.

## 5. Save

Click **Save to draft**. Once it has been saved, the editor closes and you are taken to the
[Drafts tab](Drafts-tab), where your draft now is - and where you go on to review it and submit it.
You explain what you changed when you submit, in **What did you change?** - that is what the
maintainer reads first, and the place to say **which exact board revision you have** (boards vary,
and that detail matters at review time).

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
board name (and, if you like, notes about the hardware for the maintainer who adds it to CRT's lists), choose **Create system**, and it appears in
the lists immediately as your own local draft, with CRT on the Drafts tab. From that point it is an
ordinary board in every respect - you add board images, label the components and attach files with
exactly the same tools described above.

Before it is created you are asked to accept the role of **maintainer** for it ("Becoming the
system's maintainer" - **Accept and create** or **Decline**). If you later submit the system for the
community to use, you agree to review the changes others submit for it and publish the ones that are
right. Accepting sends nothing by itself: the administrator invites a system's maintainers by email -
see [Maintainer tab](Maintainer-tab#getting-a-maintainer-account).

The full walkthrough, including importing a KiCad project so that clicking a component lights up its
real copper traces, is on [Add new board with KiCad data](Add-new-board-with-KiCad-data).

## Bringing in a board you already have work on

If you already have a board's files from earlier work - a board you set up the old way, or one you
edited in Excel or on another computer - you can put its folder into your drafts folder yourself:

* Windows: `%LocalAppData%\Classic-Repair-Toolbox\Drafts`
* Linux: `~/.local/share/Classic-Repair-Toolbox/Drafts`
* macOS: `~/Library/Application Support/Classic-Repair-Toolbox/Drafts`

Keep the same three folder levels the data uses - manufacturer, hardware, board - for example
`Drafts\Commodore\C128\310378 Open128\`, with the board's Excel file and everything that belongs to
it inside. The next time CRT starts, it takes the folder in as a draft and shows it on the
**Drafts** tab, ready to edit and submit like any other.

* A folder for a board CRT already lists becomes your draft of that board, and the Drafts tab shows
  what you changed compared with the published version.
* Any other folder becomes a new system, named after its folders. A board you set up the old way
  (with a `_UserContribution` workbook) is one of these, because it has never been published - it
  stays where it already is in the hardware and board lists.
* The board folder should hold one Excel file. For a board CRT already lists, a file with a
  different name is renamed, together with its `.json` file, to the name that board uses.
* There is no need to copy a manufacturer's `Shared files` folder in as well. CRT reads shared files
  from its own data, and a `Shared files` folder placed next to your board folders (for example
  `Drafts\Commodore\Shared files`) is ignored. A NEW or changed shared file belongs inside the
  board's own folder, under its full path - for example
  `Drafts\Commodore\C128\310378 Open128\Commodore\Shared files\Component images\` - which is also
  where CRT puts one you add through the Contribute tab.
* A folder CRT cannot take in is left exactly as it is, and the reason is written to the log file.
  The usual reasons are several Excel files in one board folder, or the Excel file being open in
  Excel while CRT starts.

## Worth knowing

Your edit takes effect in your own local view of the board as soon as you save it - the
[Drafts tab](Drafts-tab) lists every board you have local changes on, and you can discard a draft
there at any time.
If one of the draft's files is open in another program - its workbook in Excel, say - Discard
cannot remove all of it and says so at the top of the tab. Close the file there and press Discard
again.
If you discard a draft while something you sent from it is still being reviewed, the dialog says so,
and the board's maintainers are told that you discarded your draft - so they can check with you
before publishing it. Your contribution itself is not withdrawn.
Nothing changes for anyone else until you submit your draft and it is reviewed and accepted - and
it reaches everyone once it has been published to the stable source.

**Where a new component appears.** However you add one - the Contribute window, labelling it on a
schematic, or the table below - a new component goes into its category, in label order (C1, C2,
C10, not C1, C10, C2). A component in a category the board does not use yet goes at the end, and
another region of a component that is already there goes right beside it. Editing an existing
component never moves it.

**A draft you never submit is a perfectly good thing to have.** A board you built yourself keeps
working on your own machine for as long as you want it - it is not a staging area you are expected
to empty.

## When the official data changes underneath your draft

Board data is improved by other people too, so the official version of a board you are working on
can be updated while your draft sits on top of it. When that happens the "Drafts" tab says so under
that board, and the board itself shows a one-line notice.

**Your edits are not lost and nothing is overwritten.** Your draft and the synced data live in
separate folders; your changes are still applied on top of whatever the official data now says.

Click **"What changed"** to open "What changed officially". It shows the revision you started from
("You started from:") and the one published now ("Officially now:"), and then every change of
yours compared with the official data as it is now, one per line:

* "You added this. It is not in the published data."
* "You removed this. It is still in the published data, and your submission would remove it."
* "You changed:" followed by the columns you changed

Look through it for anything the official update already does differently, or no longer needs.

> [!NOTE]
> CRT cannot show you the official data *as it was* when you started, because that copy is replaced
> when the data syncs. What it shows is how your edits stand against the data as it is now.

When you have had a look and are happy, click **"I have looked at this"** and the notice goes away.
That writes down which version you have seen and **changes none of your edits** - it is not a merge,
and nothing of yours is discarded. If the official data changes again later, you will be told again.

## Editing a draft as a table

If you would rather work the way you would in Excel - for example because you are copying a lot of
values over from somewhere else - click **"Edit in table format"** on the board's row in the
[Drafts tab](Drafts-tab#the-table). The board's data opens right below that row, with one tab for each sheet of its
Excel file (Board schematics, Components, Component images and so on) and every column and row in
it. The other drafts are hidden while the table is open, and the button changes to
**"Close table"**.

Everything that differs from the official data is coloured:

* **green** - a row you added;
* **orange** - a value you changed. Only the changed cell is coloured, and hovering over it shows
  the "Published value";
* **red, and struck through** - a row you deleted. It is shown where it used to be, so you can see
  what was around it.

**Renaming is still changing the same row.** Some cells say *which* row it is - a component's board
label and region, a schematic's name, an image's board label, region, pin and name, a file's or
link's board label (or category) and name, a credit's category, sub-category and name. Change one of
those and nothing else in the row, and the row is simply changed: that cell turns orange, and
hovering over it shows the old value. Change one of them **and** another cell in the same row - a
credit's name and its contact, say - and CRT can no longer tell it is the same row, so it shows the
row as added (green) and the old one as deleted (red). The same goes for two rows that are identical
apart from those cells: deleting C10 and adding a C51 with exactly the same values reads as C10
renamed to C51.

**CRT also checks your data as you go**, with the same rules the server uses when you submit. A
cell with a problem gets a small triangle in its top-left corner - hover over the cell to see what
is wrong:

* **a red corner is an error** - something the server would refuse, such as a file name that is not
  on your computer (or is spelled with different capital letters), a file type that cannot be
  submitted, two components with the same label in the same region, or a link that is not a web
  address. **Every error has to be fixed before the draft can be submitted**;
* **an amber corner is a warning** - worth a look, but nothing stops you submitting it: a component
  that is not marked on any schematic, an oscilloscope setting CRT does not know (T/DIV, V/DIV or
  T.LVL), a file or link for a component that is not in the Components sheet, or the same part
  number used for two different chips. Two more are about the rows themselves:
  * **two or more rows describe the same thing** - on the Component images sheet, for example, the
    same board label, region, pin and name. **Every one of them** gets the amber corner, so you can
    see which rows collide. All of them are saved, but only the first is compared with the
    published data, so a change in the others is not shown to the maintainer who reviews it. Make
    them differ, or delete the extra row. (Two components with the same label in the same region,
    or two schematics with the same name, are an error instead, as above - and every one of those
    rows gets the red corner, so you can see which rows collide.);
  * **a row is missing a value it needs and will be left out when saving** - an Important signals
    row with a display name but no KiCad net name, say. It is shown like a new row until it is
    complete.

The highlights on the schematics are not in any sheet, so a problem with one - a highlight on a
schematic you renamed in the table, say - is listed in a line above the table instead.

You do not have to open the table to know: **each draft's row on the "Drafts" tab says how many
errors and warnings it has**, in a red and an amber mark beside its name, from the moment the draft
is created. If you edit a draft's Excel file while CRT is open, the numbers are checked again when
you come back to CRT's window or open the "Drafts" tab.

The colour key above the table counts each kind for the whole draft - every sheet together,
whichever one you are looking at - "2 Added", "1 Modified", "1 Errors" and so on. A kind the draft
has some of is filled with its colour; a kind it has none of is only outlined and shown faint (and
cannot be clicked - there would be nothing to see). A count you have clicked gets a firm outline.
**Click a count to see only those rows**, and click it again to see every row. Click several to see the rows of any of
them - Added, Modified and Deleted together show exactly what you changed. The tabs of
sheets with none of those rows are hidden meanwhile, so the tabs left are the places to look - and
clicking "Errors" takes you straight to the first sheet that has one. A new row you have just
inserted always stays visible.

Each row also starts with a small sign that says the same thing without relying on colour: `+`
added, `~` changed, `-` deleted. The number on each sheet's tab is how many rows were added,
changed or deleted on that sheet; errors and warnings are not counted there, because they are not
changes. To find them, click "Errors" or "Warnings" in the colour key, and only the tabs of the
sheets that have them stay. For a board you created yourself nothing is marked as added, changed or
deleted, since all of it is your own - the Added, Modified and Deleted counts are not shown, and a
line above the table says so - but errors and warnings are still shown.

**Find something in the table with the search box** above it, which works like the "Find a
previous repair" box on the Workbooks tab. Type a word and only the rows with it in any cell stay -
on every sheet, with the tabs of sheets that have none hidden - and what was found is marked in
yellow in the cells. Several words must all be in the row, though each in its own cell (`cr pinout`);
put a phrase in quotes to find it as it is (`"pinout (secondary)"`), and a minus in front leaves
out the rows with that word (`pinout -secondary`). Case does not matter, and numbers are found too -
`4164` finds the RAM. When nothing in any sheet matches, only the sheet you are on stays, empty,
with a line above the table saying so. It works together with the colour key: clicking "Modified"
while searching for `pinout` shows only the changed Pinout rows. A row you edit while searching stays on screen even
if it no longer matches, until you change the search. The cross at the end of the box empties it,
and so does closing the table or opening another draft; it stays when you switch sheets or save.

**You can look at a file without leaving the table.** Point at a file name - the *Schematic
image file* column, or *File* on the Component images, Component local files and Board local files
sheets - and a small card opens beside it straight away:

* a **picture** is shown in it. If you changed it, the published picture and yours are shown side by
  side, labelled "Published" and "Your draft", and that includes a picture you replaced under the
  same file name. A line on the card says what happened to the file - "New file", "Removed",
  "Changed to another file", "Replaced - same file name, new content" or "Unchanged" - and
  **Open full size** opens the picture;
* a **PDF** or other document gets a link that opens it in your usual program for that kind of
  file.

Move the pointer down the column and the card follows it from file to file. To click the card's
link, move the pointer across onto the card; move it anywhere else and the card closes at once.

**Editing works much as it does in Excel.** Click a cell to select it - it gets a dashed red
frame - then just start typing, double-click it, or press F2. **Tab** moves to the next cell to the right (and on to the next row
at the end of one), **Shift+Tab** to the left, and **Enter** down. **Insert row** adds an empty
row below the selected one - drag it by its handle to put it anywhere else - and **Delete row**
removes the selected row. Changed your mind? Ctrl+Z, below.

**Seeing all of a long text.** Every text is shown in full. A column is only as wide as its longest
text, up to a limit, and a longer text runs onto more lines, its row growing to fit. Make a column
narrower and its text wraps onto more lines; drag the edge of its heading to make it wider - as wide
as you like. **Double-click the edge of a heading** and the column fits its longest text, as in
Excel - every row counts, not only the ones on screen - but it never grows wider than the table, so a
very long note still wraps.

**Several rows can be deleted at once.** Click a cell in the first row, then hold Shift and click a
cell in the last: every row between is selected, shaded lightly so you can still see its colours.
Ctrl+click adds one more row, or takes a selected one out again. The button then says how many it
will delete - "Delete 8 rows" - and one Ctrl+Z brings them all back. Red rows among them are skipped,
since they are already deleted, and while a count in the colour key is clicked only the rows you can
see are selected. The Delete key on the keyboard never deletes a row.

**Deleting a component deletes everything that belongs to it.** Its rows on the *Component
images*, *Component local files* and *Component links* sheets go with it - you will see them in red
on those sheets - and its highlights on the schematics go when you save. A line above the table
says what else went. If the component has another row with the same board label for a different
region, only the images for the region you deleted go, since the rest still belong to the other
one. One Ctrl+Z brings all of it back.

**Ctrl+Z undoes, Ctrl+Y redoes** (Ctrl+Shift+Z works too, and Cmd on a Mac). That covers every
change in the table - typing into a cell, pasting, and inserting, deleting or moving rows - one
step at a time, and takes you to the sheet and row where the change was. It reaches
back to your last save. While you are still typing inside a cell, Ctrl+Z undoes just that typing,
as in Excel. For something saved earlier, hover over an orange cell to see its published value, and
a red row still shows everything it held - type or copy the values back in.

**Rows can be moved.** Drag a row by the dotted handle at its far left: while you drag, the row
turns into a red dashed empty slot that moves with the mouse, and the other rows make room, so you
can see exactly where it will land - let go of the mouse button there. Dragging above or below the
rows you can see carries it further, one row per move. Alt+Up / Alt+Down move the selected row
one place from the keyboard. A red (deleted) row cannot be moved, and while a count in the colour
key is clicked - showing only some rows - no row can. The order matters: it is the order the
components appear in on the left of the main window. A moved row is not coloured, because its
content has not changed - see the note below about order on its own.

**A new component is put in its place for you.** When you save, a component row you inserted goes
into its category, in label order - C1, C2, C10 - just as components added any other way do (see
below). If you moved it yourself before saving, it stays where you put it. After saving, the cursor
is on the row wherever it went.

**One component, several regions.** A component that differs between regions has one row per region
with the same board label - for example U1 with region PAL and U1 with region NTSC. Each is its own
row: adding the NTSC one shows it as added, and it is saved right beside its PAL twin. Giving an
existing component a region, and changing nothing else in its row, is one changed row - the region
cell turns orange - just like renaming it.

**Copy and paste one cell at a time.** Ctrl+C and Ctrl+V work on a selected cell and inside a cell
you are editing, and cells copied from Excel paste in fine. Pasting a whole block of several cells at
once is not supported yet - copy them one by one.

**Nothing is written until you click "Save changes".** If you close the table, submit, add schematic
images or KiCad data, or quit CRT while the table has edits that are not saved, you are asked
whether to save or discard them first. **"Reload"**, beside "Save changes", reads the draft again
from its file - after asking, if that would throw edits away.

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
> Some changes are saved into your draft but do not count as "rows changed": a new ORDER of rows, a
> KiCad trace calibration, a picture or file replaced under the same file name, and KiCad files. A
> draft whose only changes are of these kinds reads "0 rows changed" and cannot be submitted on its
> own. They are sent along with any other change you submit.

## Sending your work in

When a draft is ready, open the [Drafts tab](Drafts-tab) and click **"Submit"** on that board's row.

**A draft with errors is not sent.** Instead its table opens showing only the rows with errors, and
a line above it says how many there are. Fix them, save the table, and click "Submit" again.
Warnings never stop a submission.

**You do not need an account, and CRT will not ask you to make one.** The "Submit contribution"
dialog asks for two things, and both are required:

* **Your email address** - "You get an email when your contribution has been reviewed, and a
  maintainer can write to you if they need to ask something." It starts out as the address you gave
  on the [Feedback tab](Feedback-tab), if you did.
* **What did you change?** - a short summary, which is what the maintainer reads first. Say which
  exact board revision you have - boards vary, and that detail matters at review time.

**If you are a maintainer signed in on the [Maintainer tab](Maintainer-tab)**, the dialog uses your
account's email address instead and sends the submission with your account, so whoever reviews it
can see the address is checked - also when the Maintainer tab is turned off or hidden. Sign out on
the Maintainer tab first if you want to send with another address.

Before anything is sent, the dialog shows you exactly what is about to go: how many schematics,
components and highlights, how many files are referenced, and how many KiCad files when there are
any. Nothing leaves your machine until you click **Submit**.

**Files you already downloaded are not uploaded again.** CRT asks the server which files it is
missing and sends only those, so correcting a typo on a board with hundreds of images uploads
almost nothing. A large new board with its own images will take longer, and the dialog shows
progress while it works. You can cancel at any point; a cancelled submission leaves nothing behind
on the server and does not touch your draft.

**Your draft stays exactly where it is after you submit.** It is not cleared, emptied or locked,
and you can carry on using the board as normal while the contribution waits to be looked at. That
is deliberate: review takes time, and you should not lose the use of your own work while it
happens.

**What you have just sent cannot be sent again.** As long as the draft holds exactly what you last
sent from it, its "Submit" button is greyed out, and pointing at it says when you sent it. Change
anything the submission carries - a value in the table or in Excel, a component in the Contribute
tab, a highlight, a KiCad calibration, a picture or a file - and "Submit" works again. Opening the
Excel file and saving it without changing a value does not count, and neither does changing a value
and then changing it back. If the earlier send never finished - you cancelled it, or the connection
was lost - the same draft can be sent again straight away.

**The board's row on the Drafts tab shows how your last submission is doing**, in a small badge
beside its name - *Submitted - awaiting feedback from a maintainer* as soon as it is sent, then for example *Published to the BETA source*, or
*Changes requested* in orange when there is something for you to look at. It uses the same words and
colours as "My submissions", and pointing at it tells you when you sent it. It only describes
submissions sent from this draft: if you start a new draft of a board you submitted before, the badge
stays away until you send the new one.

**The board's KiCad data travels with the submission.** If the board has a "KiCad data" folder -
imported through the Drafts tab, or synced with a published board - its KiCad project files
(.kicad_pcb, .kicad_sch, .kicad_pro) are sent along with everything else, and publishing the
contribution publishes them. The maintainer sees them among the board's files, with the new and
changed ones marked. Files the server already has are not uploaded again,
so an untouched KiCad folder costs nothing to send.

**Submitting the same board again replaces what you sent before**, as long as nobody has started
on it yet. Your draft still holds everything, so the newer submission carries all of the earlier
one too. The earlier one leaves the review queue, and "My submissions" shows it as *Replaced by a
newer submission*. If a maintainer has already started on the earlier one - corrected something in
it, or given it a first approval - both are kept and looked at separately. The server knows the two
came from you by the email address you gave (or your account, if you are signed in).

**Once your contribution is published to the stable source, the draft tidies itself away.** Not
before: while it is only on the BETA source a maintainer can still take it back out for another look,
so your draft stays until the second step (see below). When the application has downloaded the
updated data and your draft holds nothing the published board does not - no other rows, no other
KiCad calibration and no other files - the draft is removed on its own, the next time the
application starts or when you close "My submissions", along with any folders it leaves empty in the
Drafts folder. If you kept working in the draft after
submitting, or it is open in the table editor with unsaved edits, it stays, and nothing of yours is
lost. **A whole new system is tidied away the same way**, once CRT lists the published system in
its hardware and board lists. Until then your draft is the only place you can see it, so it stays.
A board you set up the old way, with a `_UserContribution` workbook, is never counted as published
by its own copy in your data folder - only the real published board can tidy its draft away.

If the submission cannot be sent - no internet connection, or a file the data refers to has been
moved or deleted since you added it - you are told which file and what went wrong. Nothing is lost
and your draft is unchanged.

**A few things are refused before anything is uploaded**, each with a message saying which file
and why:

* **Only the kinds of file boards actually use can be sent** - pictures (PNG, JPG, GIF, BMP,
  WebP), PDF documents, plain text and web pages (HTML), plus KiCad project files (.kicad_pcb,
  .kicad_sch, .kicad_pro) inside the board's own "KiCad data" folder. A file of any other kind, or
  a hidden file whose name starts with a dot, is refused.
* **A file must really be what its name says.** A ".png" has to contain a picture and a ".pdf" a
  PDF document; a text file has to be plain text, saved as UTF-8.
* **A board can only change its own files and the shared folders.** You can use a file that belongs
  to another board, exactly as it is, but you cannot change it from here - change it through that
  board's own draft instead.
* **Every file sent must be used by the board.** CRT only ever sends the files your board refers
  to, plus the "KiCad data" folder, so this only matters if a file was added by hand.
* **There is a daily limit per internet connection**, generous enough that it only stops something
  going badly wrong. If you reach it, the message says when you can try again, and your draft is
  kept.

## Checking how a contribution is getting on

The **"My submissions"** button on the Drafts tab lists what you have sent, with the state of each
one and anything the maintainer has said. **"Check for updates"** asks the server for the latest.
Every state and its colour is listed on the Drafts tab page, under
[Submission states](Drafts-tab#submission-states).

You do not have to keep checking. CRT asks the server about anything still waiting when it starts,
and again every minute while its window is open and not minimised. When a maintainer has decided
something or written to you, a red number appears on the **Drafts** tab - and on the "My
submissions" button inside it - counting the contributions with news you have not read yet. Each one
is marked **New** in "My submissions", with a **"Mark as read"** button; the number goes down as you
press it. Opening "My submissions" by itself marks nothing as read.

**When a maintainer asks for changes**, the state reads *Changes requested* and the maintainer's
words are under "Feedback from maintainer". Change your draft - in the Contribute tab, the label
editor or the table - and click **"Submit"** again. The new submission is a new entry in the list.

**An accepted contribution is published in two steps.** First it is published to the **BETA
source**, where a maintainer gives the board a final check; "My submissions" then says *Published to
the BETA source*. After that it is published to the **stable source** that everyone downloads from,
and the row says *Published to the stable source*. You get an email at each step. CRT looks for the
second step at every check for 30 days after the first, and once a week after that; "Check for
updates" asks at any time.

**Checking your work on the BETA source?** Ticking **Download data from the BETA source instead of
the stable source** on the Configuration tab lets you see your contribution as it will look, before
everyone else gets it. CRT tells you when there is something to check: as soon as a contribution of
yours is in the BETA source, a note under its tabs says so and names that box (and, if "Check for new
or updated data at application launch" is off, says to tick that one first), and "My submissions"
says it under the contribution too. The note goes away when you tick the box, or when you close it.
When your contribution then reaches the stable source, CRT shows a note saying so, as a reminder to
untick that box again - the BETA source is for checking, not for everyday use. That note goes away
when you untick the box, or when you close it.

**A brand-new board** is added to CRT's hardware and board lists by the maintainer who accepts it:
they choose the names it is shown under and where in the lists it goes, so it may be listed a little
differently from how you named it.

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

For a brand new hardware and board, the administrator may invite you, by email, to be a
**maintainer** of that system once it is published, so you can look after it from then on -
reviewing what others send in for it, and publishing what is right, in CRT's
[Maintainer tab](Maintainer-tab). Accepting the maintainer role when you created the system does not
do that by itself - the invitation does (see
[Getting a maintainer account](Maintainer-tab#getting-a-maintainer-account)). It is never something
you need before contributing.

## That's it

Thank you - every correction makes the data better for the next person repairing the same board.
