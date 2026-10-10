[Wiki Home](Home)

Your own local changes to board data - edit them, check them, and send them in for review.

---

The **Drafts** tab lists every board you have changed locally, and is where you send those changes in.
Nothing on it is seen by anyone else until you submit it, and a draft you never submit keeps working
on your own computer for as long as you want it.

The step-by-step walkthrough, from the first edit to a published contribution, is
[Contribute data via CRT](Contribute-data-via-CRT). This page says what each part of the tab does.

## When the tab is there

The tab sits right after **Contribute**. It is only shown while you have at least one draft - or
while a maintainer has replied to something you sent and you have not read it yet, so the reply is
never out of reach. When only such a reply keeps it open, it says "A maintainer has replied to
something you sent. Open "My submissions" above to read it."

A red number on the tab counts the contributions with news you have not read yet: a comment from a
maintainer, or a new state. It is the same number as on the **My submissions** button inside the tab.

## What a draft is

A draft is your own copy of one board's folder: its Excel file, its `.json` file (the component
highlights and the KiCad trace calibration), the board's files and its `KiCad data` folder. While a
board has a draft, CRT shows the board from the draft - with your changes - instead of the downloaded
one. To see a board as everybody else does, see
[View boards from online source](View-boards-from-online-source).

CRT works out what you changed by comparing the draft with the published board, so it does not
matter where you make a change: in CRT, or in the draft's Excel file.

Drafts are kept in:

* Windows: `%LocalAppData%\Classic-Repair-Toolbox\Drafts`
* Linux: `~/.local/share/Classic-Repair-Toolbox/Drafts`
* macOS: `~/Library/Application Support/Classic-Repair-Toolbox/Drafts`

with one folder per board below it - manufacturer, hardware, board - for example
`Drafts\Commodore\C64\250407\`. There is no button in CRT that opens this folder. The
[command-line parameters](Commandline-parameters) can put it somewhere else.

## How a draft is made

A draft is made the first time you save a change to a board, as a copy of the board as CRT has it
downloaded - or straight away, when you ask for one:

* **Edit board as draft** on the [Contribute tab](Contribute-tab) - CRT makes the draft of the board you
  have selected and brings you here with its table open. A board has only one draft: if it has one
  already, CRT says so and offers to open it
* **Save to draft** in the component editor on the [Contribute tab](Contribute-tab) - CRT then brings
  you here
* saving new or changed component labels in the label editor on the [Schematics tab](Schematics-tab)
* **Apply KiCad calibration** on the Schematics tab
* **Add a new board** - on the Contribute tab, or at the top of this tab - for a hardware and board
  that is not in the lists at all. Fill in **Manufacturer**, **Hardware**, **Board** and, if you like,
  **Notes (optional)**, choose **Create board**, then **Accept and create** in "Becoming the board's
  maintainer". CRT selects the new board and brings you here. The full walkthrough is
  [Add new board with KiCad data](Add-new-board-with-KiCad-data).
* copying a board folder you already have into the drafts folder yourself - see
  [Bringing in a board you already have work on](Contribute-data-via-CRT#bringing-in-a-board-you-already-have-work-on)

Accepting the maintainer role when you create a board sends nothing by itself. A board's
maintainers are invited by the administrator, by email - see [Maintainer tab](Maintainer-tab).

## A draft's row

Each board with a draft has one row, in alphabetical order:

| Part of the row | What it says |
| --- | --- |
| The name | The hardware and board |
| A coloured badge | How your last submission from this draft is getting on, in the same words and colour as **My submissions** (see [Submission states](#submission-states)). Point at it to see when you sent it |
| **N errors** (red) | Problems the server would refuse. Every one has to be fixed before the draft can be submitted |
| **N warnings** (amber) | Worth a look, but they never stop a submission |
| **N rows changed** | How many rows differ from the published board. A board you created yourself says "New board, nothing added yet" or "New board, N rows so far" |
| A line about the official data | Only when the published board has changed since you started: "The official data for this board has been updated since you started. Your edits are still applied." |

The numbers are worked out again whenever you open the tab or come back to CRT's window, so a change
you saved in Excel shows up.

The buttons on the row:

| Button | What it does |
| --- | --- |
| **What changed** | Only when the published board has changed since you started - see [What changed](#what-changed) |
| **Edit in table format** | Opens the draft as a table right below the row - see [The table](#the-table). Reads **Close table** while it is open |
| **Schematic images** | Add, remove and order the board's schematic images |
| **KiCad data** | Import or remove the board's KiCad project |
| **Submit** | Send the draft in for review - see [Submit](#submit) |
| **Discard** | Throw the draft away - see [Discard](#discard) |

At the top of the tab are **My submissions** (with the red number when there is news) and
**Add a new board**.

## What changed

When the published board has been updated since you started your draft, its row says so and offers
**What changed**. CRT also shows a one-line notice while that board is on screen, with its own
**What changed** button.

The window, "What changed officially", shows the revision you started from ("You started from:") and
the one published now ("Officially now:"), then every change of yours, one per line:

* "You added this. It is not in the published data."
* "You removed this. It is still in the published data, and your submission would remove it."
* "You changed:" followed by the columns you changed

Your edits are still applied on top of the new official data - nothing of yours is overwritten.
**I have looked at this** stops the notice until the official data changes again. It changes none of
your rows.

## Schematic images and KiCad data

Both buttons open the same window, each on its own part. Files are always copied into the draft,
never moved. If the table is open with unsaved changes, you are asked about them first.

**Schematic images** - drag image files onto the window, or use **Choose image files**. Each image is
one schematic of the board. Drag the images up or down to change the order they are listed in, and
**Remove** takes one out.

**KiCad data** - drag the KiCad project folder onto the window, or use **Choose KiCad folder**.
Sub-folders are included, so a project with several sheets keeps them all; footprint libraries, 3D
models and gerbers are left out. Under "What the KiCad data lines up with", the window lists the
component labels with no match in the KiCad data, the KiCad components that have no label yet, and
the CAD names you can use for a schematic image.

The full walkthrough is [Add new board with KiCad data](Add-new-board-with-KiCad-data).

## The table

**Edit in table format** opens the draft's data the way its Excel file holds it, one tab per sheet:
Board schematics, Components, Component images, Component local files, Component links, Board local
files, Board links, Important signals and Credits. The other drafts are hidden while it is open. A
sheet's tab says how many of its rows changed, for example "Components (3)".

The long-form explanation is
[Editing a draft as a table](Contribute-data-via-CRT#editing-a-draft-as-a-table). In short:

### Colours

* **green** - a row you added, marked `+`
* **orange** - a cell you changed, in a row marked `~`. Point at it to see the value it had in the
  data you downloaded - "BETA source value" or "Stable source value", after the source you picked
  in the [Configuration tab](Configuration-tab)
* **red, struck through** - a row you deleted, marked `-`, shown where it used to be

A row where only the cells saying which row it is changed - a component's board label or region, a
credit's name - is one changed row, not a deleted row and an added one.

A board you created yourself has nothing published to compare with, so nothing is coloured and only
the Errors and Warnings counts are shown.

### The colour key is the filter

Above the table are **Added**, **Modified**, **Deleted**, **Errors** and **Warnings**, each with its
count for the whole draft, every sheet together. Click one to see only those rows, and click it again
to see every row. Click several to see the rows of any of them. Meanwhile the tabs of sheets with none
of those rows are hidden. A count of 0 is faint and cannot be clicked. No row can be moved while a
count is clicked, and a row you have just inserted always stays visible.

### Search

The box "Find in the table" keeps only the rows with your text in any cell - numbers too - on every
sheet, and marks what it found in yellow. Several words must all be in the row, each in its own cell;
`"quotes"` find a phrase as it is, and `-word` leaves out the rows with that word. Case does not
matter. It works together with the colour key. When nothing on any sheet matches, only the sheet you
are on stays, empty, with "Nothing in any sheet matches ..." above it. The cross at the end of the box
empties it.

### Editing

| Key or action | What it does |
| --- | --- |
| Click | Selects a cell - it gets a dashed red frame |
| Start typing, double-click or F2 | Edits the cell |
| Tab / Shift+Tab | The next cell to the right / left, on to the next row at the end of one |
| Enter | The cell below |
| Ctrl+C / Ctrl+V | Copy and paste one cell. A block of several cells is not pasted |
| Ctrl+Z | Undo, one step at a time, back to your last save. While you are typing in a cell, it undoes only that typing |
| Ctrl+Y or Ctrl+Shift+Z | Redo |
| Shift+click / Ctrl+click | Select every row between / one more row (or take one out again) |
| Drag the grip at a row's far left | Move the row. It shows as a red dashed slot while you drag |
| Alt+Up / Alt+Down | Move the selected row one place |
| Drag the edge of a column heading | Make the column wider or narrower |
| Double-click the edge of a column heading | Fit the column to its longest text, never wider than the table |

On a Mac, use Cmd instead of Ctrl. Every text is shown in full, over several lines when it is long.

### Rows

* **Insert row** adds an empty row below the selected one. When you save, a new component is put into
  its category, in label order (C1, C2, C10), unless you moved it yourself.
* **Delete row** deletes the selected rows - with several selected it reads, for example,
  **Delete 8 rows**. A published row stays visible in red; a row you added goes. Deleting a component also
  deletes its rows on the Component images, Component local files and Component links sheets, and its
  highlights when you save - the line above the table says what went, and one Ctrl+Z brings it all
  back. The Delete key on the keyboard never deletes a row.

### Files

Point at a file name - "Schematic image file" on Board schematics, or "File" on Component images,
Component local files and Board local files - and a card opens beside it straight away:

* a picture is shown. If you changed it, the published one and yours are shown side by side as
  "Published" and "Your draft" - also when you replaced a picture under the same file name. The card
  says "New file", "Removed", "Changed to another file", "Replaced - same file name, new content" or
  "Unchanged", and **Open full size** opens the picture;
* any other file gets a link that opens it in your usual program for that kind of file.

The card follows the pointer from file to file, and closes when the pointer leaves the cell and the
card.

### Checks

CRT checks the draft with the same rules the server uses when you submit. A cell with a problem has a
small triangle in its top-left corner - **red for an error**, which has to be fixed before the draft
can be submitted, and **amber for a warning**, which never stops anything. Point at the cell to read
what is wrong. A problem with a highlight on a schematic, which is in no sheet, is listed above the
table. What the checks look for:
[Editing a draft as a table](Contribute-data-via-CRT#editing-a-draft-as-a-table).

### When the Excel file changes too

* **"This draft appears to be open in Excel (or another spreadsheet program)..."** - a yellow bar for
  as long as the draft's Excel file is open there. Edit in one place at a time: the table follows what
  you save in Excel, but **Save changes** is greyed out until the file is closed there.
* **"The draft was changed outside this table..."** - the file was saved somewhere else while the
  table had unsaved changes. Those changes can no longer be saved into it: copy what you need, then
  press **Reload** on the bar. With nothing unsaved, the table simply reads the file again and says
  "Updated - the draft was changed outside this table."

### Saving

Nothing is written until you press **Save changes** (greyed out while there is nothing to save). It
says "Saved." when done, and the draft's row and the board on the other tabs follow. **Reload** beside
it reads the draft again from its file.

Closing the table, submitting, **Schematic images**, **KiCad data** or quitting CRT while the table has
unsaved changes asks first, in an "Unsaved edits" window: **Save changes**, **Discard edits** or
**Cancel**. When the changes cannot be saved - the file is open in Excel, or changed since - and when
you press **Reload**, only **Discard edits** and **Cancel** are offered. Enter and Escape always
cancel.

While the table has unsaved changes, **Save to draft** in the Contribute tab and the label editor's
save will not save that board - they ask you to save or discard the table's changes here first.

## Submit

**Submit** sends the draft in for review. It is greyed out - point at it to see why - when:

* there is nothing to send yet: "There is nothing to send yet. Add some data to this board first."
* the draft holds exactly what you last sent from it: "You have already sent this draft as it is now,
  on [date]. Change something to send it again." Saving the Excel file without changing a value, or
  changing one and changing it back, does not count. If the earlier send never finished, the same
  draft can be sent again straight away.

**A draft with errors is not sent.** Its table opens showing only the rows with errors, with a line
above it saying how many there are. Fix them, save the table, and press **Submit** again.

The "Submit contribution" window first shows what is about to go - how many schematics, components,
highlights and referenced files, and KiCad files when there are any - and asks for two things, both
required:

* **Your email address** - "You get an email when your contribution has been reviewed, and a
  maintainer can write to you if they need to ask something." It is the same address the
  [Feedback tab](Feedback-tab) uses, and there is no account to create. If you are signed in on the
  [Maintainer tab](Maintainer-tab), your account's address is used and cannot be changed here.
* **What did you change?** - a sentence is enough; it is the first thing the maintainer reads. This is
  the place to explain your change, and to say which exact board revision you have.

**Submit** then sends it, uploading only the files the server does not have yet, with a progress bar.
**Cancel** stops it, leaving nothing behind. The window ends with "Contribution sent", "Contribution
not accepted" (with the reasons) or "Contribution not sent" (for example with no internet connection).
Your draft is not changed either way, and you can keep using it while the contribution is reviewed.

Sending the same board again replaces your earlier submission, as long as no maintainer has started
on it yet. More: [Sending your work in](Contribute-data-via-CRT#sending-your-work-in).

## My submissions

**My submissions** lists what you have sent from this computer - the list is kept on this computer
only. Each entry shows the board, when it was sent, what you wrote, and its state in colour (below).
What a maintainer wrote to you is shown under "Feedback from maintainer".

News you have not read - a comment, or a new state - is marked **New**, with a **Mark as read** button
beside it. The red number on the Drafts tab and on **My submissions** counts the entries with unread
news, and goes down as you press **Mark as read**. Opening the window by itself marks nothing as read.

CRT asks the server about your open submissions when it starts, and every minute while its window is
open and not minimised. A submission in the BETA source is asked about like that for 30 days after it
got there, then once a week. **Check for updates** asks straight away. **Remove** takes an entry off
the list on this computer only: the contribution is not withdrawn, but CRT can no longer show you how
it went.

### Submission states

| State | Colour | Means |
| --- | --- | --- |
| Submitted - awaiting feedback from a maintainer | amber | Waiting for a maintainer |
| Approved, waiting to be published | green | One of two approvals given. A change that replaces a shared file other boards use needs a maintainer of the board AND the administrator |
| Published to the BETA source | green | Accepted, and in the BETA data for a final check |
| Published to the stable source | green | Out to everybody |
| Changes requested | orange | A maintainer asks you to change something - see below |
| Taken back out of BETA - waiting for review again | orange | A maintainer took it back out of the BETA data for another look |
| Not accepted | red | Rejected - the feedback says why |
| Replaced by a newer submission | amber | You sent the same board again before this one was reviewed |
| No longer on the server | amber | It was deleted on the server, for example together with its board. You can send it again |
| Never finished sending | red | The upload did not complete - send it again |
| Expired before it was finished | red | The upload was left unfinished for too long - send it again |
| Not checked yet | amber | CRT has not heard from the server about it yet |

**After "Changes requested":** read the feedback in **My submissions**, change your draft - in CRT or
in the table - and press **Submit** again. The new submission is listed as a new entry.

## Discard

**Discard** throws the draft away, after asking in "Discard draft": everything you changed locally for
that board is lost, and the published data is not affected. Enter and Escape both cancel.

* If something you sent from the draft is still being reviewed, the window says so. Your submission is
  not withdrawn, but the board's maintainers are told that you discarded the draft.
* If a file of the draft is open in another program - its workbook in Excel, say - part of it stays,
  and a line at the top of the tab says so. Close the file there and press **Discard** again.
* A board you created yourself disappears from the hardware and board lists together with its draft.

## When your work is published

Once a contribution reaches the stable source, CRT removes its draft by itself - as soon as CRT has
downloaded the published board and the draft holds nothing the published board does not. Not before:
while it is only in the BETA source, a maintainer can still take it back out. A draft you kept working
on after submitting stays, and so does one whose table has unsaved changes. A whole new board's draft
goes once CRT's hardware and board lists include the published board.

## BETA and stable notices

When a contribution of yours is published to the BETA source while you download from the stable
source, a notice under the tabs says so and names the box to tick on the
[Configuration tab](Configuration-tab) to try it before everybody else: "Download data from the BETA
source instead of the stable source". When it later reaches the stable source while you download from
BETA, a notice reminds you that you can switch back. **Open Configuration** takes you to the setting,
and the cross closes the notice.

## When CRT has to be updated

Now and then the server that drafts are sent to changes in a way an older CRT cannot follow. While
you have drafts (or the [Maintainer tab](Maintainer-tab) is turned on), CRT asks the server about this
when it starts - or, with no drafts yet, as soon as **Edit board as draft** or **Add a new board** is
pressed, before your first draft is made - and it notices it whenever the server answers that way. The whole Drafts tab is then covered by **CRT has to be updated**, which cannot be closed:
until CRT is updated, drafts cannot be submitted and the server cannot be asked how your submissions
are getting on. Your drafts stay on this computer, untouched, and are all there again after the
update. Everything else in CRT works as before.

The button on it is the way out. If CRT has already found a newer version - with **Check for new
version at application launch** ticked on the [Configuration tab](Configuration-tab) - **Install
update** downloads and installs it, as the update notice at the top of the window does. Otherwise
**Open download page** opens the page listing every CRT release in your web browser.

If a draft's table holds edits you have not saved, installing the update asks first, as quitting
CRT does: save them into the draft, discard them, or cancel and install nothing. While the tab is
covered, **Edit board as draft** and **Add a new board** on the [Contribute tab](Contribute-tab) make
nothing and say why, since their next step is on this tab.

## Limits

* Changes that are not rows read "0 rows changed", and a draft with only such changes cannot be
  submitted on its own: a new row order, a KiCad trace calibration, a picture or file replaced under
  the same file name, or KiCad files only. They are sent along with any other change.
* Copy and paste work one cell at a time.
* The counts you clicked in the colour key are forgotten when CRT closes.
* There is no button that opens the drafts folder - see [What a draft is](#what-a-draft-is) for where
  it is.
