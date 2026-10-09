[Wiki Home](Home)

Fix or add board data from inside the application, and send it in for review.

---

All the hardware knowledge in CRT is data that people contributed. When you spot something wrong or
missing for the board in front of you, this tab is how you fix it.

You can:

* correct a component's **friendly name, technical name, part number or description**
* **add a new component**, or delete one that is not on the board
* add **component images**, **files** and **links** - a datasheet, a photo, a reference page
* add **board-level files and links**
* **edit the whole board in table form**, the way you would in Excel
* **add a whole new board** - a hardware and board that is not in the lists at all

The tab lists the components of the board that is selected, in one column per category. Point at a
board label to see its full name, and click it to open the component editor. **Add new component**
opens the same editor on a blank component, for one the board data does not have yet - it is greyed
out until a board is loaded. The line **"Board Excel data last revisioned:"** says how fresh the
board's data is.

Every edit is saved into your own local **draft** for this board straight away - you see it on the
board immediately, and the [Drafts tab](Drafts-tab) lists every board you have local changes on.
Nothing you edit here changes what anyone else sees until you submit it for review from the Drafts
tab.

A small amber **Draft** badge marks where your drafts are: on each board with a draft in the
**Board** drop-down, on its hardware in the **Hardware** drop-down, and on each changed component in
the component list.

If the official data for a board is updated while you have a draft on it, the Drafts tab says so
under that board and offers a **"What changed"** view. Your edits are still applied - see
[Drafts tab](Drafts-tab#what-changed) for what that view shows.

**The full walkthrough, step by step:
[Contribute data via CRT](Contribute-data-via-CRT).**

## Editing the whole board as a draft

**Edit board as draft** makes your own copy of the board in front of you - a **draft** - and opens it
in the table on the [Drafts tab](Drafts-tab), where you can change many values at once, the way you
would in Excel. Nothing anybody else sees changes until you submit it from there. It is greyed out
until a board is loaded.

A board can have only **one** draft. If you already have one for this board - from this button, or
from saving a component or labels - nothing new is made: a window says so and offers **Open the
draft**, which opens the draft you have in the same table. Every change you make to a board goes into
that one draft.

While the Drafts tab says **CRT has to be updated**, this button and **Add a new board** below make
nothing and say why: their next step is on the Drafts tab, which cannot be used until CRT is updated
(see [Drafts tab](Drafts-tab)).

## Adding a whole new board

The **Add a new board** button creates a brand-new hardware and board of your own. You give it a
manufacturer, hardware and board name - and, if you like, notes about the hardware for the maintainer
who adds it to CRT's lists - and choose **Create board**. It appears in the lists straight away as a local draft, empty and ready
for you to add board images, label components and attach files exactly as you would edit any other
board, and CRT takes you to the [Drafts tab](Drafts-tab), where the next steps are.

Creating one asks you to accept the role of **maintainer** for it, in "Becoming the board's
maintainer": if you submit the board for the community to use, you agree to review the changes others
submit for it and publish the ones that are right. The board is only created if you choose
**Accept and create**. Accepting sends nothing by itself - the administrator invites a board's
maintainers by email, see [Maintainer tab](Maintainer-tab).

Unlike **Add new component**, this button does not need a board loaded first. That is the point: it
is for the board that is not there.

A board you never submit is not a half-finished thing sitting in a queue - it keeps working locally
for as long as you want it. The full walkthrough, including importing KiCad data:
[Add new board with KiCad data](Add-new-board-with-KiCad-data).

## What goes elsewhere

| | |
| --- | --- |
| A component highlighted in the wrong place | The label editor - see [The Schematics tab](Schematics-tab) |
| Schematic images, or a board's KiCad data | The **Schematic images** and **KiCad data** buttons on the [Drafts tab](Drafts-tab) |
| Many values at once, the way you would in Excel | **Edit board as draft** above, or **Edit in table format** on the [Drafts tab](Drafts-tab) |
| Something wrong with the application itself | [The Feedback tab](Feedback-tab) |
