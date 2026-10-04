[Wiki Home](Home)

Every setting in CRT, and the button that opens your data folder.

---

## Appearance

**Theme** — `Light`, `Dark` or `User preference`.

`User preference` uses your own colours instead of the built-in ones. They live in
`Classic-Repair-Toolbox.settings.json` under `userThemeColors`, as name/colour pairs — the file is
created with the defaults filled in the first time you use the option, so you have a full list to
edit rather than a blank page.

Edit the file, then click **Reload user preference colors** to see the result without restarting.

## Component popup

**Open multiple component info windows** — off, each component you click reuses the same popup
window. On, every component opens its own, so you can put two chips side by side and compare.

**Detach thumbnails into their own window** — moves the [Schematics](Schematics-tab) tab's thumbnail
strip out of the tab and into its own resizable window, freeing up the space it used for the main
image. The window fits every thumbnail to whatever size you give it, so a small window shows them
smaller and a maximized one shows them bigger - never scrolled. Selecting or dragging a thumbnail
to reorder it there works the same as in the tab, and you can now also drag sideways. Turning this
off puts the thumbnail strip back in the tab exactly as before; closing the window directly does
the same. You can also turn it on with a right-click directly on the thumbnail strip, without
coming here.

**Remember thumbnail window settings per board** — only available while the above is ticked. Off
(the default), detached/embedded and the window's size and position are one setting for every
board. On, each board remembers its own: one board's thumbnail window can be maximized, another's a
small window in a particular spot, and another can simply have thumbnails embedded in the tab
instead of shown in a window at all. A board you have not touched yet starts embedded.

Switching this on or off takes effect straight away for the board you are on, so the thumbnails may
move into a window or back into the tab as you tick it - that is the setting for that particular
board taking over from the shared one, or handing back to it.

## Data

**Check for new or updated data at application launch** — CRT downloads new and corrected board data
from the project's server. Leave it on; there are frequent updates. Turning it off also skips the
board-data and image sync entirely, and greys out the two settings below, which have nothing to act
on without it.

**Download data from the BETA source instead of the stable source** — CRT has two sources of board
data. The **stable source** is the tested data everybody gets, and is what CRT uses unless you say
otherwise. The **BETA source** holds data a maintainer has just approved, before it has had its
final check and gone out to everybody; it can be incomplete or wrong. Tick this to check a
contribution of your own once it has been approved, then untick it again. Ticking it refreshes your
data from the BETA source straight away; unticking it takes effect at the next application launch.
Greyed out unless launch-time data checking is on. If you ticked it to check a contribution of your
own, CRT tells you when that contribution reaches the stable source, so you remember to untick it.

**Delete orphan and non-used files** — removes files in your data folder that no board refers to any
more. Housekeeping for a data folder that has been through many updates. Greyed out unless
launch-time data checking is on, since the cleanup runs as part of that check.

## Updates

**Check for new version at application launch** — tells you when a new CRT release is out. You can
then update from inside the application.

**Allow notification for BETA versions** — also offers BETA releases. A BETA is a version where
development has reached a mature level and everything should work as intended, but it is still for
testing only. Leave it off unless you want to help find problems before a release.

**Allow notification for ALPHA versions** — also offers ALPHA releases. An ALPHA is a development
version where almost anything could be broken. Leave it off unless you have agreed with the
developer to try one.

The two boxes are independent, and each one only adds its own kind of release. Ticking BETA alone
never offers you an ALPHA, even when a newer ALPHA exists. With both unticked you are offered
ordinary releases only — which you always are, whatever these are set to.

## Oscilloscope

**Enable network connected oscilloscope tab** — shows or hides the "Oscilloscope" tab. See
[Synchronize oscilloscope](Synchronize-oscilloscope).

## MiniPro

**Enable MiniPro programmer functionality** — shows the IC-testing features. See
[MiniPro programmer](MiniPro-programmer).

**Enable MiniPro programmer simulated demo mode** — pretends a programmer is attached, for
development. You do not need this.

## Workbooks

**Enable Workbooks tab** — shows or hides the [Workbooks](Workbooks-tab) feature and its bar above the
tabs.

**Scope in Workbooks tab** — whether the tab lists workbooks for the selected board only (the
default), or every workbook on every board.

**Currency used for worklog costs** — pick the country you live in, and its currency is shown
beside every cost in a workbook: on the entry cards, in the summary strip, in the worklog editor,
and in exported PDF and ZIP documents. The field you type a cost into names it too, so
"Cost (DKK)" says what the number will be recorded as.

The list shows the country with its currency code, e.g. `Denmark (DKK)`, and defaults to
`United States (USD)`. Several countries share a currency — the euro countries all store `EUR` —
so reopening this tab may show a different country with the same code. Nothing is converted: the
setting only changes how your own figures are labelled, and changing it relabels costs you have
already recorded rather than recalculating them.

## Contributing and maintaining data

**View boards as officially coming from online source (hide my local draft changes)** — hides your
own local, unpublished draft edits, so a board shows exactly what CRT downloaded and everyone else
sees. That covers everything in a draft: its rows, and also any schematic image or file you
replaced, KiCad data you imported and KiCad calibration you changed. Untick it again to see your
edits marked and applied as usual. The "?" beside it opens
[View boards from online source](View-boards-from-online-source), and
[Contribute data via CRT](Contribute-data-via-CRT) says what a draft is.

**Enable Maintainer tab** - shows or hides the [Maintainer tab](Maintainer-tab), where maintainers
of the hardware data review and publish what other people send in. Off by default: it needs a
maintainer account, which the administrator gives by invitation, and it does nothing without one.
Turning it off again does not sign you out: as long as you are signed in, the Feedback tab and the
Submit dialog use your maintainer account's email address. To sign out, turn the tab on and sign out
there.

**Hide the Maintainer tab while no work is waiting for me** - with this on, the Maintainer tab only
appears while its badge shows a number, which is exactly when there is something for you to do: a
contribution to review, or a system waiting to go from BETA to stable. The tab you are already on
is never taken away under you - it goes once you move to another tab - and neither is a tab
holding table changes you have not saved. Greyed out unless the Maintainer tab is enabled, which it
narrows rather than replaces. A tab that is not signed in stays visible, so you can always get back
to the sign-in screen. Hiding the tab does not sign you out either.

## Visible hardware, boards and schematics

The right-hand panel lists every hardware, board and schematic image CRT knows about, each with
its own checkbox, fully expanded. Untick anything you do not want cluttering the hardware and board
drop-downs or the schematic thumbnails - useful once the list of supported hardware grows past what
you personally work on.

Unticking an item disables its children (they grey out) rather than clearing their own ticks, so
nothing is lost: retick the parent and everything underneath reappears exactly as you left it. A
hardware with no ticked boards left, or a board with no ticked schematics left, drops out of its
drop-down entirely.

An unticked item is shown in red, but only while it is still active - once its parent is also
unticked (and greyed out), its own red is hidden too, so unticking a whole hardware does not turn
its entire board and schematic list red as well.

Every hardware and board also has a small arrow to collapse and expand it - handy for hiding a
hardware or board you have unticked a lot of, so it stops taking up space in the list. Each one
remembers whether you left it collapsed or expanded.

Drag the splitter between the two panels to resize; the width is remembered.

## Application related files

Three buttons, one per folder:

**Open data folder** - the hardware reference data CRT downloads: schematics, component images and
the board files behind them.

**Open workbooks folder** - your own workbooks, with the worklogs, photos and attached files in
them. One folder per workbook.

**Open logs and settings folder** - the log file, the crash log and your settings file. These are
the files to attach when reporting a problem.

Each button opens the folder CRT is **actually** using, so if you have moved the data or workbooks
folder with a command-line parameter, the button follows it there.

The log is the first place to look when something did not work — it names every problem CRT found
in the data at startup.

To put the downloaded data or your workbooks somewhere else, see
[Command-line parameters](Commandline-parameters).
