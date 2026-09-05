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

## Data

**Check for new or updated data at application launch** — CRT downloads new and corrected board data
from the project's server. Leave it on; there are frequent updates. Turning it off also skips the
board-data and image sync entirely.

**Delete orphan and non-used files** — removes files in your data folder that no board refers to any
more. Housekeeping for a data folder that has been through many updates.

## Updates

**Check for new version at application launch** — tells you when a new CRT release is out. You can
then update from inside the application.

**Allow notification for BETA versions** — also offers pre-release versions. These are for testing;
leave it off unless you want to help find problems before a release.

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

## Your files

**Open data/workbooks/log/settings folder** opens the folder holding your settings file, the log and
your workbooks.

The log is the first place to look when something did not work — it names every problem CRT found
in the data at startup.

To put the downloaded data or your workbooks somewhere else, see
[Command-line parameters](Commandline-parameters).
