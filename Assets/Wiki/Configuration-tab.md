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
board-data and image sync entirely, and greys out the two settings below, which have nothing to act
on without it.

**Download data from BETA source** — fetches board data from the project's test server instead of
the live one. This is for coordinated testing of data that is not ready yet, and it can leave you
with board data that is incomplete or wrong; only use it in agreement with the developer, or at your
own risk. Ticking it refreshes your data from the BETA source straight away; unticking it takes
effect at the next application launch. Greyed out unless launch-time data checking is on.

**Delete orphan and non-used files** — removes files in your data folder that no board refers to any
more. Housekeeping for a data folder that has been through many updates. Greyed out unless
launch-time data checking is on, since the cleanup runs as part of that check.

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

**Currency used for worklog costs** — pick the country you live in, and its currency is shown
beside every cost in a workbook: on the entry cards, in the summary strip, in the worklog editor,
and in exported PDF and ZIP documents. The field you type a cost into names it too, so
"Cost (DKK)" says what the number will be recorded as.

The list shows the country with its currency code, e.g. `Denmark (DKK)`, and defaults to
`United States (USD)`. Several countries share a currency — the euro countries all store `EUR` —
so reopening this tab may show a different country with the same code. Nothing is converted: the
setting only changes how your own figures are labelled, and changing it relabels costs you have
already recorded rather than recalculating them.

## Your files

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
