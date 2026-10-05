[Wiki Home](Home) · [Workbooks](Workbooks-tab)

Exporting a repair, where the files live, and deleting.

---

## Export

Two buttons in the "Workbooks" tab header, acting on the selected workbook.

| | PDF | ZIP |
| --- | --- | --- |
| The document | Yes | Yes - the same PDF, inside |
| Photos | Embedded, page sized | The original files |
| Attached files | Listed by name only | The original files |
| Use it for | Sharing the write-up | Keeping or sharing everything, originals included |

Both are named like this, and you can rename it in the save dialog:

```
Workbook_<number>_<hardware>_<board>_<date of export>
Workbook_3_Commodore-64_250407-long-board_20260904.pdf
```

The hardware and board are the names in the two drop-downs ("Commodore 64" and "250407 (long board)"), with spaces, brackets and other punctuation turned into hyphens. The date is the day you exported, as year, month, day.

The workbook description is deliberately not in the file name - it often holds personal details, on a file you may be about to send to someone.

The file is not opened after export - you get it where you saved it. If an export fails, no file appears and no message is shown - the reason is written to the log ("Configuration" tab -> **Open logs and settings folder**).

### What the PDF contains

* A header: the workbook description and status, its number, board, when it started (or ended), and the date it was exported.
* The workbook note, then the totals.
* One section per schematic, in alphabetical order - the first follows the totals, every later one starts on a new page. Each shows the schematic at full page width with the marked areas drawn on it in their category colours.
* Then each worklog on that schematic: its title, category and state, description, components in scope (finished ones marked "(done)"), work done, comments, links, the names of its attached files, and its photos.

Each photo is shown with its file name and comment, so a recipient can find that exact file in the ZIP. Web links in the text, and the worklog's links, are clickable in the PDF.

### What the ZIP contains

```
Workbook_3_Commodore-64_250407-long-board_20260904.zip
├── Workbook_3_Commodore-64_250407-long-board_20260904.pdf
├── worklog_1/
│   ├── 5v-rail-ripple.png
│   └── 7805-datasheet.pdf
└── worklog_2/
    └── vic-socket.jpg
```

One folder per worklog, named the same as on your own disk.

## Where your files are

Everything is on your own machine, in a `Workbooks` folder next to your settings and log:

* Windows: `%LocalAppData%\Classic-Repair-Toolbox\Workbooks`
* Linux: `~/.local/share/Classic-Repair-Toolbox/Workbooks`
* macOS: `~/Library/Application Support/Classic-Repair-Toolbox/Workbooks`

The "Configuration" tab has a button `Open workbooks folder` that takes you straight there.

One folder per workbook, holding everything belonging to it:

```
Workbooks/
├── index.json                  <- bookkeeping (the next workbook number)
├── workbook_1/
│   ├── index.json              <- the workbook
│   ├── worklog_1/
│   │   ├── index.json          <- worklog #1, all of it
│   │   └── 5v-rail.png         <- and its photos and files
│   └── worklog_2/
│       └── index.json
└── workbook_2/
```

**Every folder holds one `index.json`, and that is the whole record.** A workbook is a folder; a worklog is a folder inside it, holding its own details together with its photos and files.

That means **you can delete a worklog, or a whole workbook, by deleting its folder** - the application simply stops showing it, and there is nothing left over to tidy up.

The `index.json` at the top of `Workbooks/` is not a workbook. It records which numbers have been handed out, so a deleted workbook's number is never given to a new one.

**To back up: copy the `Workbooks` folder.** That is all of it. To move to another machine, copy it across - nothing else needs doing.

Close the application first, so you do not catch a file mid-write.

You can put the folder somewhere else with `--workbooks-root=`, see [Commandline parameters](Commandline-parameters).

## Deleting

Both are permanent - there is no undo. Each asks you to confirm first.

**Delete a worklog** - the "Delete worklog" button on its card. Removes the worklog and its photos and files.

**Delete a workbook** - the "Delete workbook" button in the header. Removes the whole repair: every worklog, photo and file in it. If it was the workbook you were working in, the newest remaining workbook on that board takes its place.

Export to ZIP first if there is any chance you want it back.

### Numbers are not reused

Delete workbook #2 of two, and the next one you create is **#3**. The gap stays.

That is deliberate: #2 may name a PDF you have already exported and sent on, and reusing the number would make that document describe a different repair.

Same for worklog numbers, and deleting one does not renumber the rest.
