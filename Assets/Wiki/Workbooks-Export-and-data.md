[Wiki Home](Home) · [Workbooks](Workbooks)

Handing the job over, where the files live, and deleting.

---

## Export

Two buttons in the "Workbooks" tab header, acting on the selected workbook.

| | PDF | ZIP |
| --- | --- | --- |
| The document | Yes | Yes - the same PDF, inside |
| Photos | Embedded, page sized | The original files |
| Attached files | No | Yes |
| Use it for | Emailing a customer | Handing over everything |

Both are named like this, and you can rename it in the save dialog:

```
Workbook_3_Commodore_C64_20260904.pdf
```

The workbook description is deliberately not in the file name - it often holds a customer's own details, on a file about to be emailed.

### What the PDF contains

The totals first, then one section per schematic, each starting on a new page: the schematic at full page width with the marked areas drawn on it in their category colours, then each worklog with its description, category, state, work done, comments, links and photos.

Each photo is shown with its file name and comment, so a recipient can find that exact file in the ZIP.

### What the ZIP contains

```
Workbook_3_Commodore_C64_20260904.zip
├── Workbook_3_Commodore_C64_20260904.pdf
├── worklog_1/
│   ├── 5v-rail-ripple.png
│   └── 7805-datasheet.pdf
└── worklog_2/
    └── vic-socket.jpg
```

One folder per worklog, named the same as on your own disk.

The file is not opened after export - you get it where you saved it.

## Where your files are

Everything is on your own machine, in a `Workbooks` folder next to your settings and log:

* Windows: `%LocalAppData%\Classic-Repair-Toolbox\Workbooks`
* Linux and macOS: `~/.local/share/Classic-Repair-Toolbox/Workbooks`

The "Configuration" tab has a button `Open data/workbooks/log/settings folder` that takes you there.

One folder per workbook, holding everything belonging to it:

```
Workbooks/
├── 1/
│   ├── index.json          <- the workbook
│   ├── entries.json        <- all its worklogs
│   └── worklog_1/          <- worklog #1's photos and files
│       └── 5v-rail.png
└── 2/
```

**To back up: copy the `Workbooks` folder.** That is all of it. To move to another machine, copy it across - nothing else needs doing.

Close the application first, so you do not catch a file mid-write.

You can put the folder somewhere else with `--workbooks-root=`, see [Commandline parameters](Commandline-parameters).

## Deleting

Both are permanent - there is no undo.

**Delete a worklog** - the ✕ on its card. Removes the worklog and its photos and files.

**Delete a workbook** - the button in the header. Removes the whole job: every worklog, photo and file in it.

Export to ZIP first if there is any chance you want it back.

### Numbers are not reused

Delete workbook #2 of two, and the next one you create is **#3**. The gap stays.

That is deliberate: #2 may name a PDF you have already emailed to a customer, and reusing the number would make that document describe a different repair.

Same for worklog numbers, and deleting one does not renumber the rest.
