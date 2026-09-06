[Wiki Home](Home) · [Workbooks](Workbooks-tab)

Exporting a repair, where the files live, and deleting.

---

## Export

Two buttons in the "Workbooks" tab header, acting on the selected workbook.

| | PDF | ZIP |
| --- | --- | --- |
| The document | Yes | Yes - the same PDF, inside |
| Photos | Embedded, page sized | The original files |
| Attached files | No | Yes |
| Use it for | Sharing the write-up | Handing over everything |

Both are named like this, and you can rename it in the save dialog:

```
Workbook_3_Commodore_C64_20260904.pdf
```

The workbook description is deliberately not in the file name - it often holds personal details, on a file you may be about to send to someone.

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

> **Upgrading from an older version:** the layout changed and **nothing is converted automatically**. Older versions kept every worklog of a workbook in one `entries.json`, and named workbook folders by a bare number (`1` instead of `workbook_1`). Those are no longer read, so workbooks written by an older version will not appear. Keep a copy of the folder before upgrading if you need that data. The same applies to the older `counters.json`, which is no longer read and can be deleted.
>
> Two fields inside a workbook's `index.json` were renamed at the same time, for the same reason the folders were: `entryCount` and `lastEntryId` are now `worklogCount` and `lastWorklogId`, so the file says "worklog" like the rest of the application. These are also not converted - a workbook from an older version shows both as `0` until its next change, and no worklog number is ever re-used as a result.

**To back up: copy the `Workbooks` folder.** That is all of it. To move to another machine, copy it across - nothing else needs doing.

Close the application first, so you do not catch a file mid-write.

You can put the folder somewhere else with `--workbooks-root=`, see [Commandline parameters](Commandline-parameters).

## Deleting

Both are permanent - there is no undo.

**Delete a worklog** - the ✕ on its card. Removes the worklog and its photos and files.

**Delete a workbook** - the button in the header. Removes the whole repair: every worklog, photo and file in it.

Export to ZIP first if there is any chance you want it back.

### Numbers are not reused

Delete workbook #2 of two, and the next one you create is **#3**. The gap stays.

That is deliberate: #2 may name a PDF you have already exported and sent on, and reusing the number would make that document describe a different repair.

Same for worklog numbers, and deleting one does not renumber the rest.
