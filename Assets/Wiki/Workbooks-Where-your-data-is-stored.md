# Where your data is stored

← [Back to Workbooks](Workbooks)

Workbooks are **entirely local**. They are never uploaded, never synced, and are not part of the
hardware data the app downloads from `classic-repair-toolbox.dk`.

---

## The location

Workbooks live in a **`Workbooks`** folder inside the app's own folder in your user profile,
alongside your settings file and the log:

| OS | Path |
| --- | --- |
| **Windows** | `%LOCALAPPDATA%\Classic-Repair-Toolbox\Workbooks\` |
| **Linux** | `~/.local/share/Classic-Repair-Toolbox/Workbooks/` |
| **macOS** | `~/.local/share/Classic-Repair-Toolbox/Workbooks/` |

There is a button that takes you straight there:
**Configuration → "Open data/workbooks/log/settings folder"**.

This folder is **outside** the synced `Data` folder on purpose, and it survives application updates.

---

## The layout

One folder per workbook, named after its id. Everything belonging to that workbook is inside it.

```
Workbooks/
├── counters.json               ← the highest workbook id ever handed out
├── 1/
│   ├── index.json              ← the workbook itself
│   ├── entries.json            ← all of its worklogs
│   ├── worklog_1/              ← worklog #1's photos and files
│   │   ├── 5v-rail.png
│   │   └── 7805-datasheet.pdf
│   └── worklog_3/
│       └── vic-socket.jpg
├── 2/
│   ├── index.json
│   └── entries.json
└── 4/                          ← #3 was deleted; the gap is deliberate
    └── index.json
```

**There is deliberately no central index.** Every lookup scans the folders, so there is no
bookkeeping file to fall out of sync or go stale. A folder without a readable `index.json` is simply
skipped.

### The files

| File | Holds |
| --- | --- |
| `index.json` | One workbook: id, board, description, note, status, start date, worklog count, and its entry-id counter |
| `entries.json` | Every worklog in that workbook, with all their lists — links, comments, work done, and the *metadata* for photos and files |
| `worklog_{id}/` | The actual photo and file **bytes** for that worklog |
| `counters.json` | The highest workbook id ever handed out — see [Deleting things](Workbooks-Deleting-things#why-numbers-are-never-reused) |

All the JSON is plain, indented and readable. Writes are atomic, so an interrupted save cannot leave
a half-written file.

---

## Backing up

**Copy the `Workbooks` folder.** That is the whole backup — there is nothing stored anywhere else.

To restore, put the folder back. To move to another machine, copy it across. Nothing is registered
elsewhere and nothing needs importing.

> Close the app before copying, or you may catch a file mid-write.

---

## Using a different folder

Launch with `--workbooks-root=` to point the app somewhere else:

```
Classic-Repair-Toolbox.exe --workbooks-root="D:\Repairs\Workbooks"
```

This works exactly like `--data-root=` does for the synced hardware data: case-insensitive,
surrounding quotes stripped, first match wins, and the folder is created if it does not exist.

Useful for keeping workbooks on a synced drive, on a NAS, or beside the rest of a workshop's records.

Note that the app reads **one** workbook root per launch — it does not merge several.

---

## Editing the files by hand

You can. The formats are plain JSON and the app reads whatever it finds. A few things to know:

- **A folder must be named after the id inside its `index.json`**, or the workbook will not be found.
- **Photo and file bytes live in `worklog_{id}/`**, and the names in `entries.json` must match. An
  attachment whose file is missing is skipped in exports rather than breaking them.
- **A folder renamed from an older build's naming will lose its attachments.** Earlier versions
  called these folders `entry-{id}-files`; those are not migrated, and their attachments simply stop
  being found.
- **Do not renumber things to close a gap.** See
  [why numbers are never reused](Workbooks-Deleting-things#why-numbers-are-never-reused).

Close the app before editing, and keep a copy.

---

**Next:** [Deleting things](Workbooks-Deleting-things) · [Troubleshooting and FAQ](Workbooks-Troubleshooting-and-FAQ)
