[Wiki Home](Home)

Everything CRT knows about hardware is data, not code - so adding a board is a data job, not a programming job. This page shows which file holds what.

---

## The layout

```
<data root>/
├── Classic-Repair-Toolbox.v2.0.0.xlsx                    ← every hardware + board, and every scope
│
├── Commodore/C64/250407/                                 ← one folder per board
│   ├── Data C64 250407 v2.0.0.xlsx                       ← what the components ARE
│   ├── Data C64 250407 v2.0.0.json                       ← WHERE they are on the images
│   ├── Board Layout 250407 NTSC.png                      ← schematic images
│   ├── C64 250407 PCB Replica.1.1 Top.png
│   │
│   ├── KiCad data/                                        ← optional - makes traces clickable
│   └── Scope baseline/                                    ← optional - scope images from a good board
│
├── Commodore/Shared files/                               ← files any Commodore board can use
└── Generic shared files/                                 ← files any board can use
```

The C64 `250407` board is the reference implementation. **When in doubt, copy what it does.**

A file that several boards use lives in a shared folder rather than in one board's folder:
`<Manufacturer>/Shared files/` for that manufacturer's boards, and `Generic shared files/` for every
board. A board's rows can point at files in its own folder and in these two shared folders. A file
in another board's folder can be used too, but only exactly as it is.

### Your own work is a draft

You never edit the data root itself - the online sync keeps it equal to the published data. Every
board you change, and every board you add with **Add a new system**, is a **draft**: a folder of the
same shape in your own drafts folder (see [Command-line parameters](Commandline-parameters) for where
that is). The sync never touches it.

```
<drafts root>/
└── Commodore/C64/250407/
    ├── Data C64 250407 v2.0.0.xlsx                       ← the same files as a published board
    ├── Data C64 250407 v2.0.0.json
    ├── ...
    └── .crt-draft.json                                    ← bookkeeping; leave this one alone
```

You send a draft with **Submit** on the [Drafts tab](Drafts-tab). Once a maintainer has accepted it,
it is published to the BETA source first, and then to the stable source everyone downloads from -
see [Contribute data via CRT](Contribute-data-via-CRT).

## Which file do I need?

| I want to | File |
| --- | --- |
| Add my board to the drop-downs | None - use **Add a new system** on the Contribute or Drafts tab. A maintainer adds it to the [Main Excel](Main-Excel) when it is published |
| Say what a component is - name, value, part number, datasheet | [Board Excel](Board-Excel) |
| Make a component light up on a schematic image | [Board JSON](Board-JSON) |
| Make the copper traces clickable | [KiCad folder](KiCad-folder) |
| Add scope readings from a known good board | [Scope baseline folder](Scope-baseline-folder) |
| Do all of the above for a brand new board | [Add a new board with KiCad data](Add-new-board-with-KiCad-data) - the walkthrough |

## The one thing to understand

Two files in a board folder share a name - the `.xlsx` and the `.json`:

> **The Excel file says *what* a component is. The JSON file says *where* it is.**

CRT finds the JSON by taking the Excel path and swapping the extension. So if you rename or version-bump one, **rename the other in the same go.**

> [!WARNING]
> Nothing crashes if you forget. The board still opens and the component list is still complete - but every highlight on every schematic is silently gone. The logfile then has a `Board highlight JSON file not found` line, and a "not marked on any schematic" warning for every component.

## Rules for every Excel file

These apply to [Main Excel](Main-Excel) and [Board Excel](Board-Excel) alike.

* **Column headers must be spelled exactly** (capital letters do not matter). CRT finds the header
  row by its names, so the rows above it and the order of the columns are free - but a sheet that
  lacks one of its required headers is read as empty, and the logfile says `Header row not found`.
* **Blank rows are skipped.** They do no harm, and CRT leaves them out when it saves the file.
* **All paths use `/`, never `\`** - so they work on Linux and macOS too.
* **Folder and file names are case-sensitive**, for the same reason.
* **Every cell is read as the text it shows. Write numbers with a dot** (`0.3`, `1.4V`), never a comma.
* **No formatting carries over.** Colours, bold and italic in Excel do nothing in CRT. When CRT saves
  a draft, or the server publishes a board, the file is written afresh: only the values in the
  documented columns are kept.
* **Check your work in the draft's table.** "Edit in table format" on the Drafts tab marks every
  problem on its cell, with the same rules the server uses when you submit. For the boards you
  have downloaded, the logfile lists the same problems at each launch.

## If it does not work

| What you see | Almost always means |
| --- | --- |
| Your board is not in the drop-downs | Its folder is not in your drafts folder, or CRT could not take it in - the logfile says why (usually several Excel files in one board folder) |
| Board opens, but nothing highlights | The `.json` name no longer matches the `.xlsx` - check the logfile |
| One component never highlights | Its label is spelled differently in the two files, or its `Region` excludes yours |
| Components highlight, but no copper lights up | The board label is not the KiCad reference designator |
| Your files vanished or changed back after a sync | They were edited in the data root, which the sync keeps equal to the published data. Work in a draft instead |

The application is deliberately forgiving: bad data is a warning, not a crash. **That makes the draft's table - and, for downloaded boards, the logfile - the review tool for this whole job.**
