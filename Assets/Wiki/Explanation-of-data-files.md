[Wiki Home](Home)

Everything CRT knows about hardware is data, not code — so adding a board is a data job, not a programming job. This page shows which file holds what.

---

## The layout

```
<data root>/
├── Classic-Repair-Toolbox.v2.0.0.xlsx                    ← every hardware + board, and every scope
├── Classic-Repair-Toolbox.v2.0.0_UserContribution.xlsx   ← your own boards (you create this)
│
└── Commodore/C64/250407/                                 ← one folder per board
    ├── Data C64 250407 v2.0.0.xlsx                       ← what the components ARE
    ├── Data C64 250407 v2.0.0.json                       ← WHERE they are on the images
    ├── Board Layout 250407 NTSC.png                      ← schematic images
    ├── C64 250407 PCB Replica.1.1 Top.png
    │
    ├── KiCad data/                                        ← optional — makes traces clickable
    └── Scope baseline/                                    ← optional — scope images from a good board
```

The C64 `250407` board is the reference implementation. **When in doubt, copy what it does.**

## Which file do I need?

| I want to | File |
| --- | --- |
| Add my board to the drop-downs | [Main Excel](Main-Excel) — or rather, your own `_UserContribution` copy of it |
| Say what a component is — name, value, part number, datasheet | [Board Excel](Board-Excel) |
| Make a component light up on a schematic image | [Board JSON](Board-JSON) |
| Make the copper traces clickable | [KiCad folder](KiCad-folder) |
| Add scope readings from a known good board | [Scope baseline folder](Scope-baseline-folder) |
| Do all of the above for a brand new board | [Add a new board with KiCad data](Add-new-board-with-KiCad-data) — the walkthrough |

## The one thing to understand

Two files in a board folder share a name — the `.xlsx` and the `.json`:

> **The Excel file says *what* a component is. The JSON file says *where* it is.**

CRT finds the JSON by taking the Excel path and swapping the extension. So if you rename or version-bump one, **rename the other in the same go.**

> [!WARNING]
> Nothing crashes if you forget. The board still opens and the component list is still complete — but every highlight on every schematic is silently gone, and the only trace is one `Board highlight JSON file not found` line in the logfile.

## Rules for every Excel file

These apply to [Main Excel](Main-Excel) and [Board Excel](Board-Excel) alike.

* **No empty rows in the middle of your data.** The first blank row is treated as the end.
* **All paths use `/`, never `\`** — so they work on Linux and macOS too.
* **Folder and file names are case-sensitive**, for the same reason.
* **A yellowish cell** means that value still needs checking or correcting by hand.
* **No formatting carries over.** Colours, bold and italic in Excel do nothing in CRT.
* **Read the logfile after startup.** It names every problem it found in your data.

## If it does not work

| What you see | Almost always means |
| --- | --- |
| Your board is not in the drop-downs | It is not listed in your `_UserContribution` file, or the path to the board Excel is wrong |
| Board opens, but nothing highlights | The `.json` name no longer matches the `.xlsx` — check the logfile |
| One component never highlights | Its label is spelled differently in the two files, or its `Region` excludes yours |
| Components highlight, but no copper lights up | The board label is not the KiCad reference designator |
| Your files vanished after a sync | The board is not in your `_UserContribution` file, so it was not protected |

The application is deliberately forgiving: bad data is a warning in the logfile, not a crash. **That makes the logfile the review tool for this whole job.**
