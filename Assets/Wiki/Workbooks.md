# Workbooks

**Workbooks** is Classic Repair Toolbox's repair-tracking feature. It lets you record what you found
on a board, where on the schematic you found it, what you did about it, how long it took and what it
cost — and then hand the whole job to a customer as a PDF or a ZIP.

A **workbook** is one repair job on one board. A **worklog** is one thing you noted down inside that
job — a bad capacitor, a lifted trace, an observation to come back to. A workbook holds as many
worklogs as the repair needs.

Everything is stored locally on your own machine. Workbooks are never uploaded, never synced and are
not part of the hardware data the app downloads from `classic-repair-toolbox.dk`.

---

## Pages

| Page | What it covers |
| --- | --- |
| **[Getting started](Workbooks-Getting-started)** | Turning the feature on, creating your first workbook, and adding your first worklog |
| **[Concepts and vocabulary](Workbooks-Concepts-and-vocabulary)** | Workbook vs. worklog, categories, states, the active workbook, id numbering |
| **[The worklog bar](Workbooks-The-worklog-bar)** | The strip above the tabs: the workbook picker, "Show worklogs", "Add worklog" |
| **[The Workbooks tab](Workbooks-The-Workbooks-tab)** | The full workbook browser — list, header, board pane, worklog list |
| **[The worklog editor](Workbooks-The-worklog-editor)** | The main editing window: fields, the seven lists, photos, files, components |
| **[Marking areas on a schematic](Workbooks-Marking-areas-on-a-schematic)** | Drawing an area, "Show marked area", pills, parked pills, overlays |
| **[Components in scope](Workbooks-Components-in-scope)** | The two component checklists and how they are populated |
| **[Searching](Workbooks-Searching)** | "Find a previous repair" — the query grammar and what is searched |
| **[The summary strip](Workbooks-The-summary-strip)** | Totals for a workbook, and where else they appear |
| **[Attaching oscilloscope captures](Workbooks-Attaching-oscilloscope-captures)** | Filing a scope capture straight into a worklog from the component popup |
| **[Exporting a workbook](Workbooks-Exporting-a-workbook)** | PDF and ZIP export — contents, naming, what the document looks like |
| **[Where your data is stored](Workbooks-Where-your-data-is-stored)** | Folder layout, file formats, backup, moving data, `--workbooks-root=` |
| **[Deleting things](Workbooks-Deleting-things)** | Deleting worklogs and workbooks, what is removed, and why numbers are not reused |
| **[Troubleshooting and FAQ](Workbooks-Troubleshooting-and-FAQ)** | Common questions and things that surprise people |

---

## At a glance

```
Workbook  "Dead C64, no video — customer J. Hansen"      #3, Open
│
├── Worklog #1  "VIC-II socket has a cracked pin"        Issue, Closed
│     ├── marked area on schematic "Video"
│     ├── components in scope: U19
│     ├── work done: 0.5 h, 40.00
│     └── 2 photos, 1 comment
│
├── Worklog #2  "Check the 9V AC rail"                   Note, Open
│     └── parked pill (no marked area)
│
└── Worklog #3  "Scratched case lid"                     Cosmetic, Open
```

A workbook closes itself once every worklog inside it is Closed, and reopens if you add or reopen
one. You never set a workbook's status by hand.

---

## Where things are in the app

- **Configuration tab → "Enable Worklog"** — turns the whole feature on or off. When it is off, the
  worklog bar and the Workbooks tab are both hidden.
- **The worklog bar** — a strip directly above the tab headers, on every tab. Shows the active
  workbook and carries "Show worklogs" and "Add worklog".
- **Workbooks tab** — the full browser: workbook list, board previews, worklog details, search and
  export.
- **Schematics tab** — where worklog areas are drawn and where their pills appear on the board and
  in the thumbnail gallery.
