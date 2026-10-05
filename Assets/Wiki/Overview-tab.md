[Wiki Home](Home)

Every component on the board as a list, and printable versions of it.

---

The tab shows one row per component on the selected board:

| Column | |
| --- | --- |
| **Component** | The board label - `U19`, `C64` |
| **Technical name** | The value or part type - `906114-01`, `100nF` |
| **Friendly name** | What people call it - `VIC-II`, `PLA` |
| **Part-number** | Vendor part number, where there is one |
| **Short description** | One line about what it does |
| **Notes** | The note on the component's first image, where there is one |
| **Files and links** | Datasheets and other material for that component - `F:` marks a local file, `W:` a web link. Click one to open it. |

Click the **component label** (the first column, in link colour) to open that component's popup, the
same one the [Schematics tab](Schematics-tab) opens. Clicking anywhere else in a row ticks or
unticks it for printing - see below.

The list follows the component list on the left of the window: when components are selected there,
only those are shown, and otherwise it shows what the category list and the search box leave in
view.

## Printing

Two buttons, bottom right, open a printable document in your browser, and the print dialog opens by
itself:

* **Print component list** - a checklist with the columns Component, Technical name and Friendly
  name, plus an empty column for your own notes
* **Print BOM** - a bill of materials with the columns Type, Components, Technical name, Friendly
  name and Quantity: components with the same category, technical name and friendly name share one
  line, with all their labels listed and counted

The BOM is the one to use when ordering parts: it tells you that you need four of a given capacitor
rather than listing them four times.

Both print only the rows that are shown and ticked, so the filter on the left carries over. With
nothing ticked, the buttons do nothing.

## Choosing what appears

Each row has a tick box deciding whether that component appears in the printed lists, and there is a
tick-all box in the header, which acts on the rows currently shown. Everything starts ticked, so
untick the parts you do not want - connectors and mechanical items, say - before printing.
