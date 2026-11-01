[Wiki Home](Home) · [Data files](Explanation-of-data-files)

The raw KiCad files that make a board's traces clickable.

---

A board's KiCad project is a set of its files whose names start with `KiCad data/`, and it is what
makes the traces on a board clickable: click a pin or a trace on a picture, and the whole copper net
lights up, across every picture lined up with a KiCad view.

```
KiCad data/C64-250469-KiCad.kicad_pcb
KiCad data/C64-250469-KiCad.kicad_sch
KiCad data/Pages/vic.kicad_sch
```

## Adding it to a board

On the "Contribute" tab, drop the KiCad project's **folder** onto the tab - or open **Files** under
**KiCad** on the left (it shows how many KiCad files the board has) and import it there, on the right
of the tab. CRT copies in only the files it reads, and says what it found (components,
nets, sheets), what it left out, and what is left to do.

While **Add the KiCad components to the board** is ticked (it is, until you untick it), the import
also adds every component of the KiCad board that the board does not have yet, with its package and
pins where the footprint says them - one step, so **Undo** takes them all back. Imported without it,
the summary offers **Add N components from KiCad**.

From there CRT does most of the rest for you:

* each picture card's **Shows KiCad view** ties a picture to a KiCad view (a side of the board, or a
  schematic sheet) - and **Overview** suggests the pairs it can tell from the names
* **Align KiCad data** on a picture card lines the KiCad view up with the picture. On a picture of a side
  of the board with nothing marked on it yet, pressing **Done** there also marks every component on
  it from its footprint - one step, so **Undo** takes it back
* **Mark all components automatically from KiCad** marks the components still missing on an aligned picture of a
  side of the board, for instance after you added more components
* each picture card sets the color and opacity of the traces on **this side** and on the **other
  side** of the board, as the "Schematics" tab draws them when a net lights up
* **Fill from KiCad** in a component's pins adds each pin's number and name, and shows the net it is
  on
* **Important signals** under **KiCad** suggests signals from the net names

## Rules

* Only `.kicad_pcb`, `.kicad_sch` and `.kicad_pro` are read (KiCad 6 and newer). An older KiCad's
  `.brd` / `.sch` is not - open the project in a newer KiCad and save it again. EAGLE's files use the
  same extensions; import those into KiCad first.
* Sub-folders are read too, so a multi-sheet project can keep its pages in, say, `Pages/vic.kicad_sch`.
  KiCad's own `*-backups` folders, and files and folders whose name starts with a dot, are skipped.
* Every sheet of a multi-sheet project must be a `.kicad_sch` file of the project. A sheet the
  schematic points to somewhere else - another drive, a network share, or out of the folder with
  `..` - is not read, and a submission carrying such a schematic is refused. Keep the whole project
  in one folder and save it again from KiCad.
* Keep exactly **one** `.kicad_pcb`. It is what components are matched against to light up their
  copper, and with several only the first is used.
* Footprint libraries, 3D models, gerbers and backups are not taken - everything here is downloaded
  by every user.

A board does not need KiCad data. Without it the board works as normal, just without clickable
traces.
