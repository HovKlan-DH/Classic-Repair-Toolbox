[Wiki Home](Home) · [Data files](Explanation-of-data-files)

The raw KiCad files that make a board's traces clickable.

---

A folder named exactly `KiCad data`, placed directly inside a board folder. It holds the raw KiCad files, and it is what makes the traces on a board clickable.

```
Data/Commodore/C64/250469/
└── KiCad data/
    ├── C64-250469-KiCad.kicad_pcb
    └── C64-250469-KiCad.kicad_sch
```

Rules:

* The folder name is `KiCad data`, spelled exactly like that - capital letters included.
* Only `.kicad_pcb`, `.kicad_sch` and `.kicad_pro` are read (KiCad 6 and newer). Legacy `.brd` / `.sch` files are ignored.
* Sub-folders are read too, so a multi-sheet project can keep its pages in, say, `Pages/vic.kicad_sch`. KiCad's own `*-backups` folders, and folders whose name starts with a dot, are skipped.
* Keep exactly **one** `.kicad_pcb`. It is what components are matched against to light up their copper, and with several only the first is used.
* Nothing to register. The application finds the folder on its own.
* Do not add footprint libraries, 3D models, gerbers or backups. Everything here is downloaded by every user.

For your own board, use the **KiCad data** button on the board's row in the [Drafts tab](Drafts-tab): it creates the folder in your draft, copies only the three file types above, and shows which components will light up and which **CAD name** each image can use. Only those three file types travel with your submission.

The board is not required to have this folder. Without it the board works as normal, just without clickable traces.

See [How to add a new board with KiCad data](Add-new-board-with-KiCad-data) for the full walkthrough, and [Board Excel](Board-Excel#column-cad-name) for the `CAD name` column that ties an image to a KiCad view.
