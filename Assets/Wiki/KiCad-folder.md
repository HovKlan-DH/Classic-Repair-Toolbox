# KiCad folder

The raw KiCad files that make a board's traces clickable.

[Wiki Home](Home) · [Data files](Explanation-of-data-files)

---

A folder named exactly `KiCad data`, placed directly inside a board folder. It holds the raw KiCad files, and it is what makes the traces on a board clickable.

```
Data/Commodore/C64/250469/
└── KiCad data/
    ├── C64-250469-KiCad.kicad_pcb
    └── C64-250469-KiCad.kicad_sch
```

Rules:

* Only `.kicad_pcb`, `.kicad_sch` and `.kicad_pro` are read (KiCad 6 and newer).
* Nothing to register. The application finds the folder on its own.
* Do not add footprint libraries, 3D models, gerbers or backups. Everything here is downloaded by every user.

The board is not required to have this folder. Without it the board works as normal, just without clickable traces.

See [How to add a new board with KiCad data](Add-new-board-with-KiCad-data) for the full walkthrough.
