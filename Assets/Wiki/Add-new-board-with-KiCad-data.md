[Wiki Home](Home)

The full walkthrough: from a bare board photo to clickable traces.

---

This page walks through contributing a **complete new board** whose schematics and PCB views are backed
by a KiCad project, so that selecting a component highlights its actual copper traces, nets can be
hovered, and pin 1 can be marked.

A board without KiCad data is the same job minus steps 5–7. Everything else applies either way.

---

## What you are building

A single self-contained folder in the data tree, plus one line registering it in the main Excel data file:

```
<data root>/
├── Classic-Repair-Toolbox.v2.0.0.xlsx                  <- the main Excel data file (do not edit this directly)
├── Classic-Repair-Toolbox.v2.0.0_UserContribution.xlsx <- the new add-on main Excel file that YOU create (step 4)
└── Commodore/C64/250407/                               <- your board folder containing all your data
    ├── Data C64 250407 v2.0.0.xlsx                     <- board Excel data (step 2)
    ├── Data C64 250407 v2.0.0.json                     <- highlights + calibration (steps 6-7)
    ├── Board Layout 250407 NTSC.png                    <- schematic image (one or several)
    ├── C64 250407 PCB Replica.1.1 Top.png              <- schematic image (one or several)
    ├── C64 250407 PCB Replica.1.1 Bottom.png           <- schematic image (one or several)
    ├── C64 250407 PCB Replica 1of2.png                 <- schematic image (one or several)
    ├── C64 250407 PCB Replica 2of2.png                 <- schematic image (one or several)
    ├── KiCad data/                                     <- folder containing raw KiCad files (step 3)
    │   ├── 250407_.kicad_pcb
    │   ├── 250407_.kicad_sch
    │   └── 250407_sheet2.kicad_sch
    └── Scope baseline/                                 <- optional folder for oscilloscope baseline images
```

The C64 `250407` board is the reference implementation — when in doubt, copy what it does.

---

## Before you start: what makes a KiCad project usable here

The app parses the raw KiCad files itself. There is no conversion step and no export to prepare, but
the project has to satisfy four things:

1. **Modern KiCad files.** Only `.kicad_pcb`, `.kicad_sch` and `.kicad_pro` are read (KiCad 6 and newer
   S-expression format). Legacy `.brd` / `.sch` files are ignored.
2. **Reference designators must match your board labels.**

   > [!WARNING]
   > A component labelled `U17` finds its copper by looking for a footprint named `U17`. Call it
   > `PLA/U17` in your workbook and nothing will highlight — with no error anywhere. Matching is
   > case-insensitive and trimmed, but otherwise exact. This is the number one cause of
   > "I did everything and no traces appear".
3. **Nets should be named.** Net names are what the overlay groups copper by, and they are what you put
   in the `Important signals` sheet. KiCad's auto-generated names (`Net-(U1-Pad3)`) work, but they are
   useless to a human reading the signal list.
4. **One image per view you want interactive.** The KiCad data is drawn *on top of* an image — a PCB
   render, a photo, or an exported schematic sheet. No image, no overlay.

Size is not a blocker: existing boards ship a 45 MB `.kicad_pcb`. The project is parsed in the
background while the schematic image is already on screen, which is what the "KiCad data initializing…"
indicator in the bottom-right corner is telling you.

---

## Step 1 — Create the board folder

Board folders live under the data root as `<Manufacturer>/<Hardware>/<Board>/`. The data root is the
`Data` folder next to the executable, unless you started the app with `--data-root=<path>`.

---

## Step 2 — Build the board Excel data file

Copy an existing board `.xlsx` and empty the rows (maybe keep some for examples).\
Keep the sheet names exactly, but reference the documentation for the many sheets and columns, [Explanation of data files](Explanation-of-data-files) and [Board Excel](Board-Excel)

---

## Step 3 — Drop the KiCad files into `KiCad data/`

Create a folder named exactly **`KiCad data`** directly inside the board folder. The app discovers its
contents automatically — there is nothing to register.

* **Top level files only.** Files in sub-folders are not picked up.
* **Copy all KiCad files with extensions:** `*.kicad_*`.
* **Do not include the file `KiCad-traces.json`.** Some older existing boards still uses this from an older
  conversion pipeline; the application no longer reads it.
* **Ship only what is needed.** Every file in this folder is synced to every user of the app. Footprint
  libraries, 3D models, gerbers and backups do not belong here.

---

## Step 4 — Register the board so the app can see it

Boards appear in the dropdowns because they are listed in the main Excel data file - which you do not edit, as it will be overwritten at the next data synchronization!\
Instead, the application merges in the **personal contribution Excel data file** that you create in the data root:

```
Classic-Repair-Toolbox.v<version>_UserContribution.xlsx
```

Same name as the main Excel data file that the application resolved, with `_UserContribution` before the extension
(e.g. `Classic-Repair-Toolbox.v2.0.0_UserContribution.xlsx`).

The content of the user contribution file must be same format as the main Excel data file, but of course only with your data in it (one new board). View the format here, [Explanation of data files](Explanation-of-data-files) and [Main Excel](Main-Excel).

Its entries are merged into the hardware/board dropdowns at startup (an entry that duplicates one from
the main Excel data file is skipped, with a warning in the log).

> [!TIP]
> This file does more than list your board. Everything it references — the board workbook, its JSON
> sidecar, every file the workbook points at, and the whole `KiCad data/` folder — becomes
> **protected**: the online sync will _not_ overwrite it and the orphan cleanup will not delete it. This
> is how you work on a board for weeks without the online sync ruining your edits. The sync
> banner tells you how many files are being protected.

---

## Step 5 — Find the real CAD names, then fill in the column

The `CAD name` value must match a KiCad *view* the application built from your files. Views are generated like
this:

| View | Display name |
| --- | --- |
| PCB, top side | `<pcb file base name> - PCB Top` |
| PCB, bottom side | `<pcb file base name> - PCB Bottom` |
| A schematic sheet | Its sheet name, or the file base name when it has none |

So `250407_.kicad_pcb` yields `250407_ - PCB Top` and `250407_ - PCB Bottom`, while `250407_sheet2.kicad_sch`
yields `250407_sheet2`.

**Do not guess them — read them from the logfile.** Start the application, select your board, and look for:

```
KiCad information for [Commodore 64] [250407]:
  File [250407_.kicad_pcb]; nets [249], footprints [...], pads [...], segments [...], vias [...]
    Display name [250407_ - PCB Top]; type [pcb_top], source_kind [pcb], id [pcb:0:top]
    Display name [250407_ - PCB Bottom]; type [pcb_bottom], source_kind [pcb], id [pcb:0:bottom]
  File [250407_.kicad_sch]; wires [...], polylines [...], labels [...], symbols [...]
    Display name [250407_]; type [schematic], source_kind [schematic], id [schematic:0]
```

Copy those `Display name` values verbatim into the `CAD name` column. The same block also tells you
whether the parse worked at all — zero nets or zero footprints means the file did not give up its data.

There are two conveniences if `CAD name` is blank: a view named exactly `Top (replica)` or
`Bottom (replica)` falls back to the first PCB top/bottom view, and a name shaped like `#1 of 2` falls
back to the *n*-th schematic sheet. **Fill in `CAD name` explicitly anyway** — the fallbacks only cover
one PCB file and depend on sheet order.

> [!WARNING]
> Nothing validates the `CAD name`. A typo produces no error, no warning and no overlay — it simply
> behaves as if the board had no KiCad data at all. If traces do not appear, suspect this column first.

---

## Step 6 — Calibrate each KiCad-backed view

KiCad works in millimetres and your image is pixels, so each view needs a one-time alignment. Fill in
`CAD name` **before** calibrating — the calibration is stored together with the CAD name it was made for.

1. Schematics tab → settings panel → tick **Enable contributor mode**.
2. Select the view you want to calibrate.
3. **Right-click** the image (a click, not a drag) → **Calibrate KiCad traces**.
4. Drag the calibration box so the KiCad outline lands on the board in the image. For a view showing the
   board from the other side, drag the box inside-out to mirror it.
5. **Apply KiCad calibration**.

This writes a `KiCad calibration points` entry into the board JSON — offset, scale and mirror flags for
that one view. See [Board JSON](Board-JSON) for what the stored numbers mean. Nothing there is worth
hand-editing; recalibrating is faster and correct.

Repeat for every view that has a `CAD name`.

---

## Step 7 — Label the components

With contributor mode still on: right-click → **Enable component label editor**, draw a rectangle around
each component, and press **Apply all editor changes**. Rectangles go into the board JSON; brand-new
board labels are appended to the workbook's `Components` sheet automatically.

The full rules live in [Board JSON](Board-JSON). The one that matters most here: **the board label you
type must be the KiCad reference designator**, or the component will highlight on the image but light up
no copper.

There is a walkthrough video: [How to use component label editor](https://youtu.be/u-UkD-m4Z6o)

---

## Step 8 — Important signals (optional, KiCad-only)

The `Important signals` sheet powers the side panel that lets a user light up a whole supply or clock
net without hunting for a component:

| Display name | KiCad net name |
| --- | --- |
| `+5VDC` | `+5V` |
| `+5VDC CAN` | `CAN+5V` |

* You may write either the full KiCad net name or its last path segment — `/CPU/PHI2` and `PHI2` both
  resolve. Matching is case-insensitive.
* Several rows may share a display name; they become one entry lighting up all of those nets.
* A row that matches no net in the project is skipped silently in normal use — **but in contributor
  mode the logfile names it**:

  ```
  Important signals debug: Excel KiCad net name [PH12] for display name [PHI2] did not match any
  loaded KiCad net name
  ```

  Contributor mode also logs the total counts and any duplicate mappings. Leave it on while you wire up
  a new board.

---

## Step 9 — Verify your edits

The logfile is the review tool for this whole job — the application is deliberately forgiving at runtime, so
most mistakes are a warning in the logfile rather than a visible failure. Make sure there are no warnings visible in the logfile for the board you launch with.

---

## Troubleshooting

| Symptom | Cause to check first |
| --- | --- |
| No overlay at all on any view | `CAD name` typo, files not directly in `KiCad data/`, or a non-modern KiCad format |
| Overlay on PCB views but not schematic views | The schematic sheets were not loaded, or their sheet names differ from what you put in `CAD name` — check the log listing |
| Traces appear but are offset or stretched | Calibration missing or stale for that view |
| Component highlights, but no copper lights up | Reference designator ≠ board label |
| Only some nets ever light up | Those nets are unnamed in KiCad, or the component's pads have no net assignment |
| Signal missing from the Important signals panel | Net name does not resolve — turn on contributor mode and read the log |
| "Mark first pin" is not offered | Only PCB views carry pad data; schematic views cannot mark pin 1 |
| "KiCad data initializing…" for a long time | Normal on a large `.kicad_pcb`; it loads in the background |
| Your files were replaced or vanished after a sync | The board is not listed in your `_UserContribution` workbook, so it was not protected |

---

## Submitting your data

When you select a whole new board/system, then you should follow [Contribute data GitHub](Contribute-data-via-GitHub) as this is a major new thing.

**Make sure you submit data in a good quality!** No one wants to see a rough and fast implementation, as this only gives frustration when missing something or something is plain wrong. This is a fine balance though, because "quality data" is a never ending story and it can always improve, but at least do your best to satisfy the many users that will benefit from your data 🙏

## Connection with application developer

There can be many causes why data misbehaves, so if e.g. you have some KiCad data not showing as expected or you do not understand something (probably due to none or wrong documentation), then please do not hesitate to connect with the developer. [View GitHub page for contact](https://github.com/HovKlan-DH/Classic-Repair-Toolbox#contact-developer).
