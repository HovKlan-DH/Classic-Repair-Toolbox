[Wiki Home](Home)

The full walkthrough: from a bare board photo to clickable traces.

---

This page walks through contributing a **complete new board** whose schematics and PCB views are backed
by a KiCad project, so that selecting a component highlights its actual copper traces, nets can be
hovered, and pin 1 can be marked.

A board without KiCad data is the same job minus steps 3 and 5. Everything else applies either way.

---

## What you are building

A new board of your own, held as a **local draft** until you decide to submit it. Everything you add
goes into your own drafts folder - the application creates the folders, the board data and the
registration for you, and nothing you do here can be overwritten by the online data sync.

```
<drafts root>/
└── Commodore/C64/250407/            <- created for you when you add the system
    ├── Data C64 250407.xlsx         <- your board data, in the ordinary board format
    ├── Data C64 250407.json         <- component highlights and KiCad calibration
    ├── Board Layout 250407 NTSC.png
    ├── C64 250407 PCB Replica.1.1 Top.png
    ├── KiCad data/                  <- the raw KiCad files you imported
    │   ├── 250407_.kicad_pcb
    │   ├── 250407_.kicad_sch
    │   └── Pages/                   <- sub-folders are kept exactly as they were
    │       └── vic.kicad_sch
    ├── Scope baseline/              <- created empty for you; put scope images from a good board here
    └── .crt-draft.json              <- bookkeeping; leave this one alone
```

The `Scope baseline` folder starts out empty. It is there so that, when you add your first scope
image from a known good board, you can pick it straight away instead of creating it yourself. See
[Scope baseline folder](Scope-baseline-folder).

This is deliberately the **same shape as any published board**. The C64 `250407` board is the
reference implementation - when in doubt, look at what it does.

> [!TIP]
> **You can edit the spreadsheet directly if you prefer.** It is an ordinary board workbook, so
> opening it in Excel and typing into it does exactly what doing the same thing in the
> application does - there is no separate draft format to keep in step. Work whichever way
> suits the task: the application reads the file fresh every time, so it always sees your
> latest edits, and it never writes back over them.
>
> Close the file in Excel before saving from the application, as you would with any document
> open in two places.
>
> The one file to leave alone is `.crt-draft.json`. It records which published revision your
> draft started from, and it is never submitted or published.

> [!NOTE]
> **Your board is a real, working board on your own machine from the moment you create it.** You can
> use it exactly like any other board - browse it, label components, attach files, record worklogs -
> without ever submitting it. A system you never submit stays yours and keeps working.

---

## Before you start: what makes a KiCad project usable here

The app parses the raw KiCad files itself. There is no conversion step and no export to prepare, but
the project has to satisfy four things:

1. **Modern KiCad files.** Only `.kicad_pcb`, `.kicad_sch` and `.kicad_pro` are read (KiCad 6 and newer
   S-expression format). Legacy `.brd` / `.sch` files are ignored.
2. **Reference designators must match your board labels.** A component labelled `U17` finds its copper
   by looking for a footprint named `U17`. Call it `PLA/U17` in your board data and nothing will
   highlight. Matching is case-insensitive and trimmed, but otherwise exact.

   > [!TIP]
   > This used to be the number one cause of "I did everything and no traces appear", because nothing
   > told you. **The KiCad import now checks this for you and lists every label that will not light
   > up** - see step 5.
3. **Nets should be named.** Net names are what the overlay groups copper by, and they are what you put
   in the `Important signals` sheet. KiCad's auto-generated names (`Net-(U1-Pad3)`) work, but they are
   useless to a human reading the signal list.

   > [!TIP]
   > **A hierarchical project renames a net at every sheet boundary, and that is handled for you.**
   > If your userport sheet draws a line as `CNT1` while the PCB calls the same copper
   > `/I{slash}O/Serial Bus/CNT` - because the two sheets are wired together through pins of those
   > two names - selecting a component still highlights it on both sheets. The application follows
   > the sheet pins to work out that the two names are one net.
   >
   > The one case it deliberately leaves alone is an ambiguous one: if a sheet reaches the same wire
   > through two differently named pins, no alias is used, because highlighting the *wrong* net
   > would be worse than highlighting none.
4. **One image per view you want interactive.** The KiCad data is drawn *on top of* an image - a PCB
   render, a photo, or an exported schematic sheet. No image, no overlay.

Size is not a blocker: existing boards ship a 45 MB `.kicad_pcb`. The project is parsed in the
background while the schematic image is already on screen, which is what the "KiCad data initializing..."
indicator in the bottom-right corner is telling you.

---

## Step 1 - Add the system

On the **Contribute** tab, click **Add a new system**. (The same button is on the **Drafts** tab once
you have at least one draft.)

Fill in three names:

| Field | What it is | Example |
| --- | --- | --- |
| **Manufacturer** | The company that made it. Pick from the list if it is already there. | `Commodore` |
| **Hardware** | The machine itself. | `C64` |
| **Board** | The board revision. | `250407` |

There is a **Notes** box as well, which is optional and shows up on the Overview tab.

The box at the bottom previews exactly what will be created. Click **Create system** and you are asked
to accept the role of **maintainer**: if you later submit this system for the community to use, you will
be registered as its maintainer, and you will review the changes others submit for it and publish the
ones that are right. A new system can only be created if you accept. **Decline** takes you back to the form with nothing created.

Click **Accept and create** and the board is created, appears in the hardware and board lists straight
away, and is selected for you.

> [!TIP]
> Pick from the **Manufacturer** suggestions where you can. Typing `Comodore` instead of `Commodore`
> creates a second folder beside the real one, and nothing will warn you.

**That is the whole of what used to be steps 1, 2 and 4** - creating the folders, building a board
Excel file by copying and emptying another one, and creating a version-named `_UserContribution`
workbook to register the board in. None of that is needed any more.

---

## Step 2 - Add your board images

On the **Drafts** tab, find your system and click **Schematic images**.

Drag your image files onto the box, or use **Choose image files...**. PNG, JPG, GIF, BMP and WEBP all
work. Each image becomes one board view, named after the file, and a **copy** is taken into your
draft - your original files stay where they are.

---

## Step 3 - Import the KiCad data (optional)

On the **Drafts** tab, click **KiCad data** for your system. Drag your KiCad project folder onto the
box, or use **Choose KiCad folder...**.

* **Pick the whole KiCad project folder.** Only the `.kicad_pcb`, `.kicad_pro` and `.kicad_sch` files
  are copied, so footprint libraries, 3D models, gerbers, netlists and BOM files are all left behind -
  which is what you want, since everything imported here travels with your board when you submit it.
* **Sub-folders are searched too, and kept.** A multi-sheet project usually keeps its pages in a
  sub-folder (`Pages/vic.kicad_sch`), and those pages hold the actual circuitry - the root sheet only
  points at them. They are copied into the same sub-folder they came from, because a root sheet
  refers to its pages by relative path. KiCad's own `*-backups` folder is skipped, so the import does
  not pick up a second copy of every sheet.
* **Re-import at any time.** If you change something in KiCad, import the folder again; the files are
  replaced and the report is rebuilt.
* **Remove a file you do not want.** Every imported file is listed with its own **Remove** button.
  It deletes the copy in your draft only - your own KiCad project is not touched, and importing the
  folder again brings the file back. The report is rebuilt for the files that are left.

---

## Step 4 - Label the components

Go to the Schematics tab with your new board selected.

1. Settings panel -> tick **Enable contributor mode**.
2. Right-click the image -> **Enable component label editor**.
3. Draw a rectangle around each component, give it the board label, and press **Apply all editor
   changes**.

The full rules live in [Board JSON](Board-JSON). The one that matters most here: **the board label you
type must be the KiCad reference designator**, or the component will highlight on the image but light up
no copper.

There is a walkthrough video: [How to use component label editor](https://youtu.be/u-UkD-m4Z6o)

You can also use the **Contribute** tab to add components one at a time, with their friendly name, part
number, category, files and links.

---

## Step 5 - Check what the KiCad data lines up with

Open **KiCad data** again. Under the list of KiCad files the application tells you exactly where you stand:

| What it says | What it means |
| --- | --- |
| **N components will NOT light up** | These labels have no component of that name in the KiCad data. This is the list to fix - rename them to the bare reference designator (`U17`, not `PLA/U17`). |
| **N components match** | These will highlight copper correctly. |
| **N components are not labelled yet** | The KiCad data knows about these but your board has no label for them. This is your to-do list. |
| **CAD names you can use** | The view names the KiCad project generates - see the next step. |

You no longer have to read any of this out of the logfile.

---

## Step 6 - Fill in the CAD name for each view

Each board image that should be backed by KiCad data needs a **CAD name** naming which KiCad view it
shows. The names are listed for you in the window above - copy one across verbatim.

Views are generated like this:

| View | Display name |
| --- | --- |
| PCB, top side | `<pcb file base name> - PCB Top` |
| PCB, bottom side | `<pcb file base name> - PCB Bottom` |
| A schematic sheet | Its sheet name, or the file base name when it has none |

So `250407_.kicad_pcb` yields `250407_ - PCB Top` and `250407_ - PCB Bottom`, while `250407_sheet2.kicad_sch`
yields `250407_sheet2`.

> [!WARNING]
> A typo in a `CAD name` produces no error and no overlay - it simply behaves as if the board had no
> KiCad data at all. If traces do not appear on one view, suspect this first.

---

## Step 7 - Calibrate each KiCad-backed view

KiCad works in millimetres and your image is pixels, so each view needs a one-time alignment. Fill in
`CAD name` **before** calibrating - the calibration is stored together with the CAD name it was made for.

1. Schematics tab -> settings panel -> tick **Enable contributor mode**.
2. Select the view you want to calibrate.
3. **Right-click** the image (a click, not a drag) -> **Calibrate KiCad traces**.
4. Drag the calibration box so the KiCad outline lands on the board in the image. For a view showing the
   board from the other side, drag the box inside-out to mirror it.
5. **Apply KiCad calibration**.

This saves offset, scale and mirror flags for that one view into your own local draft. See
[Board JSON](Board-JSON) for what a `KiCad calibration points` entry looks like once it is eventually
published. Nothing in it is worth hand-editing; recalibrating is faster and correct.

Repeat for every view that has a `CAD name`.

---

## Step 8 - Important signals (optional, KiCad-only)

The `Important signals` data powers the side panel that lets a user light up a whole supply or clock
net without hunting for a component:

| Display name | KiCad net name |
| --- | --- |
| `+5VDC` | `+5V` |
| `+5VDC CAN` | `CAN+5V` |

* You may write either the full KiCad net name or its last path segment - `/CPU/PHI2` and `PHI2` both
  resolve. Matching is case-insensitive.
* Several rows may share a display name; they become one entry lighting up all of those nets.
* A row that matches no net in the project is skipped silently in normal use - **but in contributor
  mode the logfile names it**:

  ```
  Important signals debug: Excel KiCad net name [PH12] for display name [PHI2] did not match any
  loaded KiCad net name
  ```

  Contributor mode also logs the total counts and any duplicate mappings. Leave it on while you wire up
  a new board.

---

## Step 9 - Verify your work

Everything you have added shows up on the board immediately, so the real check is simply to use it:
select components, look for copper lighting up, hover nets.

Two extra checks worth doing:

* Re-open **KiCad data** and confirm nothing is left in the "will NOT light up" list.
* Untick **"View boards as officially published"** on the Configuration tab and back on again. For a
  brand-new system, ticking it shows an empty board - that is correct, because officially your system
  does not exist yet.

The logfile is still worth a glance for anything unexpected; the application is deliberately forgiving
at runtime, so some mistakes are a warning there rather than a visible failure.

---

## Troubleshooting

| Symptom | Cause to check first |
| --- | --- |
| No overlay at all on any view | `CAD name` typo, or a non-modern KiCad format. Re-open **KiCad data** and read the report. |
| Overlay on PCB views but not schematic views | The schematic sheets were not loaded, or their sheet names differ from what you put in `CAD name` - the report lists the real names |
| A component highlights most of its traces but not all | Hover the missing line: if it lights up, the net exists. On a hierarchical project this used to mean the sheets named that net differently; that is now followed automatically, so a remaining gap is worth reporting |
| Traces appear but are offset or stretched | Calibration missing or stale for that view |
| Component highlights, but no copper lights up | Reference designator != board label. **The report names every one of these** - see step 5 |
| Only some nets ever light up | Those nets are unnamed in KiCad, or the component's pads have no net assignment |
| Signal missing from the Important signals panel | Net name does not resolve - turn on contributor mode and read the log |
| "Mark first pin" is not offered | Only PCB views carry pad data; schematic views cannot mark pin 1 |
| "KiCad data initializing..." for a long time | Normal on a large `.kicad_pcb`; it loads in the background |
| Your board shows as empty | "View boards as officially published" is ticked on the Configuration tab. A draft-only system has nothing published yet, so that view is correctly blank |

---

## Boards made the old way

Before drafts existed, a new board was registered by creating a
`Classic-Repair-Toolbox.v<version>_UserContribution.xlsx` workbook in the data root by hand, listing the
board in it, and building the board's own `.xlsx` alongside. **Boards made that way keep working exactly
as they always have** and are still protected from the online sync - nothing has been taken away.

New boards are not made that way any more, and the application will not create such a file. If you have
a board set up the old way and it works, leave it alone.

---

## Submitting your data

Submitting drafts from inside the application is being built. Until it lands, a whole new board is still
best contributed through [Contribute data GitHub](Contribute-data-via-GitHub).

**Make sure you submit data in a good quality!** No one wants to see a rough and fast implementation, as
this only gives frustration when missing something or something is plain wrong. This is a fine balance
though, because "quality data" is a never ending story and it can always improve, but at least do your
best to satisfy the many users that will benefit from your data 🙏

## Connection with application developer

There can be many causes why data misbehaves, so if e.g. you have some KiCad data not showing as expected or you do not understand something (probably due to none or wrong documentation), then please do not hesitate to connect with the developer. [View GitHub page for contact](https://github.com/HovKlan-DH/Classic-Repair-Toolbox#contact-developer).
