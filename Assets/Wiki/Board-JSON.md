[Wiki Home](Home) · [Data files](Explanation-of-data-files)

Where each component sits on each schematic image. Written by the app, not by hand.

---

Every board folder in the data tree holds two files that share the same name - one `.xlsx` and one `.json`:

```
Assets/Data/Commodore/C64/250407/
├── Data C64 250407 v2.0.0.xlsx   <- the board's text data
└── Data C64 250407 v2.0.0.json   <- the board's coordinates   (this page)
```

The short version: _the Excel file says **what** a component is, the JSON file says **where** it is._

Nothing in the JSON is descriptive. There are no names, no values, no datasheets, no links - it is pure
geometry: the rectangle that lights up when you select `U17` in the component list, and the alignment
that makes a KiCad copper overlay sit exactly on top of a board photo.

---

## The file name is not a free choice

The app never searches for this file by name. It takes the board Excel path that is registered in the
main Excel file (`Classic-Repair-Toolbox.v2.0.0.xlsx`, sheet `Hardware & Board`, column `Excel data file`) and
simply swaps the extension for `.json` - same folder, same base name, every time. A draft works the same
way, beside the draft's own Excel file.

| Board Excel file | Board JSON file the app will look for |
| --- | --- |
| `Commodore/C64/250407/Data C64 250407 v2.0.0.xlsx` | `Commodore/C64/250407/Data C64 250407 v2.0.0.json` |
| `Amstrad/CPC 664/MC0005A/Data CPC 664 MC0005A v2.0.0.xlsx` | `Amstrad/CPC 664/MC0005A/Data CPC 664 MC0005A v2.0.0.json` |

> [!IMPORTANT]
> If you rename or version-bump the board Excel file, **rename the JSON in the same go**. Nothing
> crashes if you forget: the board still opens, the component list is still complete - but every
> highlight on every schematic is silently gone. The logfile then has a
> `Board highlight JSON file not found` line, and a "not marked on any schematic" warning for every
> component.

### A note on older board files

In board files from before `v2.0.0` this data lived in a `Component highlights` worksheet inside the
workbook itself. From the `v2.0.0` board files onward that sheet is gone and the JSON sidecar is the
only source of truth. If you are looking at an old file, or porting old data forward, that is why the
sheet exists in one and not in the other.

---

## Structure

The file has two independent top-level sections ("roots"), both optional:

```json
{
  "Component highlights": { },
  "KiCad calibration points": { }
}
```

Most boards only have the first. Saving one never disturbs the other.

---

## `Component highlights`

Three levels of nesting: **schematic name -> board label -> a list of rectangles.**

```json
{
  "Component highlights": {
    "Board layout": {
      "C1": [
        { "X": 1448, "Y": 996, "Width": 102, "Height": 36 }
      ],
      "U17": [
        { "X": 2210, "Y": 1512, "Width": 148, "Height": 258 }
      ]
    },
    "Schematics #1 of 2": {
      "U20": [
        { "X": 3960, "Y": 2854, "Width": 147, "Height": 256 },
        { "X": 1239, "Y": 4032, "Width": 148, "Height": 258 }
      ]
    }
  }
}
```

### Both names are joins into the Excel file

| JSON key | Must match |
| --- | --- |
| Schematic name (`"Board layout"`) | Sheet `Board schematics`, column `Schematic name` |
| Board label (`"C1"`, `"U17"`) | Sheet `Components`, column `Board label` |
| A `KiCad calibration points` key (`"Top (replica)"`) | Sheet `Board schematics`, column `Schematic name` |

Matching is case-insensitive and surrounding spaces are trimmed, but otherwise the strings have to be
identical. A label that exists in one file and not in the other is an *orphan*, and CRT warns about it
in **both** directions. In a draft's table on the "Drafts" tab the warning is an amber corner on the
component's row (or, for a highlight with no component, a line above the table); for every board, the
logfile gets a line at each launch:

```
Excel data file [...] sheet [Components] row [12] column [Board label] has warning
[component.no_highlight]: Component [C42] is not marked on any schematic, so CRT cannot
point it out on the board. Mark it with the label editor on the Schematics tab. - please fix!

Excel data file [...] JSON file [...] has warning [highlight.orphan]: Schematic
[Board layout] marks component [C42], which is not in the Components sheet. - please fix!
```

Orphans are the most common defect in contributed board data, and they are cheap to find - open the
draft's table, or run the app once with your data and read the logfile.

A highlight on a schematic name that is not in `Board schematics` is not just a warning: it is an
error, and the draft cannot be submitted until it is fixed.

### Coordinates are source-image pixels

`X`, `Y`, `Width` and `Height` are pixels in the **full-resolution image** for that schematic (the file
named in `Board schematics` -> `Schematic image file`), with the origin at the top-left corner.

* They are **not** percentages, **not** screen pixels, and **not** relative to your current zoom. That is
  exactly why the same numbers work in the main view, at any zoom level, and in the thumbnails.
* They are always **whole numbers** - the editor rounds when it saves.
* A rectangle with no width or height, one starting outside the image (a negative `X` or `Y`), or one
  whose numbers cannot be read is an error, and the draft cannot be submitted until it is fixed.
* Consequently: if you replace a schematic image with a scan at a different resolution, or crop it,
  **every rectangle on that image becomes wrong**. Re-cropping an image is a re-labelling job, so get the
  image right before you start drawing.

### One component can have several rectangles

The value is a list, not a single object. If a component appears twice on the same sheet - a chip drawn
in two places on a schematic, a part mounted on both sides - give it two rectangles, as `U20` above.
They highlight and blink together as one component.

### PAL / NTSC does **not** belong in this file

There is no region field here, and you should not fork rectangles per region. Region lives on the
component rows in the Excel `Components` sheet, and the rule is:

* The label's rows have a blank `Region` -> highlighted in **every** region.
* The label's rows name regions (one region per row) -> highlighted **only** in those.

One label, one set of rectangles; the Excel rows decide whether it is drawn.

### Let the app format the file

The app writes the file, and the server writes it in its own layout when it publishes new highlights.
Spacing and sorting do not change what a maintainer is shown or what is published - the highlights
themselves are compared - so there is nothing to gain from re-sorting or pretty-printing it by hand.

---

## `KiCad calibration points`

Only present for boards that ship a `KiCad data/` folder. One entry per schematic view, recording how the
KiCad geometry has to be stretched to sit on top of *that particular image*:

```json
{
  "KiCad calibration points": {
    "Top (replica)": {
      "CadName": "250407_ - PCB Top",
      "OffsetX": 50.87193460490473,
      "OffsetY": 451.9972752043597,
      "ScaleX": 0.9769080633767282,
      "ScaleY": 0.812766517468314,
      "MirrorX": false,
      "MirrorY": false
    }
  }
}
```

| Field | Meaning |
| --- | --- |
| `CadName` | Which KiCad view this image was matched against, e.g. a PCB layer or a schematic sheet |
| `OffsetX` / `OffsetY` | Top-left corner of the calibration box, in image pixels |
| `ScaleX` / `ScaleY` | Box size divided by full image size (`1.0` = the box covers the whole image) |
| `MirrorX` / `MirrorY` | The box was dragged inside-out - used when the image shows the board from the other side |

These numbers are the result of dragging a box in the app until the copper lines up with the image, so
they come out with fifteen decimals and are meaningless in isolation. **Do not hand-edit them.** If an
overlay drifts, redo the calibration; it takes a minute and it is the only reliable way to get it right.

---

## How you actually produce all this: you do not type it

This file is written by the app. Editing it by hand is possible, but it is not the intended workflow and
it is very easy to produce something that is valid JSON and still wrong.

1. "Schematics" tab -> in the settings panel, tick **Enable contributor mode**.
2. **Right-click** the schematic image (a click, not a drag) to open the floating menu.
3. Choose **Enable component label editor**, or **Calibrate KiCad traces** for the KiCad root.

**Component label editor:**

4. Drag to draw a rectangle; drag the handles to resize; drag the middle to move; right-click a rectangle
   to delete it. A brand-new rectangle asks you for a board label and a category.
5. Press **Apply all editor changes**.

What that save actually does:

* Saves into your own local **draft** for this board, not into the board's Excel file or JSON directly -
  nothing you edit here changes what anyone else sees until you submit it, and the [Drafts tab](Drafts-tab) lists
  every board you have local changes on. See [Command-line parameters](Commandline-parameters) for where
  drafts are stored, and the Configuration tab's "View boards as officially coming from online source" option ([View boards from online source](View-boards-from-online-source)) to check
  what the board looks like without your edit.
* Replaces the highlights for the **currently selected schematic only** - every other schematic keeps
  whatever your draft already has for it (to begin with, the published highlights).
* Drafts a new component entry for any board label that did not exist yet. If you introduced no new
  labels, nothing is drafted for the `Components` sheet at all.
* Refuses to save, with a message, if a rectangle has no label, no category, or zero size, or if the same
  label was given two different categories.

> 🎬 There is a walkthrough video: [How to use component label editor](https://youtu.be/u-UkD-m4Z6o)

**Calibrate KiCad traces:**

4. Drag the box (or its handles) until the KiCad copper lines up with the image; drag across an edge to
   flip the board.
5. Press **Apply KiCad calibration** to save, or **Discard KiCad calibration** (or **Esc**) to cancel
   without saving.

Saving writes one calibration entry into your local **draft** for this board, the same as the label
editor above - not into the board JSON directly. It replaces only the calibration for the currently
selected schematic view; every other view's calibration, official or already drafted, is untouched.

---

## Symptoms and their causes

| What you see | Almost always means |
| --- | --- |
| No highlights at all on a board, everything else fine | JSON missing, renamed out of sync with the `.xlsx`, or malformed - check the logfile |
| One component never highlights | Label typo (join mismatch), or its `Region` in the Excel excludes the current region |
| Highlights land on the wrong schematic, or nowhere | Schematic name key does not match `Board schematics` -> `Schematic name` |
| Every rectangle on one image is shifted or scaled | The image was replaced or cropped after the labelling was done |
| KiCad copper drifts away from the image | Calibration is stale - recalibrate that view |
| Orphan warnings in the table or the logfile | Excel and JSON disagree about which components exist |

A missing or broken JSON never blocks the app. It logs a warning, treats the board as having no
highlights, and carries on - convenient at runtime, and easy to miss while contributing. Read the log.

---

## Contributing checklist

* **Keep the `.xlsx` and the `.json` together**, with identical base names. In a draft CRT does that for
  you, and **Submit** sends both - the highlights and calibrations travel with the rest of the board.
* **Let the app write the JSON** with the label editor and the calibration, rather than by hand.
* **Look at the draft's table before submitting** - orphan components, missing files and anything
  the server would refuse are marked there (see [Contribute data via CRT](Contribute-data-via-CRT)),
  and the logfile has the same for every board at each launch.
* **Do not put descriptive data here.** Names, values, part numbers, regions, links, credits and
  oscilloscope baselines all belong in the Excel file.

---

## How the file reaches other users

The board JSON is part of the online data manifest and is checksum-synced from
`classic-repair-toolbox.dk` like any other data file - independently of application releases. At startup
the app syncs the main Excel file and the board Excel files first; the board JSON arrives with the
background pass that fetches the remaining data files shortly after the window opens.

Your own work never goes through the sync. It lives in your **draft**, which the sync does not touch,
and it reaches others only after you submit it and a maintainer accepts it: it is published to the
BETA source first, and then to the stable source everyone downloads from.
