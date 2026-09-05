[Wiki Home](Home)

The board, its components and its traces. This is where you will spend your time.

---

## Getting around the image

* **Mouse wheel** zooms, **drag** pans
* The **thumbnails** down the side switch between schematic images and PCB photos
* **Click a component** on the image to open its popup — part number, datasheet links, local files, and any [oscilloscope baseline](Scope-baseline-folder) images
* Drag a thumbnail to reorder it

Clicking a component in the left-hand list highlights it here. Click again to unhighlight.

## Traces

If the board has [KiCad data](KiCad-folder), the copper is live: click a pin or a trace and the whole net it belongs to lights up, across every image that has been calibrated.

Two side panels sit beside the image:

* **Important signals** — the named signals worth knowing on this board, e.g. the clocks and the data bus. Click one to light up its net.
* **Netlist names** — every net on the board, when you know exactly which one you are after.

Boards without KiCad data still work normally, just without clickable copper.

## Settings

The settings panel has a **Global settings** block that applies everywhere, and per-board rows below it.

| Setting | What it does |
| --- | --- |
| **Highlight trace on hover** | Light up copper as the pointer passes over it, without clicking |
| **Hold SHIFT to highlight trace on hover** | Only hover-highlight while SHIFT is held. Useful on dense boards where everything lights up otherwise. |
| **Highlight component traces on select** | Selecting a component also lights up the copper attached to it |
| **Highlight component traces on hover** | The same, on hover |
| **Mark first pin on component** | Draw a marker on pin 1, so you can orient a chip at a glance |
| **Labels visible** | Show component labels drawn on the image |
| **Traces visible** | Show the copper overlay at all |

Per-board settings are remembered for that board.

## Drawing your own traces

You can draw trace overlays by hand on any image — useful for a board with no KiCad data, or for marking a repair you have made. Each gets its own colour, and they are listed in the panel beside the image so you can hide or delete them.

## Recording what you find

With the [Workbooks](Workbooks) feature on, **Add worklog** in the bar above the tabs lets you drag a rectangle on the schematic and write up what you found there. Markers for saved worklogs appear on the image and on the thumbnails — see [Workbooks: daily use](Workbooks-Daily-use).

## Contributor mode

Ticking **Enable contributor mode** in the settings panel unlocks two editors, both reached by **right-clicking the image**:

* **Enable component label editor** — draw or adjust the rectangles that make a component highlight. This is how a wrong or missing highlight gets fixed.
* **Calibrate KiCad traces** — line the KiCad outline up with the board in the image, so the copper overlay lands in the right place.

Both write to the board's [JSON file](Board-JSON), and both are covered step by step in [Add a new board with KiCad data](Add-new-board-with-KiCad-data).

> 🎬 There is a walkthrough video: [How to use component label editor](https://youtu.be/u-UkD-m4Z6o)

## If it does not work

| What you see | Almost always means |
| --- | --- |
| A component does not highlight | It has no highlight data yet, or the **Region** filter is excluding it |
| Clicking a trace does nothing | That board has no [KiCad data](KiCad-folder), or **Traces visible** is off |
| Traces light up but land in the wrong place | The calibration for that image is stale — recalibrate it |
| Everything lights up as you move the mouse | Turn on **Hold SHIFT to highlight trace on hover** |
| "KiCad data initializing..." | The project is still being parsed in the background. It only happens once per board load. |
