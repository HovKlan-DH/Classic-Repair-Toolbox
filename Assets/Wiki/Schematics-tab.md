[Wiki Home](Home)

The board, its components and its traces. This is where you will spend your time.

---

## Getting around the image

* **Mouse wheel** zooms in and out around the pointer. **Hold the right mouse button and drag** to
  pan. On a trackpad, pinch to zoom and drag with two fingers to pan.
* A drag with the **left** mouse button on an empty spot draws a trace line of your own - see
  "Drawing your own traces" below
* The **thumbnails** down the side switch between schematic images and PCB photos. Drag a thumbnail
  to reorder it; the order is remembered for that board.
* **Click a component** on the image to highlight it and open its popup - part number, datasheet
  links, local files, and any [oscilloscope baseline](Scope-baseline-folder) images. **Right-click** a
  component to highlight it, or take its highlight away, without opening the popup.
* **Fullscreen** in the left-hand panel, or `F11`, shows the schematic view in its own maximized
  window. `Escape` or `F11` puts it back in the tab.

In the left-hand panel:

* Clicking a component in the component list highlights it here. Click again to unhighlight.
* **Clear** takes every highlight away and empties the search box; **Mark all** highlights every
  component in the list.
* **Blink selected** makes the highlighted components blink, so they are easier to spot on a busy
  board.

**Detach thumbnails into their own window** in [Configuration](Configuration-tab) moves this
thumbnail strip out of the tab into its own resizable window, giving the main image more room.
Selecting and reordering thumbnails there works the same as here, and since the window lays them
out in a grid rather than a single column, you can drag sideways too. **Right-click anywhere in
the thumbnail strip** for the same option without leaving this tab.

## The component popup

The popup shows the component's images (pinouts, oscilloscope baselines and so on) with each
image's note, its one-liner, and its local files and links.

* **PAL (n)** / **NTSC (n)** - which region's images to show in this window, with the number of
  images for each. Only shown when the board has region-specific parts.
* **Mousewheel will zoom image** - ticked, the mouse wheel zooms the image under the pointer;
  unticked, it steps through the images.
* **Synchronize oscilloscope** and **Numpad controls oscilloscope** - only shown while the
  Oscilloscope tab is enabled. See [Synchronize oscilloscope](Synchronize-oscilloscope) and
  [Controlling oscilloscope with keyboard](Controlling-oscilloscope-with-keyboard).
* **Test IC with MiniPro programmer** - for ICs CRT can test, see
  [MiniPro programmer](MiniPro-programmer).

Keys in the popup: left/right arrow for the previous/next image, type a pin number (e.g. `1` then
`4`) to jump to that pin's image, `Space` or `Enter` for the first image (with **Numpad controls
oscilloscope** ticked, `Enter` captures the scope screen instead), and `Escape` to close.
Whether each component opens its own popup is set by **Open multiple component info windows** in
[Configuration](Configuration-tab).

## Traces

If the board has [KiCad data](KiCad-folder), the copper is live. With **Highlight trace on hover** on
(the default), point at a pin or a trace and the whole net it belongs to lights up, across every
image that has been calibrated. Click it to keep it lit, and click it again - or right-click it - to
let it go. A right-click on an empty spot lets go of every lit net. The pointer's label names the
pin it is over.

Clicking copper only works while hover-highlighting is active: with **Highlight trace on hover**
off, the copper cannot be clicked, and with **Hold SHIFT to highlight trace on hover** on, hold
SHIFT while you point and click.

Two panels sit over the image:

* **Important signals** - the named signals worth knowing on this board, e.g. the clocks and the
  data bus. Click one to light up its net; **Clear All** lets them all go.
* **Netlist names** - appears while a net is lit. It names the lit net(s) and lists, under
  "Connected to:", every component pin on them (`U7 pin 3`). **Clear All** lets every lit net go.

Boards without KiCad data still work normally, just without clickable copper.

## Settings

The settings sit in panels in the bottom-left corner of the image, each with a **Collapse** button.

The **Global settings** panel applies to every board - except **Mark first pin on component**,
which each board remembers for itself. Most rows only appear when they can do something on the
image you are looking at:

| Setting | What it does | Shown |
| --- | --- | --- |
| **Highlight trace on hover** | Light up copper as the pointer passes over it. Also needed for clicking copper. | Images with KiCad data |
| **Hold SHIFT to highlight trace on hover** | Only hover-highlight while SHIFT is held. Useful on dense boards where everything lights up otherwise. | Images with KiCad data |
| **Highlight component traces on select** | Selecting a component also lights up the copper attached to it | Images with KiCad data |
| **Highlight component traces on hover** | The same, on hover | Images with KiCad data |
| **Show traces from opposite side** | Also draw the lit copper on the other side of the board | PCB images |
| **Show zones** | Draw the copper zones (filled areas such as ground planes) of lit nets | PCB images |
| **Show traces and pads** | Hide or show the copper while calibrating, to compare it with the photo underneath | Only while calibrating |
| **Mark first pin on component** | Draw an orange marker on pin 1 of the selected components, so you can orient a chip at a glance | Images with KiCad pad data |
| **Enable contributor mode** | See "Contributor mode" below | Always |

The **Labels visible** panel writes text on each component on the image: tick **Board label**,
**Technical name** and/or **Friendly name** for what to show, and **Selected components only** to
show it only on the highlighted components.

## Drawing your own traces

You can draw trace lines by hand on any image - useful for a board with no KiCad data, or for
marking a repair you have made:

* Drag with the left mouse button from an empty spot of the image (not on a component) to draw a
  line.
* Drag one of a line's points to move it, or drag anywhere along a line to add a bend there. Hold
  SHIFT while dragging to line a point up straight with the points next to it.
* Right-click a point to remove it, or right-click a line to delete the whole line.
* After drawing or moving, a small palette appears beside the line: pick its colour, choose a
  custom colour, or click the bin to delete it.

The **Traces visible** panel appears once there are traces on the image: one row per colour, with
how many lines use it and a tick box to show or hide them. **Clear All** deletes them all, and
**Undo deletion** brings back the last line deleted. Your traces are saved for each board and image
in your own settings folder (`Classic-Repair-Toolbox.traces.json`), never in the board data.

## Recording what you find

With the [Workbooks](Workbooks-tab) feature on, **Add worklog** in the bar above the tabs lets you drag a rectangle on the schematic and write up what you found there. Markers for saved worklogs appear on the image and on the thumbnails - see [Workbooks: daily use](Workbooks-Daily-use).

## Contributor mode

Ticking **Enable contributor mode** in the **Global settings** panel adds a menu that opens when you
**right-click an empty spot of the image**, with two editors:

* **Enable component label editor** - draw or adjust the rectangles that make a component highlight. This is how a wrong or missing highlight gets fixed. Finish with **Apply all editor changes** or **Discard all editor changes**; while the editor is open, switching hardware, board or schematic is locked.
* **Calibrate KiCad traces** - line the KiCad outline up with the board in the image, so the copper overlay lands in the right place. Only offered on images with KiCad data. Finish with **Apply KiCad calibration** or **Discard KiCad calibration**.

Both save into your own local draft of the board - CRT creates the draft if there is none - and not
into the downloaded board files. You then send the change from the [Drafts tab](Drafts-tab); see
[Contribute data via CRT](Contribute-data-via-CRT). The highlights and calibrations are stored the way
[Board JSON](Board-JSON) describes, and both editors are covered step by step in
[Add a new board with KiCad data](Add-new-board-with-KiCad-data).

> 🎬 There is a walkthrough video: [How to use component label editor](https://youtu.be/u-UkD-m4Z6o)

## If it does not work

| What you see | Almost always means |
| --- | --- |
| A component does not highlight | It has no highlight data yet, or the **PAL** / **NTSC** region choice on the left is excluding it |
| Clicking a trace does nothing | That board or image has no [KiCad data](KiCad-folder), **Highlight trace on hover** is off, or **Hold SHIFT to highlight trace on hover** is on and SHIFT is not held |
| Dragging draws a line instead of moving the image | Pan with the **right** mouse button; the left one draws your own traces |
| Traces light up but land in the wrong place | The calibration for that image is stale - recalibrate it |
| Everything lights up as you move the mouse | Turn on **Hold SHIFT to highlight trace on hover** |
| "KiCad data initializing..." | The project is still being parsed in the background. It only happens once per board load. |
