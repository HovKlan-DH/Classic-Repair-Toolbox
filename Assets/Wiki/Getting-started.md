[Wiki Home](Home)

What CRT does, and what to do in your first ten minutes with it.

---

CRT shows you the schematics, component data and known-good oscilloscope readings for a vintage board, so you can work out what is wrong with the one in front of you.

## 1. Pick your board

Two drop-downs, top left: **Hardware** and **Board**.

The board number is printed on the PCB itself — `250407`, `250469` and so on. Pick the one that matches, because component positions and part numbers differ between revisions of the same machine.

The first launch downloads the data for it, which takes a moment.

## 2. Find a component

The list below the drop-downs is every component on the board. Type in the search box to narrow it — `U19`, `VIC`, `capacitor`.

**Click one, and it lights up on the schematic.** Click it again to unhighlight.

Above the list are two filters:

* **Region** — `PAL` / `NTSC`, where a board has region-specific parts
* **Category** — capacitors, ICs, connectors and so on

## 3. Look at the board

The "Schematics" tab is where you will spend your time.

* Mouse wheel zooms, drag pans
* The thumbnails down the side switch between schematic images and PCB photos
* **Click a component on the image** to open its popup — datasheet links, part numbers, local files, and any oscilloscope baseline images

If the board has [KiCad data](KiCad-folder), clicking a pin or a trace highlights the whole copper net it belongs to, across every image.

## 4. The other tabs

| Tab | What it is for |
| --- | --- |
| **[Schematics](Schematics-tab)** | The board, its components and its traces |
| **[Overview](Overview-tab)** | Every component as a list, and a printable bill of materials |
| **[Resources](Resources-tab)** | Datasheets, service manuals and links for this board |
| **[Workbooks](Workbooks-tab)** | Record what you find as you repair |
| **[Oscilloscope](Oscilloscope-tab)** | Drive a network-connected scope |
| **[Contribute](Contribute-tab)** | Send corrections and additions back |
| **[Configuration](Configuration-tab)** | Settings, and the button that opens your data folder |
| **[Feedback](Feedback-tab)** | Send an issue or a question to the developer |
| **[About](About-tab)** | Versions, links, and who contributed this board's data |

**Workbooks** and **Oscilloscope** can both be hidden from the "Configuration" tab, if you do not use them.

## 5. Record what you find

The **Workbooks** tab lets you write down each fault as you find it, mark where it is on the schematic, log time and cost, attach photos, and export the finished repair as a PDF to keep or to share.

That is the one feature worth reading about before you need it — see [Workbooks](Workbooks-tab).

## Where next

* [Workbooks](Workbooks-tab) — track a repair from first fault to finished write-up
* [Synchronize oscilloscope](Synchronize-oscilloscope) — if you have a network-capable scope
* [MiniPro programmer](MiniPro-programmer) — to test a logic IC out of the board
* [Contribute data via CRT](Contribute-data-via-CRT) — when you spot something wrong or missing

## If it does not work

| What you see | Almost always means |
| --- | --- |
| No hardware in the drop-down | The first data download has not finished, or was blocked — check "Configuration" |
| Your board revision is not listed | Nobody has contributed it yet — see [Add a new board](Add-new-board-with-KiCad-data) |
| A component does not highlight | It has no highlight data yet, or the **Region** filter is excluding it |
| Clicking a trace does nothing | That board has no [KiCad data](KiCad-folder) — not every board does |
| No **Workbooks** or **Oscilloscope** tab | Both can be switched off in "Configuration" — check they are still ticked |
