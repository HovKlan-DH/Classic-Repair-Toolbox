[Wiki Home](Home)

What CRT does, and what to do in your first ten minutes with it.

---

CRT shows you the schematics, component data and known-good oscilloscope readings for a vintage board, so you can work out what is wrong with the one in front of you.

## 1. Install it

Download the installer for your system from the [releases page](https://github.com/HovKlan-DH/Classic-Repair-Toolbox/releases) and run it. CRT updates itself after that - it tells you when a new version is ready. On Linux, see [Installing application in Linux](Installing-application-in-Linux).

The first launch sets up the data for every board - copied from the installer, then checked against the online source for anything newer - which takes a moment. Newer pictures keep downloading in the background after CRT has opened.

## 2. Pick your board

Two drop-downs, top left: **Hardware** and **Board**.

The board number is printed on the PCB itself - `250407`, `250469` and so on. Pick the one that matches, because component positions and part numbers differ between revisions of the same machine.

## 3. Find a component

The list below the drop-downs is every component on the board. Type in the search box under it to narrow it - `U19`, `VIC`, `capacitor`.

**Click one, and it lights up on the picture.** Click it again to unhighlight.

Two filters narrow the list:

* **Category** - above the list: capacitors, ICs, connectors and so on
* **Region** - the `PAL` / `NTSC` buttons below the list, where a board has region-specific parts

## 4. Look at the board

The "Schematics" tab is where you will spend your time.

* Mouse wheel (or a trackpad pinch) zooms
* Drag with the **right** or **middle** mouse button (or two fingers on a trackpad) to move the picture
* The thumbnails down the side switch between schematics and photos of the board
* **Click a component on the picture** to open its popup - what it is, its pins, datasheet links, replacement parts, local files, and any oscilloscope baseline readings
* **Right-click a component** to highlight it, or to correct it with **Edit this component**
* **Draw trace** lets you draw lines of your own over the picture, while you follow a signal

If the board has [KiCad data](KiCad-folder), clicking a pin or a trace lights up the whole copper net it belongs to, across every picture.

## 5. The other tabs

| Tab | What it is for |
| --- | --- |
| **Schematics** | The board, its components and its traces |
| **Overview** | Every component as a list, and a printable bill of materials |
| **Resources** | Datasheets, service manuals and links for this board |
| **[Workbooks](Workbooks-tab)** | Record a repair from first fault to finished write-up - and keep your own machines, each with its repairs |
| **Oscilloscope** | Drive a network-connected scope - see [Synchronize oscilloscope](Synchronize-oscilloscope) |
| **Contribute** | Correct or add board data, or add a whole new board, and send it in |
| **[Maintainer](Maintainer-tab)** | For maintainers only: review and publish what others send in |
| **Configuration** | Settings, and the buttons that open your data, workbooks and settings folders |
| **Feedback** | Send an issue or a question to the developer |
| **About** | Versions, links, and who contributed this board's data |

**Workbooks** and **Oscilloscope** can both be hidden from the "Configuration" tab if you do not use them. **Maintainer** shows only when "Enable Maintainer tab" is ticked there - it needs a maintainer account.

Each tab says what it is for as you use it. Where a setting has more to it, **Learn more** beside it opens the page about it here.

## 6. Record what you find

The **Workbooks** tab lets you write down each fault as you find it, mark where it is on the picture, log time and cost, attach photos, and export the finished repair as a PDF or a ZIP to keep or to share.

That is the one feature worth reading about before you need it - see [Workbooks](Workbooks-tab).

## 7. Something wrong or missing?

The **Contribute** tab is always there. Change anything on the board you are looking at - CRT keeps your changes in a draft of your own, saved as you go - and **Submit** it when you are done. The board's maintainers look at it, and once it is published everybody gets it. No account is needed.

## If it does not work

| What you see | Almost always means |
| --- | --- |
| No hardware in the drop-down | The first data download has not finished, or was blocked - check "Configuration" |
| Your board revision is not listed | Nobody has added it yet - **Add a new board** on the "Contribute" tab |
| A component does not highlight | It is not marked on that picture yet, or the **Region** filter is excluding it |
| Clicking a trace does nothing | That board has no [KiCad data](KiCad-folder) - not every board does - or **Traces visible** is off |
| No **Workbooks** or **Oscilloscope** tab | Both can be switched off in "Configuration" - check they are still ticked |
| No **Maintainer** tab | It shows only with "Enable Maintainer tab" ticked in "Configuration" |
