[Wiki Home](Home)

Drive a connected scope from the numpad, without leaving the schematic.

---

CRT can control a network connected oscilloscope from the keyboard/numpad, from within the
component information window - the popup that opens when you click a component. The controls are
_not_ available on the "Oscilloscope" tab itself.

To enable and use the controls, do this:

* Check `Enable network connected oscilloscope tab` on the "Configuration" tab (it is on by default).
* Go to the "Oscilloscope" tab, pick your scope's vendor and series or model, type its address, and
  press **Connect to oscilloscope**.
* Click a component that has an oscilloscope baseline (images depicting a working board), and
  select one of its images.
* In the component information window, check `Numpad controls oscilloscope`. The checkbox is only
  available while CRT is actually connected to the oscilloscope.
* Make sure `NumLock` is on.

The keys work while the component information window has focus. You can use these keys:

<img width="1051" height="358" alt="image" src="https://github.com/user-attachments/assets/8f339e2c-bf05-49bd-ab8d-9cad2a3b018b" />

A few things worth knowing:

* While the controls are on, the numpad digits no longer select pins (e.g. typing `1` will not
  select the image for pin 1), but the non-numpad digits `0`-`9` still do that.
* While the controls are on, `Enter` captures the scope's screen instead of jumping to the first
  image. `Space` still jumps to the first image.
* `Escape` keeps its normal meaning, and will close the window.
* The left/right arrow keys also keep their normal meaning, and will navigate to the previous/next
  image.
* The up/down arrow keys move the trigger level in steps of 0.25 V.
* `+` and `-` step the time base to the next of your scope model's time per division steps, and `1`
  and `2` only work when `1V` and `2V` are among its volts per division steps.
* A capture is saved in the **Image save folder** on the "Oscilloscope" tab - an
  `Oscilloscope images` folder beside CRT's data folder, unless you choose another with
  **Select folder**. The capture is shown in the window as a "User image", saved in
  that folder as `<board label>_<pin>_<region>.png`, and - when the Workbooks feature is on and the
  board has a workbook - an **Attach image to worklog** button lets you file it in a repair - see
  [Workbooks: daily use](Workbooks-Daily-use).

## Capturing a whole baseline pin by pin

The same keys work in **Capture from oscilloscope pin-by-pin** - the button on a component's
"Oscilloscope baseline" page on the "Contribute" tab, shown while the oscilloscope is connected. It
opens a window of its own, which remembers its size and position, with the component's case and every
pin listed. Select the pins you want to measure, or click them on the case - the buttons over the list
select a group of pins and go to the first of them. Probe the pin shown and press `Enter`: the scope's
screen and how it is set are captured, and the pin stays shown so you can check it. A pin that had a
baseline shows the old one on the left and the new capture on the right, each labeled, and any setting
that changed says what it was. Nothing goes into your draft yet: press `Enter` again to retake it, or
**Accept and continue to next pin** (`Right`) to make it that pin's baseline and go on, with the scope
set as the next pin's old baseline was measured. `Left` goes back without un-accepting anything. Click
any picture to see it larger - a click outside it closes it again. Here the numpad always controls the
scope, no checkbox needed, and **Close capturing** or `Escape` closes the window - first asking about
any capture not yet accepted.

**Capture again** on one baseline opens this window on that pin and captures it there, so you can check
it before you accept it.

If your scope is not in the list, or does not work properly, make a model of your own: **My scope
is not listed** on the "Oscilloscope" tab - see
[Synchronize oscilloscope](Synchronize-oscilloscope#my-scope-is-not-listed). Fill in both lists of
steps as well as the commands: the stepping keys find the next value in those lists, and do nothing
when the scope's current value is not in them.
