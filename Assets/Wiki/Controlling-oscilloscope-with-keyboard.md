[Wiki Home](Home)

Drive a connected scope from the numpad, without leaving the schematic.

---

CRT can control a network connected oscilloscope from the keyboard/numpad, from within the
component information window - the popup that opens when you click a component. The controls are
_not_ available on the "Oscilloscope" tab itself.

To enable and use the controls, do this:

* Tick `Enable network connected oscilloscope tab` on the "Configuration" tab (it is on by default).
* Go to the "Oscilloscope" tab, fill in the details for your oscilloscope and connect to it - see
  [Oscilloscope tab](Oscilloscope-tab).
* Click a component that has an oscilloscope baseline (images depicting a working board), and
  select one of its images.
* In the component information window, tick `Numpad controls oscilloscope`. The checkbox is only
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
* `+` and `-` step the time base to the next value in your scope's `TIME/DIV` list, and `1` and `2`
  only work when `1V` and `2V` are in its `VOLTS/DIV` list.
* Capturing an image requires that you have chosen an **Image save folder** with **Select folder**
  on the "Oscilloscope" tab first. The capture is shown in the window as a "User image", saved in
  that folder as `<board label>_<pin>_<region>.png`, and - when the Workbooks feature is on and the
  board has a workbook - an **Attach image to worklog** button lets you file it in a repair - see
  [Workbooks: daily use](Workbooks-Daily-use).

If your oscilloscope is not in the list, or it does not work properly, then please do investigate
which **SCPI commands** work for your specific oscilloscope model, as this varies quite a lot - even
within the same vendor. I do not know all oscilloscopes, nor do I have access to anything other than
my own, so you will need to provide this data yourself. You can add and test the required data in
the main Excel data file `Classic-Repair-Toolbox.v2.0.0.xlsx` in the sheet `Oscilloscope` - see
[Main Excel](Main-Excel). Note that the
`TIME/DIV` and `VOLTS/DIV` value lists must be filled in as well as the command columns - the
stepping keys resolve their target from those lists, and do nothing if the current value is not
found there. Also note that the launch-time data check replaces a changed data file with the
server's copy at the next launch, unless **Check for new or updated data at application launch** is
unticked in [Configuration](Configuration-tab).
