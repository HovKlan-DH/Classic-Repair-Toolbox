Go to [Wiki Home](https://github.com/HovKlan-DH/Classic-Repair-Toolbox/wiki).

_CRT_ can control a network connected oscilloscope from the keyboard/numpad, from within the
component information window - the popup that opens when you click a component. The controls are
_not_ available on the "Oscilloscope" tab itself.

To enable and use the controls, do this:

* Tick `Enable network connected oscilloscope tab` on the "Configuration" tab.
* Go to the "Oscilloscope" tab, fill in the details for your oscilloscope and connect to it.
* Click a component that has an oscilloscope baseline (images depicting a working system), and
  select one of its images.
* In the component information window, tick `Numpad controls oscilloscope`. The checkbox is only
  available while _CRT_ is actually connected to the oscilloscope.
* Make sure `NumLock` is on.

The keys work while the component information window has focus. You can use these keys:

<img width="1051" height="358" alt="image" src="https://github.com/user-attachments/assets/8f339e2c-bf05-49bd-ab8d-9cad2a3b018b" />

A few things worth knowing:

* While the controls are on, the numpad digits no longer select pins (e.g. typing `1` will not
  select the image for pin 1), but the non-numpad digits `0`-`9` still do that.
* `Escape` keeps its normal meaning, and will close the window.
* The left/right arrow keys also keep their normal meaning, and will navigate to the previous/next
  image.
* Capturing an image requires that you have selected an oscilloscope image folder on the
  "Oscilloscope" tab first.

If your oscilloscope is not in the list, or it does not work properly, then please do investigate
which **SCPI commands** work for your specific oscilloscope model, as this varies quite a lot - even
within the same vendor. I do not know all oscilloscopes, nor do I have access to anything other than
my own, so you will need to provide this data yourself. You can add and test the required data in
the main Excel data file `Classic-Repair-Toolbox.xlsx` in the sheet `Oscilloscope`. Note that the
`TIME/DIV` and `VOLTS/DIV` value lists must be filled in as well as the command columns - the
stepping keys resolve their target from those lists, and do nothing if the current value is not
found there.

Go to [Wiki Home](https://github.com/HovKlan-DH/Classic-Repair-Toolbox/wiki).