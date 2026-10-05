[Wiki Home](Home)

Connect CRT to a network-capable oscilloscope.

---

Shown while **Enable network connected oscilloscope tab** is ticked in
[Configuration](Configuration-tab), which it is by default. Unticking it hides the tab and drops
the connection.

## Connecting

* **Vendor** - the make of your scope; it narrows the list below
* **Series or model** - pick the entry matching your scope, which is what tells CRT which SCPI
  commands to send
* **IP address or FQDN** and **TCP port** - where your scope is on the network. The port starts
  at `5025`.
* **Auto-connect oscilloscope** - while ticked, CRT connects by itself: at startup, and again
  whenever the connection is lost, retrying until the scope answers
* **Connect to oscilloscope** - connects now

Once a connection has been made, or while **Auto-connect oscilloscope** is ticked, the window title
says `(oscilloscope connected)` or `(oscilloscope disconnected)`, so you can see the state from any
tab.

**Run full test suite** - available once connected - runs every command your scope's entry defines
in turn: it stops and starts acquisition, sets the trigger level, time base and volts per division,
and pulls a screen image. Each command and reply is written to the output, which makes it the
fastest way to find out whether a model entry actually fits your scope. It changes the scope's
settings while it runs.

The output pane on the right logs the commands sent and the replies, which is where to look when
something does not behave. Keyboard steps from the
[numpad controls](Controlling-oscilloscope-with-keyboard) write one summary line per step instead.

## Image save folder

Where screen captures pulled from the scope are written. Choose it with **Select folder**; **Open
folder** shows it in your file manager.

A capture is taken from the component information window with the `Enter` key - see
[Controlling oscilloscope with keyboard](Controlling-oscilloscope-with-keyboard). It is saved as a
PNG named `<board label>_<pin>_<region>.png` (the image's name instead of the pin when it has none,
and no region part when there is none) - the same naming as the
[Scope baseline folder](Scope-baseline-folder). A new capture of the same pin replaces the old file.

Captures can also be filed straight into a repair - see [Workbooks: daily use](Workbooks-Daily-use).

## What you can do with it

* **[Synchronize oscilloscope](Synchronize-oscilloscope)** - set your scope up exactly like the one
  that captured a baseline image, automatically as you click through pins
* **[Controlling oscilloscope with keyboard](Controlling-oscilloscope-with-keyboard)** - drive the
  time base, volts and trigger from the numpad without leaving the schematic

## My scope is not listed

The SCPI commands vary a lot between vendors, and even between models from one vendor. You can add
your own model to the `Oscilloscope` sheet in the main Excel data file,
`Classic-Repair-Toolbox.v2.0.0.xlsx` in the data folder, and test it with **Run full test suite** -
see [Main Excel](Main-Excel) for the columns. (CRT reads the newest
`Classic-Repair-Toolbox.v<version>.xlsx` at or below its own version.)

> [!IMPORTANT]
> The launch-time data check replaces any data file that differs from the server's copy, so your
> edited file is put back at the next launch. Untick **Check for new or updated data at application
> launch** in [Configuration](Configuration-tab) while you test, and send the working model to the
> developer through the [Feedback tab](Feedback-tab) so it can be added for everybody.
