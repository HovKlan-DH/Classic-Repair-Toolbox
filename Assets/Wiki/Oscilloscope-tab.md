[Wiki Home](Home)

Connect CRT to a network-capable oscilloscope.

---

Hidden unless **Enable network connected oscilloscope tab** is ticked in
[Configuration](Configuration-tab).

## Connecting

* **IP address or FQDN** and **TCP port** — where your scope is on the network
* **Series or model** — pick the entry matching your scope, which is what tells CRT which SCPI
  commands to send
* **Auto-connect oscilloscope** — reconnect automatically at startup

**Run full test suite** walks through every command for your scope and reports what worked. It is
the fastest way to find out whether a model entry actually fits your scope.

The output pane logs every command sent and every reply, which is where to look when something does
not behave.

## Image save folder

Where screen captures pulled from the scope are written. They can also be filed straight into a
repair — see [Workbooks: daily use](Workbooks-Daily-use).

## What you can do with it

* **[Synchronize oscilloscope](Synchronize-oscilloscope)** — set your scope up exactly like the one
  that captured a baseline image, automatically as you click through pins
* **[Controlling oscilloscope with keyboard](Controlling-oscilloscope-with-keyboard)** — drive the
  time base, volts and trigger from the numpad without leaving the schematic

## My scope is not listed

The SCPI commands vary a lot between vendors, and even between models from one vendor. You can add
your own model to the `Oscilloscope` sheet in the main Excel data file and test it with **Run full
test suite** — see [Main Excel](Main-Excel) for the columns.
