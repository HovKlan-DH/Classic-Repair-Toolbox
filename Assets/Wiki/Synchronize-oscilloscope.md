[Wiki Home](Home)

Set your scope up the same way the baseline was taken.

---

When you select a baseline image in the component information window, CRT can push the
oscilloscope settings stored with that image to your connected oscilloscope - the time base, the
volts per division and the trigger level. The idea is that your own oscilloscope is then set up
exactly like the one that captured the baseline, so what you see on your screen can be compared
directly with the image on screen.

The setup of your oscilloscope is done automatically as you click through the images, so you do not need to dial in the
settings by hand for every pin you probe.

## Enabling it

* Tick `Enable network connected oscilloscope tab` on the "Configuration" tab (it is on by default).
* Go to the "Oscilloscope" tab, pick your scope's vendor and series or model, type its address, and
  press **Connect to oscilloscope**. Your scope is not in the list? See "My scope is not listed" at
  the end of this page.
* Click a component that has an oscilloscope baseline, and select one of its images.
* In the component information window, make sure `Synchronize oscilloscope` is ticked. It is on by
  default, but it is only available while CRT is actually connected to the oscilloscope.

Each time you select an image, the settings for that image are sent to the oscilloscope. The values
being sent are also shown on the image itself, as the `T/DIV`, `V/DIV` and `T:` labels.

## What is synchronized

The values are the baseline's own, as the board's data holds them - the fields of a scope baseline
in the component's pane on the "Contribute" tab:

| Field           | Meaning       |
| --------------- | ------------- |
| `T/DIV`         | Time base     |
| `V/DIV`         | Volts/div     |
| `Trigger level` | Trigger level |

Only the values that are actually filled in are sent - an image with just a `T/DIV` value will only
have its time base synchronized, and the rest of your oscilloscope settings are left alone. An image
with none of the three, or with no pin number, is ignored entirely.

Nothing else is touched. CRT does not change your channel, probe attenuation, trigger source or
trigger slope - you have to set those up yourself.

## How the values are matched

Values are written as a number followed by a unit, using a period as the decimal separator. The
unit is required, and its capitals do not matter:

* Time: `ns`, `us` (or `µs`), `ms`, `s` - for example `500ns`, `20us`, `1ms`
* Voltage: `uV` (or `µV`), `mV`, `V` - for example `500mV`, `2V`, `1.5V`

A `T/DIV` or `V/DIV` value is only sent when it matches one of the values your scope supports - the
time and volts per division steps of your scope's model (for a model of your own, under **Edit my
model** on the "Oscilloscope" tab). A value that is
not in the list is skipped, and the output on the "Oscilloscope" tab says
`Could not map image T/DIV value [...] to a supported oscilloscope value` (or the same for
VOLTS/DIV). The trigger level is sent as written.

## Good to know

* Selecting the same image again does not resend anything. The commands are only sent when the
  values, the image or the oscilloscope actually change.
* Everything that is sent - and everything that could not be mapped - is written to the output on
  the "Oscilloscope" tab, so that is the place to look when a baseline does not seem to synchronize.
* Synchronizing works well together with
  [controlling the oscilloscope from the keyboard](Controlling-oscilloscope-with-keyboard): let the
  image set the starting point, then fine-tune the time base or trigger level from the numpad.

## My scope is not listed

The **SCPI commands** a scope takes vary a lot - even between models of one vendor - so CRT knows only
the models people have added. To add yours:

1. On the "Oscilloscope" tab, choose the model most like yours (often one of the same vendor) and
   press **My scope is not listed**.
2. Give your model its vendor and name, and press **Create model**. It is a copy of the one you
   started from, kept in a file of your own that the data check never puts back.
3. Connect to your scope, then go through the commands with your scope's programming manual open.
   **Try it** beside each command sends it and shows the scope's answer under it; a box is saved as
   you leave it.
4. **Run full test suite** tries every command in turn.
5. **Submit this model** sends it to the administrator, who publishes it for everybody. You are mailed
   either way.

Your scope needs a network (LAN) port that takes SCPI commands - not every scope has one. A port
that only works with the scope maker's own PC software, or a scope with only USB, cannot be used.
