# Synchronize oscilloscope

Go to [Documentation index](./README.md).

When you select a baseline image in the component information window, _CRT_ can push the
oscilloscope settings stored with that image to your connected oscilloscope - the time base, the
volts per division and the trigger level. The idea is that your own oscilloscope is then set up
exactly like the one that captured the baseline, so what you see on your screen can be compared
directly with the image on screen.

The setup of your oscilloscope is done automatically as you click through the images, so you do not need to dial in the
settings by hand for every pin you probe.

## Enabling it

* Tick `Enable network connected oscilloscope tab` on the "Configuration" tab.
* Go to the "Oscilloscope" tab, fill in the details for your oscilloscope and connect to it.
* Click a component that has an oscilloscope baseline, and select one of its images.
* In the component information window, make sure `Synchronize oscilloscope` is ticked. It is on by
  default, but it is only available while _CRT_ is actually connected to the oscilloscope.

Each time you select an image, the settings for that image are sent to the oscilloscope. The values
being sent are also shown on the image itself, as the `T/DIV`, `V/DIV` and `T:` labels.

## What is synchronized

The values come from the `Component images` sheet in the board's own Excel data file:

| Column  | Meaning       |
| ------- | ------------- |
| `T/DIV` | Time base     |
| `V/DIV` | Volts/div     |
| `T.LVL` | Trigger level |

Only the values that are actually filled in are sent - an image with just a `T/DIV` value will only
have its time base synchronized, and the rest of your oscilloscope settings are left alone. An image
with none of the three, or with no pin number, is ignored entirely.

Nothing else is touched. _CRT_ does not change your channel, probe attenuation, trigger source or
trigger slope - you have to set those up yourself.

## How the values are matched

Values are written as a number followed by a unit, using a period as the decimal separator:

* Time: `ns`, `us` (or `µs`), `ms`, `s` - for example `500ns`, `20us`, `1ms`
* Voltage: `uV` (or `µV`), `mV`, `V` - for example `500mV`, `2V`, `1.5V`

## Good to know

* Selecting the same image again does not resend anything. The commands are only sent when the
  values, the image or the oscilloscope actually change.
* Everything that is sent - and everything that could not be mapped - is written to the output on
  the "Oscilloscope" tab, so that is the place to look when a baseline does not seem to synchronize.
* Synchronizing works well together with
  [controlling the oscilloscope from the keyboard](./controlling-oscilloscope-with-keyboard.md): let the
  image set the starting point, then fine-tune the time base or trigger level from the numpad.

If your oscilloscope is not in the list, or it does not work properly, then please do investigate
which **SCPI commands** work for your specific oscilloscope model, as this varies quite a lot - even
within the same vendor. I do not know all oscilloscopes, nor do I have access to anything other than
my own, so you will need to provide this data yourself. You can add and test the required data in
the main Excel data file `Classic-Repair-Toolbox.xlsx` in the sheet `Oscilloscope`.

Go to [Documentation index](./README.md).