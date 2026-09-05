[Wiki Home](Home) · [Data files](Explanation-of-data-files)

Oscilloscope screenshots from a known working board.

---

A folder named `Scope baseline` inside a board folder, holding oscilloscope screenshots from a known working board. You use them to compare against your own measurements when hunting a fault.

```
Data/Commodore/C64/250407/
└── Scope baseline/
    ├── C64 250407 Introduction to scope baseline measurements.txt
    ├── U19_1_PAL.png
    ├── U19_1_NTSC.png
    └── ...
```

The images are named `<component>_<pin>_<region>.png`, so `U19_1_PAL.png` is pin 1 of U19 on a PAL board.

The text file describes the exact system the measurements were taken on - board revision, PSU, region, oscilloscope model and what the machine was doing at the time. Worth reading before you conclude your own reading is wrong.

## Getting them into the application

The images are not picked up from the folder on their own. Each one is a row in the board's Excel data file, under the `Component images` sheet, and it shows up in the component popup when the row has:

* a `Pin` number, and
* at least one of `T/DIV`, `V/DIV` or `T.LVL`

Without those the image is treated as an ordinary component image instead.

See [Board Excel](Board-Excel) for the columns.

## Using them

Click a component in the list, and the popup shows the baseline image for each pin. If your oscilloscope is connected, CRT can also set it up with the same T/DIV and V/DIV values used for the baseline - see [Synchronize oscilloscope](Synchronize-oscilloscope).

Do note that a matching reading is not a guarantee the chip is fine - it may simply not be active in the mode the machine is in.
