[Wiki Home](Home) · [Data files](Explanation-of-data-files)

One board's `board.json`: everything CRT knows about the board, in one file. For reading - every
change is made in CRT, on the "Contribute" tab, which writes this file for you.

---

```
Data v4/Commodore/C64/250407/board.json
```

The file opens with the board's names and the date of its last change, then one list per kind of
data. Each list has one entry per line, so two versions of a board compare line by line:

```json
{
  "formatVersion": 1,
  "hardwareName": "Commodore 64",
  "boardName": "250407",
  "revisionDate": "2026-October-11",
  "schematics": [
    {"name":"Board layout","imageFile":"Board Layout 250407 NTSC.png","highlightColor":"Red","highlightOpacity":"30%", ...},
    ...
  ],
  "components": [
    {"boardLabel":"C1","friendlyName":"Ceramic","technicalNameOrValue":"100nF 25V","partNumber":"251075-06","category":"Capacitor","region":"","description":""},
    ...
  ],
  ...
  "files": [
    {"name":"Board Layout 250407 NTSC.png","sha256":"829d8189...ca07","size":1006044},
    ...
  ]
}
```

## The lists

| List | One entry per | What it holds |
| --- | --- | --- |
| `schematics` | Picture of the board | Its name, its image file, the colours and opacities components are marked in (on the picture, on its thumbnail, and on the other side's traces), and the KiCad view it shows (`cadName`) |
| `components` | Component | Board label (`U8`), friendly name, technical name or value, part number, category, region (`PAL`, `NTSC` or empty for both), short description - and, from CRT 4, its package and repair notes |
| `componentPins` | Pin of a component | Pin number, its name (`VCC`, `~RESET`), its direction (Input, Output, Power, Ground ...) |
| `componentReplacements` | Replacement part | Its part number, whether it is a drop-in or needs changes, what to change, and where that is known from |
| `componentImages` | Picture of a component | A photo, a pinout, a render - with its name and a note |
| `scopeBaselines` | Oscilloscope reading of one pin | The pin, the reading expected, the screenshot, the scope settings it was taken with (`timeDiv`, `voltsDiv`, `triggerLevel`), what the machine was doing and where the reading comes from |
| `componentFiles`, `componentLinks` | File or web page about a component | A datasheet, an article |
| `boardFiles`, `boardLinks` | File or web page about the board | Service manuals, schematics, forum threads - each with a category |
| `credits` | Person thanked | What they did, their name or handle, and how to reach them |
| `importantSignals` | Signal the "Schematics" tab's panel lists | Its name (`+5VDC`) and the KiCad net it is (`+5V`) |
| `highlights` | Rectangle marking a component on a picture | The picture, the board label, and `x`, `y`, `width`, `height` in the picture's own pixels |
| `kiCadCalibrations` | Picture lined up with a KiCad view | Offset, scale and mirroring that put the KiCad copper on top of the picture |
| `schematicRenames` | Picture renamed by a publish | Old name, new name, its file, the date - so your own worklogs and settings follow it |
| `files` | File of the board | The board's name for it, its SHA-256 and size |

A component on a board with regional variants can have one entry per region; everything that
belongs to one component - its pins, pictures, baselines - names it by its board label and region.

## Files are named by the board

Every `file` or `imageFile` holds a name from the board's own `files` list - `Component
images/7406.png`, `Scope baseline/U1_1_PAL.png` - never a path into the data folder. The `files`
entry gives that name its checksum, and the bytes are at `Files/<first two letters>/<sha256>.<ext>`
in the data folder (see [Data files](Explanation-of-data-files)).

## What a later CRT adds is kept

Every entry keeps the fields this CRT does not know, and writes them back when the board is
changed - so data a later CRT adds survives a change made with an older one. A board file of a
higher `formatVersion` than CRT knows is not read at all: that CRT has to be updated.

## Where it came from

Before CRT 4, a board was an Excel workbook with a JSON file beside it for the coordinates. Every
published board was converted into this file without losing anything, and your own CRT 3 drafts are
converted the first time CRT 4 starts - the old draft is kept inside it, in a `.crt-pre-4.0` folder.
