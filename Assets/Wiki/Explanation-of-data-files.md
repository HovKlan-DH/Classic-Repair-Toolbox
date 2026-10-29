[Wiki Home](Home)

What CRT's board data is made of, and where it lives - for the curious. You never need to edit
these files: every change to a board is made in CRT, on the "Contribute" tab.

---

## The data folder

CRT keeps the downloaded data in a `Data v4` folder:

* Windows: `%LocalAppData%\Classic-Repair-Toolbox\Data v4`
* Linux: `~/.local/share/Classic-Repair-Toolbox/Data v4`
* macOS: `~/Library/Application Support/Classic-Repair-Toolbox/Data v4`

**Open data folder** on the "Configuration" tab opens it, and `--data-root` moves it - see
[Command-line parameters](Commandline-parameters). A folder called `Data` beside it is CRT 3's,
which CRT 4 only reads (to take over your CRT 3 drafts and scope models) and never changes.

```
Data v4/
├── Classic-Repair-Toolbox.v4.0.0.json     <- the main file: every board, every oscilloscope
├── Commodore/
│   └── C64/
│       └── 250407/
│           └── board.json                 <- one board, everything about it
├── ZX Spectrum/ ...
├── Files/
│   ├── 82/
│   │   └── 829d81899b3c...ca07.png        <- every file's bytes, stored once
│   └── ...
└── MiniPro/
    └── IC tests/                          <- the IC tests the MiniPro programmer runs
```

## One board, one file

Each board has a folder of its own - `Commodore/C64/250407` - holding just its `board.json`.
That one file is the whole board: its pictures, its components and their pins, pictures, scope
baselines, files and links, its credits and important signals, where each component is marked on
each picture, and how each picture lines up with the KiCad data. See [Board JSON](Board-JSON).

The folder's path is the board's **id**, which is how everything else - the main file, your
drafts, your workbooks - names it.

## Files are a board's own, stored once

Every file a board uses - a picture, a datasheet, a scope screenshot - has a name of the board's
own, such as `Component images/7406.png`. The board's file lists each name with the file's
SHA-256 checksum, and the bytes are stored under that checksum in `Files/`: `Files/82/829d...ca07.png`.

So a file used by eleven boards is downloaded and stored once, while each board still owns its copy:
replacing the 7406 datasheet on one board makes new bytes under a new checksum, and the other ten
boards are untouched. To bring the same change to them, the "Contribute" tab's **Also on other
boards** offers it to each board, as a change of its own.

## The main file

`Classic-Repair-Toolbox.v4.0.0.json` holds two lists:

* **boards** - CRT's hardware and board drop-downs, top to bottom: each board's id, the hardware and
  board names it is listed under, and the hardware notes the "Overview" tab shows
* **oscilloscopes** - every scope model CRT knows: its vendor, series or model, network port, the wait
  between commands, the SCPI command for each thing CRT asks the scope, and the time and volts per
  division steps the numpad keys move through

Your own scope models are in a file of your own beside your settings,
`Classic-Repair-Toolbox.own-oscilloscopes.json`, made and edited with **My scope is not listed** on
the "Oscilloscope" tab - see [Synchronize oscilloscope](Synchronize-oscilloscope#my-scope-is-not-listed).
The data check never touches it.

The version in the main file's name is the first CRT version that reads it. A CRT reads the newest
main file at or below its own version.

## How CRT keeps it up to date

When CRT starts, and **Check for new or updated data at application launch** is ticked on the
"Configuration" tab, it fetches the list of every file in the online data with its checksum
(`dataChecksums.json`), and downloads what is new or changed: the main file and the board files
first, everything else in the background while you work. **Delete orphan and non-used files** removes
files the online data no longer has.

"Online data" is the stable source, or the BETA source when **Download data from the BETA source
instead of the stable source** is ticked - the BETA source carries what maintainers have accepted,
before it is published to everybody.

## Your drafts

A change you make to a board goes into your own **draft** of it, in a `Drafts` folder beside your
settings (`--drafts-root` moves it). A draft is laid out like the data: the board's folder, holding its
own `board.json` with your changes, and a `Files` folder with only the files you added. CRT saves it by
itself, and deletes it once your change is published to the stable source.

## Your workbooks and devices

Your repairs are not board data and never leave your machine - see
[Workbooks: export and your data](Workbooks-Export-and-data).

## CRT 2 and 3

CRT 2 and 3 read the board data from Excel files - a main workbook and one workbook per board, with a
JSON file beside each for the coordinates. That data is frozen as it was when CRT 4 came out: CRT 2 and
3 keep showing it, and are told to update to contribute.
