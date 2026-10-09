[Wiki Home](Home) · [Data files](Explanation-of-data-files)

The data file for one board: its schematics, components, images, files and links.

---

A board Excel file is placed in the folder for the individual board, for example `Data C64 250407 v2.0.0.xlsx`. The version number in its name matches the [Main Excel](Main-Excel) file that lists the board. A new board made with **Add a new board** starts out as `Data <Hardware> <Board>.xlsx`, and gets the version in its name when it is published.

The board Excel file is by far the most time consuming part, when building a new board from ground, as the data gathering part is hard/slow, if you want a good data quality (and yes, please... we want that).

You can fill it in from inside CRT - the Contribute tab, the label editor on the Schematics tab, or "Edit in table format" on the [Drafts tab](Drafts-tab) - or open your draft's file in Excel. Below is the documentation for each worksheet and the columns inside it.

## The shape of every worksheet

A workbook written by CRT (a draft, or a board published from a submission) has nine worksheets, in this order: `Board schematics`, `Components`, `Component images`, `Component local files`, `Component links`, `Board local files`, `Board links`, `Important signals` and `Credits`. Every worksheet opens the same way:

| Row | Holds |
| --- | --- |
| 1 | `# Hardware: Commodore 64` |
| 2 | `# Board: 250407` |
| 3 | `# Revision date: 2026-August-21` |
| 4 | (blank) |
| 5 | The title band, e.g. `Components` |
| 6 | The column headers - everything above the data stays in view while you scroll |

The data starts on row 7.

CRT does not need this exact shape to read a file: it finds the header row by its column names, wherever it is. The `# Hardware:` and `# Board:` lines are a caption for a person reading the file, and nothing in CRT depends on them. A board published without them gets the names it is listed under in the drop-downs.

### Required headers

A worksheet is only read when its header row has all of these names. Its other columns are read when they are there, and blank when they are not.

| Worksheet | Required headers |
| --- | --- |
| Board schematics | `Schematic name`, `Schematic image file` |
| Components | `Board label`, `Friendly name`, `Technical name or value` |
| Component images | `Board label`, `Pin`, `Name`, `File` |
| Component local files | `Board label`, `Name`, `File` |
| Component links | `Board label`, `Name`, `URL` |
| Board local files | `Category`, `Name`, `File` |
| Board links | `Category`, `Name`, `URL` |
| Important signals | `Display name`, `KiCad net name` |
| Credits | `Category`, `Name or handle` |

A workbook from an older version may still have a `UUID v4` column. It is ignored, and dropped the next time CRT saves the file.

### The revision date

The `# Revision date:` line says when the board's data was last published - the year, the full English month name and the day, with no leading zero on the day: `2026-August-21`, `2026-October-4`.

**You do not set it.** The server writes the publish date into every worksheet whenever a submission is published, so a date typed by hand does not survive.

CRT reads it from the first cell starting `# Revision date:` within the first 10 rows and 10 columns of `Board schematics`. It is shown on the Contribute tab and the About tab ("Board Excel data last revisioned").

One thing reads it as a date: the **drift warning**. When you have a draft of a board and the official data is updated underneath it, CRT compares this value against the one your draft was started from, and tells you the official data has been *updated* rather than merely *changed*. If the value cannot be read as a date, nothing breaks - CRT simply says the data "has changed" instead of "has been updated", which is all it can honestly claim.

### Files

Every column that names a file holds its path **relative to the `Data` folder**, with `/` between the folders and exactly the capital letters the file has on disk - for example `Commodore/C64/250407/Board Layout 250407 NTSC.png`.

* **Where:** in the board's own folder, in `<Manufacturer>/Shared files/` (files that manufacturer's boards share) or in `Generic shared files/` (files every board shares). A file in another board's folder can be used too, but only exactly as it is.
* **Which kinds:** pictures (`.png`, `.jpg`, `.jpeg`, `.gif`, `.bmp`, `.webp`), PDF documents (`.pdf`), plain text (`.txt`) and web pages (`.html`, `.htm`). Anything else is refused when you submit.
* A file that is not there, a different capital letter or a `\` in the path is an error in the draft's table, and the draft cannot be submitted until it is fixed.

### Links

Every `URL` must start with `http://` or `https://`. CRT opens nothing else, and anything else is an error in the draft's table.


## Worksheet: Board schematics

Depicts which schematic images are available for the board, in the order they are listed.

### Column: Schematic name

Exact same name as shown in the thumbnail label.\
Keep the name short, so it can fit in the label.\
A schematic name must be unique in the board - capital letters do not count, and a repeat is an error.

### Column: Schematic image file

Path and filename to the schematic image file - a picture, see [Files](#files) above.\
Use **relative** path from the `Data` folder.

Ideally this has a fairly high quality, but of course the visibility/performance ratio must be balanced, as the larger image, the harder the display of it gets.

### Columns: Schematic highlight color and Thumbnail highlight color

Which color to use for component highlighting - `Schematic highlight color` in the "Main" image, `Thumbnail highlight color` in the thumbnails.\
Any [standard colour name](https://reference.avaloniaui.net/api/Avalonia.Media/Colors/) works, e.g. `Red`, `IndianRed`, `CornflowerBlue`, and so does a hex value such as `#FF0000`. Blank means `IndianRed`.\
On a view backed by KiCad data, `Schematic highlight color` is also the colour of the component's copper traces.

### Columns: Schematic highlight opacity and Thumbnail highlight opacity

Where relevant then use a semi-transparent highlight, to allow viewing of potential information below the component highlight.\
Write a number from `0` (fully transparent) to `1` (solid), with a dot - the C64 `250407` board uses `0.3` for the schematic and `1` for the thumbnails. `30` and `30%` work too. Blank means `0.2`.

### Column: Opposite trace highlight color

On a view backed by KiCad data: the colour of the copper on the other side of the board, drawn when "Show traces from opposite side" is ticked on the Schematics tab. Blank means `DodgerBlue`.

### Column: CAD name

The KiCad view this image shows, copied exactly from "CAD names you can use" in the Drafts tab's **KiCad data** window - for example `250407_ - PCB Top`. The logfile lists them too, as `Display name [...]` lines.\
Leave it blank for an image with no KiCad data. See [How to add a new board with KiCad data](Add-new-board-with-KiCad-data).

## Worksheet: Components

The rows are shown in this order in the component list and on the Overview tab.

### Column: Board label

Very short label representing the component name.\
Ideally it should be 2-5 characters long only.\
A board label must be unique within a region: it can only be on several rows when each row has its own `Region` (blank counts as one region). A repeat is an error.

Do note that `Board label` + `Friendly name` + `Technical name or value` is concatenated in the component list, so do consider to make this as short and precise as possible.

### Column: Friendly name

Typically components have "human readable" or "friendly" names.\
Could also be that component is most often referred to as this name.\
Should still be as short as possible.

Do note that `Board label` + `Friendly name` + `Technical name or value` is concatenated in the component list, so do consider to make this as short and precise as possible.

### Column: Technical name or value

The value of the component or its technical name, depending on its nature.

### Column: Part-number

Typically the part-number from the vendor. In many cases there exists lists of part-numbers, so it can be a good reference, as often you can directly lookup technical details for a part-number.\
The same part-number on two components with different technical names gets a warning, as it has usually been copied from the wrong row.

### Column: Category

Could be `Capacitor`, `Resistor`, `IC`, `Connector`, `Misc` or whatever else suits as a group identified for the component.
Keep the list short, so there is not many categories, but also do make sure to group it where it makes sense.

### Column: Region

Should be either empty (blank), `PAL` or `NTSC`.\
Use only a specific region (PAL or NTSC) when this component is specific for this region only.\
Use empty (blank) when the component is generic, and for no specific region.

### Column: Short one-liner description (one short line only!)

That is the full header - in a file written by CRT, the part in brackets is on a second line of the same cell.\
Will be shown in the component information popup.\
Is a short contextual and relevant information about the component.\
Could be technical information, which could not fit in "Technical name or value".\
Must be **one line only**!

## Worksheet: Component images

A list of images shown in the component information popup.

### Column: Board label

Direct reference from the `Components` worksheet.\
A board label can be referenced many times, as it can have multiple files per component.

### Column: Region

This should be either empty (blank), `PAL` or `NTSC` to determine which region is relevant for this image.\
E.g. doing oscilloscope measurements would be nice to know if this is done on a `PAL` or `NTSC` system.\
Also for e.g. the pinout image - does this show a `PAL` or `NTSC` component, as this could differ.\
If the region is not relevant, then leave it blank.

### Column: Pin

Numeric/integer value.\
If the image is for a specific component pin.\
If the pin is not relevant, then leave it blank.

A row with a `Pin` and at least one of `T/DIV`, `V/DIV` or `T.LVL` is an oscilloscope baseline - see [Scope baseline folder](Scope-baseline-folder).

### Column: Name

A pinout image should ideally show the legs and what is their input/output.\
If an image is for a specific pin, then document its name for easy reference.

### Column: Expected oscilloscope reading

What value is expected here when measuring this with an oscilloscope?\
This can be different things like `LOW`, `HIGH`, `Pulsing`, a frequency or voltage.

### Column: T/DIV

`T/DIV` is the "time per division".\
E.g. `1uS` mean each horizontal division is "1 micro second".\
CRT knows the 1-2-5 steps from `2nS` to `1000S`: `2nS`, `5nS`, `10nS` ... `500nS`, `1uS` ... `500uS`, `1mS` ... `500mS`, `1S` ... `500S` and `1000S` (capital letters do not matter). Anything else gets a warning.

### Column: V/DIV

`V/DIV` is the "volts per division".\
E.g. `1V` mean each vertical division is "1V".\
CRT knows `5mV`, `10mV`, `20mV`, `50mV`, `100mV`, `200mV`, `500mV`, `1V`, `2V`, `5V`, `10V`, `20V`, `50V` and `100V`. Anything else gets a warning.

### Column: T.LVL

`T.LVL` is the "trigger level in volts".\
E.g. `1.5V` mean that the scope will trigger at "1.5V".\
Write a number with a `V` after it, using a dot for decimals - `1.4V`, `-2V`. Anything else gets a warning.

### Column: File

Path and filename to the image file - a picture, see [Files](#files) above.\
Use **relative** path from the `Data` folder.\
It may be left blank when the row has a `Note` - e.g. a `Pinout` row that only gives a compatible part number. A row with neither a file nor a note is an error.

### Column: Note

The note field for the image.\
Typically (always?) the first image is the `Pinout` image, and this is special as its `Note` field is used for the component text in the "Overview" tab.

## Worksheet: Component local files

Component local files will show in both the "Overview" tab and the component information popup in CRT.\
It is a local file specifically for this component - e.g. a datasheet or technical documentation.

### Column: Board label

Direct reference from the `Components` worksheet.\
You can have multiple local files per component, so the board label is allowed to duplicate.

### Column: Name

Name for the file that will be shown in CRT.

### Column: File

Path and filename to the local file - see [Files](#files) above for the kinds of file allowed.\
The local file will be opened in whatever application you use for that kind of file.\
Use **relative** path from the `Data` folder.

## Worksheet: Component links

Component URLs will show in both the "Overview" tab and the component information popup in CRT.\
It is a URL specifically for this component - e.g. a technical documentation or troubleshooting references.

### Column: Board label

Direct reference from the `Components` worksheet.\
You can have multiple links per component, so the board label is allowed to duplicate.

### Column: Name

Name for the link that will be shown in CRT.

### Column: URL

The URL will be opened in your default browser. It must start with `http://` or `https://`.

## Worksheet: Board local files

Board local files will show in the "Resources" tab in CRT.\
It is meant as a general documentation for the board - e.g. generic diagnosing or troubleshooting.

### Column: Category

What kind of file is this - some examples are `Troubleshooting`, `Technical documentation` or alike.\
You can have multiple local files per category, so the category name is allowed to duplicate.

### Column: Name

Name for the file that will be shown in CRT.

### Column: File

Path and filename to the local file - see [Files](#files) above for the kinds of file allowed.\
The local file will be opened in whatever application you use for that kind of file.\
Use **relative** path from the `Data` folder.

## Worksheet: Board links

Board URLs will show in the "Resources" tab in CRT.\
It is meant as a general documentation for the board - e.g. generic diagnosing or troubleshooting.

### Column: Category

What kind of URL is this - some examples are `Troubleshooting`, `Technical documentation` or alike.\
You can have multiple URLs per category, so the category name is allowed to duplicate.

### Column: Name

Name for the link that will be shown in CRT.

### Column: URL

The URL will be opened in your default browser. It must start with `http://` or `https://`.

## Worksheet: Important signals

Will show the important (KiCad) signals to have in the list in the schematics image.\
It does require KiCad data.

### Column: Display name

The name to display in the CRT list in the schematics image.\
Many times the KiCad data is weird to look at, so this is the "human readable" name for it.\
A display name can be repeated many times, as it then will show all the KiCad net names belonging to it.

### Column: KiCad net name

The net name as the KiCad data has it. The full name and its last part both work (`/CPU/PHI2` or `PHI2`), and capital letters do not matter.\
With "Enable contributor mode" ticked on the Schematics tab, the logfile lists every net name of the board.\
A row missing either column is left out when the file is saved, and the draft's table warns about it.

## Worksheet: Credits

Will show who has contributed with data to this board.\
Shown in the "About" tab.

### Column: Category

What was contributed - e.g. `Schematic image`, `Component pinouts` or `Board labelling`.

### Column: Sub-category

Optional detail - e.g. which schematic a `Schematic image` credit is for.

### Column: Name or handle

The name or handle to credit.

### Column: Contact (email or web page)

Optional. An email address or a web address, which becomes clickable in the "About" tab.

## Rules for every Excel file

See [Data files](Explanation-of-data-files#rules-for-every-excel-file) - the same rules apply to every Excel file.
