Go to [Wiki Home](https://github.com/HovKlan-DH/Classic-Repair-Toolbox/wiki).\
Go to [Explanation data files](https://github.com/HovKlan-DH/Classic-Repair-Toolbox/wiki/Explanation-of-data-files).

A board Excel file is placed in the folder for the individual board (system). As an example it could `Data C64 250407.xlsx` or `Data C64 250407 v2.0.0.xlsx`. The Excel file, alongside the main Excel file, is versionized meaning the version number represents the file will work from _FROM this version and onwards_.

The board Excel file is by far the most time consuming part, when building a new system from ground, as the data gathering part is hard/slow, if you want a good data quality (and yes, please... we want that).

Below is the documentation for each of these worksheets and the columns inside those.

The Excel file has many worksheets, so this will be a long documentation.

* Worksheets:
  * [Board schematics](#worksheet-board-schematics)
    * Columns:
      * [Schematic name](#column-schematic-name)
      * [Schematic image file](#column-schematic-image-file)
      * [Highlight colors](#columns-highlight-colors)
      * [Highlight opacity](#columns-highlight-opacity)
      * [CAD name](#column-cad-name)
      * [UUID v4](#column-uuid-v4)
  * [Components](#worksheet-components)
    * Columns:
      * [Board label](#column-board-label)
      * [Friendly name](#column-friendly-name)
      * [Technical name or value](#column-technical-name-or-value)
      * [Part-number](#column-part-number)
      * [Category](#column-category)
      * [Region](#column-region)
      * [Short one-liner description](#column-short-one-liner-description)
      * [UUID v4](#column-uuid-v4-1)
  * [Component images](#worksheet-component-images)
    * Columns:
      * [Board label](#column-board-label-1)
      * [Region](#column-region-1)
      * [Pin](#column-pin)
      * [Name](#column-name)
      * [Expected oscilloscope reading](#column-expected-oscilloscope-reading)
      * [T/DIV](#column-tdiv)
      * [V/DIV](#column-vdiv)
      * [T.LVL](#column-tlvl)
      * [File](#column-file)
      * [Note](#column-note)
      * [UUID v4](#column-uuid-v4-2)
  * [Component local files](#worksheet-component-local-files)
    * Columns:
      * [Board label](#column-board-label-2)
      * [Name](#column-name-1)
      * [File](#column-file-1)
      * [UUID v4](#column-uuid-v4-3)
  * [Component links](#worksheet-component-links)
    * Columns:
      * [Board label](#column-board-label-3)
      * [Name](#column-name-2)
      * [URL](#column-url)
      * [UUID v4](#column-uuid-v4-4)
  * [Board local files](#worksheet-board-local-files)
    * Columns:
      * [Category](#column-category-1)
      * [Name](#column-name-3)
      * [File](#column-file-2)
      * [UUID v4](#column-uuid-v4-5)
  * [Board links](#worksheet-board-links)
    * Columns:
      * [Category](#column-category-2)
      * [Name](#column-name-4)
      * [URL](#column-url-1)
      * [UUID v4](#column-uuid-v4-6)
  * [Important signals](#worksheet-important-signals)
    * Columns:
      * [Display name](#column-display-name)
      * [KiCad net name](#column-kicad-net-name)
  * [Credits](#worksheet-credits)
* [Common shared considerations for Excel data files](#common-shared-considerations-for-excel-data-files)

## Worksheet: Board schematics

Depicts which schematic images are available for the board.

These are the columns and how to understand them:

### Column: Schematic name

Exact same name as shown in the thumbnail label.\
Keep the name short, so it can fit in the label.\
A schematic name must be unique in the board.

### Column: Schematic image file

Path and filename to the schematic image file.\
The image format should be either `JPG`, `PNG` or `GIF`.\
Use **relative** path from the `Data` folder.

Ideally this has a fairly high quality, but of course the visibility/performance ratio must be balanced, as the larger image, the harder the display of it gets.

### Columns: Highlight colors

Which color to use for component highlighting in the schematic images, both in the "Main" image but also for the thumbnails.\
All allowed colors are available here, [https://learn.microsoft.com/en-us/dotnet/api/system.windows.media.colors?view=windowsdesktop-7.0](https://learn.microsoft.com/en-us/dotnet/api/system.windows.media.colors?view=windowsdesktop-7.0)

### Columns: Highlight opacity

Where relevant then use a semi-transparent gradient for the highlight, to allow viewing of potential information below the component highlight.\
`0%` equals fully transparent.\
`100%` equals solid non-transparent color.

### Column: CAD name

The exact KiCad display name, which can be found in the logfile, if you have placed all `*_.kicad_*` files in the `KiCad data` folder.

### Column: UUID v4

A unique identifier, used when user contributes data via the "Contribute" tab.\
The UUID can be obtained from https://classic-repair-toolbox.dk/uuid/ and must be globally unique.

## Worksheet: Components

### Column: Board label

Very short label representing the component name.\
Ideally it should be 2-5 characters long only.\
A component label can be represented multiple times, depending if it has regional settings applied.

Do note that `Board label` + `Friendly name` + `Technical name or value` is concatenated in the component list, so do consider to make this as short and precise as possible.

### Column: Friendly name

Typically components have "human readable" or "friendly" names.\
Could also be that component is most often referred to as this name.\
Should still be as short as possible.

Do note that `Board label` + `Friendly name` + `Technical name or value` is concatenated in the component list, so do consider to make this as short and precise as possible.

### Column: Technical name or value

The value of the component or its technical name, depending on its nature.

### Column: Part-number

Typically the part-number from the vendor. In many cases there exists lists of part-numbers, so it can be a good reference, as often you can directly lookup technical details for a part-number.

### Column: Category

Could be `Capacitor`, `Resistor`, `IC`, `Connector`, `Misc` or whatever else suits as a group identified for the component.
Keep the list short, so there is not many categories, but also do make sure to group it where it makes sense.

### Column: Region

Should be either empty (blank), `PAL` or `NTSC`.\
Use only a specific region (PAL or NTSC) when this component is specific for this region only.\
Use empty (blank) when the component is generic, and for no specific region.

### Column: Short one-liner description

Will be shown in the component information popup.\
Is a short contextual and relevant information about the component.\
Could be technical information, which could not fit in "Technical name or value".\
Must be **one line only**!

### Column: UUID v4

A unique identifier, used when user contributes data via the "Contribute" tab.\
The UUID can be obtained from https://classic-repair-toolbox.dk/uuid/ and must be globally unique.

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

### Column: Name

A pinout image should ideally show the legs and what is their input/output.\
If an image is for a specific pin, then document its name for easy reference.

### Column: Expected oscilloscope reading

What value is expected here when measuring this with an oscilloscope?\
This can different things like `LOW`, `HIGH`, `Pulsing`, a frequency or voltage.

### Column: T/DIV

`T/DIV` is the "time per division".\
E.g. `1uS` mean each horizontal division is "1 micro second".

### Column: V/DIV

`V/DIV` is the "volts per division".\
E.g. `1V` mean each vertical division is "1V".

### Column: T.LVL

`T.LVL` is the "trigger level in volts".\
E.g. `1.5V` mean that the scope will trigger at "1.5V".

### Column: File

Path and filename to the image file.\
The image format should be either `JPG`, `PNG` or `GIF`.\
Use **relative** path from the `Data` folder.

### Column: Note

The note field for the image.\
Typically (always?) the first image is the `Pinout` image, and this is special as the `Note` field is used for the component text in the "Resources" tab.

### Column: UUID v4

A unique identifier, used when user contributes data via the "Contribute" tab.\
The UUID can be obtained from https://classic-repair-toolbox.dk/uuid/ and must be globally unique.

## Worksheet: Component local files

Component local files will show in both the `Overview` tab and the component information popup in _CRT_.\
It is a local file specifically for this component - e.g. a datasheet or technical documentation.

### Column: Board label

Direct reference from the `Components` worksheet.\
You can have multiple local files per component, so the board label is allowed to duplicate.

### Column: Name

Name for the file that will be shown in _CRT_.

### Column: File

Path and filename to the local file.\
The local file will be opened in whatever default application you have for the extension.\
Use **relative** path from the `Data` folder.

### Column: UUID v4

A unique identifier, used when user contributes data via the "Contribute" tab.\
The UUID can be obtained from https://classic-repair-toolbox.dk/uuid/ and must be globally unique.

## Worksheet: Component links

Component URLs will show in both the `Overview` tab and the component information popup in _CRT_.\
It is a URL specifically for this component - e.g. a technical documentation or troubleshooting references.

### Column: Board label

Direct reference from the `Components` worksheet.\
You can have multiple links per component, so the board label is allowed to duplicate.

### Column: Name

Name for the link that will be shown in _CRT_.

### Column: URL

The URL will be opened in your default browser.

### Column: UUID v4

A unique identifier, used when user contributes data via the "Contribute" tab.\
The UUID can be obtained from https://classic-repair-toolbox.dk/uuid/ and must be globally unique.

## Worksheet: Board local files

Board local files will show in the `Resources` tab in _CRT_.\
It is meant as a general documentation for the board - e.g. generic diagnosing or troubleshooting.

### Column: Category

What kind of file is this - some examples are `Troubleshooting`, `Technical documentation` or alike.\
You can have multiple local files per category, so the category name is allowed to duplicate.

### Column: Name

Name for the file that will be shown in _CRT_.

### Column: File

Path and filename to the local file.\
The local file will be opened in whatever default application you have for the extension.\
Use **relative** path from the `Data` folder.

### Column: UUID v4

A unique identifier, used when user contributes data via the "Contribute" tab.\
The UUID can be obtained from https://classic-repair-toolbox.dk/uuid/ and must be globally unique.

## Worksheet: Board links

Board URLs will show in the `Resources` tab in _CRT_.\
It is meant as a general documentation for the board - e.g. generic diagnosing or troubleshooting.

### Column: Category

What kind of URL is this - some examples are `Troubleshooting`, `Technical documentation` or alike.\
You can have multiple URLs per category, so the category name is allowed to duplicate.

### Column: Name

Name for the link that will be shown in _CRT_.

### Column: URL

The URL will be opened in your default browser.

### Column: UUID v4

A unique identifier, used when user contributes data via the "Contribute" tab.\
The UUID can be obtained from https://classic-repair-toolbox.dk/uuid/ and must be globally unique.

## Worksheet: Important signals

Will show the important (KiCad) signals to have in the list in the schematics image.\
It does require KiCad data.

### Column: Display name

The name to display in the _CRT_ list in the schematics image.\
Many times the KiCad data is weird to look at, so this is the "human readable" name for it.\
A display name can be repeated many times, as it then will show all the KiCad net names belonging to it.

### Column: KiCad net name

This is data that can be gathered from the logfile, where it will dump all these net names.

## Worksheet: Credits

Will show who has contributed with data to this board.\
Shown in the "About" tab.

# Common shared considerations for Excel data files

There are a few important things to know generally for these Excel files.

* You should not have empty rows in the middle of your data, as this will be considered as _end of data_.
* If there is a yellowish highlight in the data cell, it means the data needs to be validated or corrected manually.
* No formatting (colors, bold or italic etc.) will be carried over from Excel to _CRT_ UI.
* All paths uses `/` instead of `\` (to support Linux and macOS).
* Treat all folders and filenames as case-sensitive (to support Linux and macOS)
* Check logfile after startup, as it will reveal any errors with the data files.

Go to [Wiki Home](https://github.com/HovKlan-DH/Classic-Repair-Toolbox/wiki).\
Go to [Explanation data files](https://github.com/HovKlan-DH/Classic-Repair-Toolbox/wiki/Explanation-of-data-files).