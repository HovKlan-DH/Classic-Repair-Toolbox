Go to [Wiki Home](https://github.com/HovKlan-DH/Classic-Repair-Toolbox/wiki).\
Go to [Explanation data files](https://github.com/HovKlan-DH/Classic-Repair-Toolbox/wiki/Explanation-of-data-files).

The main Excel file is the one placed in the data root folder, named `Classic-Repair-Toolbox.xlsx` or `Classic-Repair-Toolbox.v2.0.0.xlsx`. This file depicts the different hardware in _CRT_, and it has the configuration for all supported oscilloscopes.

The Excel file, alongside the board Excel files, is versionized meaning the version number represents the file will work from _FROM this version and onwards_.

The Excel file has only two worksheets, so it is pretty simple compared to the board files.

* Worksheets:
  * [Hardware & Board](#worksheet-hardware--board)
    * Columns:
      * [Hardware name in drop-down](#column-hardware-name-in-drop-dow)
      * [Board name in drop-down](#column-board-name-in-drop-down)
      * [Excel data file](#column-excel-data-file)
      * [Hardware notes in "Overview" tab](#column-hardware-notes-in-overview-tab)
  * [Oscilloscope](#worksheet-oscilloscope)
    * Columns:
      * [Brand](#column-brand)
      * [Series or model](#column-series-or-model)
      * [Port](#column-port)
      * [Debounce-Time](#column-debounce-time)
      * [Identify](#column-identify)
      * [DrainErrorQueue](#column-drainerrorqueue)
      * [Operation-Complete](#column-operation-complete)
      * [Clear-Statistics](#column-clear-statistics)
      * [QueryActiveTrigger](#column-queryactivetrigger)
      * [Stop](#column-stop)
      * [Single](#column-single)
      * [Run](#column-run)
      * [QueryTriggerMode](#column-querytriggermode)
      * [QueryTriggerLevel](#column-querytriggerlevel)
      * [SetTriggerLevel](#column-settriggerlevel)
      * [QueryTimeDiv](#column-querytimediv)
      * [SetTimeDiv](#column-settimediv)
      * [QueryVoltsDiv](#column-queryvoltsdiv)
      * [SetVoltsDiv](#column-setvoltsdiv)
      * [DumpImage](#column-dumpimage)
      * [TIME/DIV](#column-timediv)
      * [VOLTS/DIV](#column-voltsdiv)
    * [Important notice](#important-notice)
* [Common shared considerations for Excel data files](#common-shared-considerations-for-excel-data-files)

## Worksheet: Hardware & Board

These are the columns and how to understand them:

### Column: Hardware name in drop-down

Exact same name as shown in the `Hardware` drop-down list in _CRT_.\
The exact same hardware name can be represented on multiple rows, if it has different boards.\
Keep the name short, so it can fit in the drop-down.

### Column: Board name in drop-down

Exact same name as shown in the `Board` drop-down list in _CRT_.\
The same board name should _not_ be represented more than once per hardware.\
Keep the name short, so it can fit in the drop-down.\
This is typically the model number or for a Commodore system it is the "assy" (yeah, weird naming).

### Column: Excel data file

Path and filename to the specific hardware/board Excel data file.\
Use **relative** path from the `Data` folder.

### Column: Hardware notes in "Overview" tab

A small note describing characteristics of the hardware - maybe special recognition points.\
Currently this data is not displayed anywhere, but it should be one day.

## Worksheet: Oscilloscope

These are the columns and how to understand them:

### Column: Brand

The brand name of the oscilloscope vendor.

### Column: Series or model

Ideally an entire series/family can be covered, but this could also be individual oscilloscope models.

### Column: Port

Default port suggested for network connectivity. Can be changed in the "Oscilloscope" tab.

### Column: Debounce-Time

The _debounce time_ is a pacing value that tells _CRT_ how long to wait before it talks to your oscilloscope between commands. Note that this is by far not all commands, but only some (e.g. image handling). It is also used for keyboard debouncing.

### Column: Identify

SCPI command for identifying the scope. 

For my Rigol scope it will show this:
> Vendor: RIGOL TECHNOLOGIES\
> Model: DS2202A\
> Serial: *************\
> Firmware: 00.03.06

### Column: DrainErrorQueue

SCPI command for giving the next error in the system error queue.

### Column: Operation-Complete

SCPI command to query if the scope is ready for the next command.

### Column: Clear-Statistics

SCPI command for clearing statistics/status values in scope.

### Column: QueryActiveTrigger

SCPI command to query the scope what its trigger/acquisition state is right now. This could be `Running`, `Waiting` or `triggered` etc.

### Column: Stop

SCPI command to stop the scope here-and-now (freeze the acquisition).

### Column: Single

SCPI command to stop the scope at the next trigger point.

### Column: Run

SCPI command to resume the acquisition.

### Column: QueryTriggerMode

SCPI command to query the trigger mode - e.g. `Edge` or alike.

### Column: QueryTriggerLevel

SCPI command to query the trigger level (voltage).

### Column: SetTriggerLevel

SCPI command to set the trigger (will be both the trigger mode and the trigger level).

### Column: QueryTimeDiv

SCPI command to query the `T/DIV` (time per division).\
E.g. `1uS` mean each horizontal division is `1 micro second`.

### Column: SetTimeDiv

SCPI command to set the `T/DIV` (time per division).

### Column: QueryVoltsDiv

SCPI command to query the `V/DIV` (volts per division).\
E.g. `1V` mean each vertical division is `1V`.

### Column: SetVoltsDiv

SCPI command to set the `V/DIV` (volts per division).

### Column: DumpImage

SCPI command to dump the current image from scope via network.
This is quite different per vendor and even per scope.

### Column: TIME/DIV

A list of all the `T/DIV` being available for the scope.\
E.g. `1uS` mean each horizontal division is `1 micro second`.

### Column: VOLTS/DIV

A list of all the `V/DIV` being available for the scope.\
E.g. `1V` mean each vertical division is `1V`.

### Important notice

> [!WARNING]
> If some of the SCPI commands (in the Excel file) it states `{0}` which you should see as the actual value passed. The value is handled by _CRT_.

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