[Wiki Home](Home) · [Data files](Explanation-of-data-files)

The top-level file: every hardware and board, and every oscilloscope.

---

The main Excel file is the one placed in the data root folder, named `Classic-Repair-Toolbox.xlsx` or `Classic-Repair-Toolbox.v2.0.0.xlsx`. This file depicts the different hardware in CRT, and it has the configuration for all supported oscilloscopes.

The Excel file, alongside the board Excel files, is versioned meaning the version number represents the file will work from _FROM this version and onwards_.

The Excel file has only two worksheets, so it is pretty simple compared to the board files.


## Worksheet: Hardware & Board

These are the columns and how to understand them:

### Column: Hardware name in drop-down

Exact same name as shown in the `Hardware` drop-down list in CRT.\
The exact same hardware name can be represented on multiple rows, if it has different boards.\
Keep the name short, so it can fit in the drop-down.

### Column: Board name in drop-down

Exact same name as shown in the `Board` drop-down list in CRT.\
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

One row per scope brand/series. Copy the row for a scope closest to yours and adjust it — **you will
need your scope's programming manual**, because SCPI commands vary between vendors and even between
models from the same vendor.

`{0}` in a command is where CRT substitutes the actual value.

| Column | What goes in it |
| --- | --- |
| `Brand` | The vendor name |
| `Series or model` | Ideally a whole series/family, but a single model works too |
| `Port` | Default network port. Changeable in the "Oscilloscope" tab. |
| `Debounce-Time` | How long to wait between commands, in ms. Used for image handling and keyboard repeat. |
| `Identify` | Returns vendor, model, serial and firmware |
| `DrainErrorQueue` | Returns the next error from the scope's error queue |
| `Operation-Complete` | Asks whether the scope is ready for the next command |
| `Clear-Statistics` | Clears statistics/status values |
| `QueryActiveTrigger` | Returns the trigger state: `Running`, `Waiting`, `triggered` |
| `Stop` | Freezes acquisition now |
| `Single` | Stops at the next trigger point |
| `Run` | Resumes acquisition |
| `QueryTriggerMode` | Returns the trigger mode, e.g. `Edge` |
| `QueryTriggerLevel` | Returns the trigger level in volts |
| `SetTriggerLevel` | Sets it — writes both the trigger mode and the level |
| `QueryTimeDiv` / `SetTimeDiv` | Time per division, e.g. `1uS` = 1 microsecond per horizontal division |
| `QueryVoltsDiv` / `SetVoltsDiv` | Volts per division, e.g. `1V` per vertical division |
| `DumpImage` | Fetches the current screen over the network. Differs a lot per vendor and model. |
| `TIME/DIV` | Every `T/DIV` value your scope offers |
| `VOLTS/DIV` | Every `V/DIV` value your scope offers |

> [!WARNING]
> `TIME/DIV` and `VOLTS/DIV` must both be filled in, not just the command columns. The numpad
> stepping keys look up the current value in those lists, and do nothing if it is not there.


## Rules for every Excel file

See [Data files](Explanation-of-data-files#rules-for-every-excel-file) — the same rules apply to every Excel file.
