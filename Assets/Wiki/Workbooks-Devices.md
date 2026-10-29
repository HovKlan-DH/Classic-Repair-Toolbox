[Wiki Home](Home) · [Workbooks](Workbooks-tab)

Your own machines, each with its repairs, the changes made to it and its photos.

---

A **device** is one computer you own - "Breadbin #2", "The C128 from the flea market". Its repairs are ordinary workbooks that say which device they are on, so a machine you have fixed three times over the years keeps all three together.

Devices live in the "Workbooks" tab and need the Workbooks feature switched on (see [Getting started](Workbooks-Getting-started)).

## Add a device

"Workbooks" tab -> **Add device** beside "My devices":

| Field | What goes in it |
| --- | --- |
| Name | What you call it - `Breadbin #2` |
| Hardware and Board | The board inside it, from CRT's own lists |
| Region | PAL or NTSC, where the board has both |
| Status | Working, Faulty, In repair, For parts, or Not mine any more |
| Serial number | Optional |
| Where it came from | Optional - a flea market, a friend, new in 1985 |

## The device page

Click a device's card under "My devices" and its page takes the place of the workbook:

* What it is, and its status
* **Repairs** - every workbook on the device, on any board, with the time and cost of all of them added up
* **Changes** - what was done to it outside a repair: a replaced power supply, a fitted ROM. Pick the date, type what was changed, **Add change**. Swapping the board in **Edit device** is recorded here by itself
* **Notes** - saved when you leave the box - and **photos** (**Add photos**)

Buttons on the page:

* **New repair** - a new workbook on this device, on its board
* **Edit device** and **Delete device** - deleting asks first, and its repairs stay in the workbook list, on no device
* **Export history to PDF** / **Export history to ZIP** - see below

## Repairs on a device

* **Create new workbook** always asks **On my device**: None, or one of your devices with the board shown. **Add device** beside it adds one there and then.
* Under a workbook's title in the "Workbooks" tab, **On my device** moves an existing workbook onto a device (or off it).
* The workbook drop-down in the bar names the device: `#3 | Breadbin #2 - No video (C64/250407)`.

## The device picker

Once you have a device, a **Device** drop-down sits above **Hardware**. It shows the device the bar's workbook is on, or "No device".

Pick a device and CRT goes to its repair: the newest open one on any board, else the newest on its board - switching board first. A device with no repairs yet gets its board selected and the "New repair on Breadbin #2" form; cancel it and the picker goes back.

## Export history

**Export history to PDF** writes one document for the whole machine: the device on every page, then what it is, the totals of every repair, the list of repairs, its changes and photos - and after that every repair, oldest first, each drawn as its own workbook export on its own board's pictures.

**Export history to ZIP** holds the same PDF, the device's photos under `device photos/`, and each repair's files under `workbook_<id>/worklog_<id>/`.

## Where it is stored

Beside your workbooks, one folder per device:

```
Workbooks/
├── index.json          <- also the next device number
├── device_1/
│   ├── index.json      <- the device, its changes and notes
│   └── photo-1.jpg     <- its photos
└── workbook_3/
    └── index.json      <- a repair, naming its device
```

As with workbooks, a deleted device's number is never handed out again, and copying the `Workbooks` folder backs up everything.

## If it does not work

| What you see | Almost always means |
| --- | --- |
| No **Device** drop-down above Hardware | No device added yet, or the Workbooks feature is off in "Configuration" |
| A device is not offered in Create new workbook | It is on another board - only devices with the board shown are offered |
| Picking a device does nothing | Its board is unticked in the "Configuration" tab's list of boards shown |
