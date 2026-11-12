[Wiki Home](Home) · [Workbooks](Workbooks-tab)

Your own machines, each with its workbooks, comments and photos.

---

A **device** is one computer you own - "Breadbin #2", "The C128 from the flea market". Its repairs are ordinary workbooks that say which device they are on, so a machine you have fixed three times over the years keeps all three together.

Devices live in the "Workbooks" tab and need the Workbooks feature switched on (see [Getting started](Workbooks-Getting-started)).

## Add a device

**Add device** in the bar at the top, after **Add workbook**. No workbook is needed - add your machines first, and start repairs on them when you need to.

The form is laid out like the worklog editor: what the device is and its comments on the left, its photos on the right.

| Field | What goes in it |
| --- | --- |
| Name | What you call it - `Breadbin #2` |
| Hardware and Board | The board inside it, from CRT's own lists - any board, not only the one shown. Locked while the device has workbooks, as a workbook belongs to a device of its own board |
| Status | Unknown, Working, Faulty, In repair or For parts |
| Region | Unknown, or one of the regions the board comes in (PAL and NTSC, or its own) - only offered for a board that comes in regions |
| Condition | 0 to 5 stars - click a star; click the same star again for none |
| Serial number or unique identifier | Optional |
| Description | Optional, as many lines as you like - where it came from, what came with it. It can be formatted - see [Formatting your text](Workbooks-Daily-use#formatting-your-text) |

* **Links of interest** - **Add link**, a headline and an address, as on a worklog. Click a link to open it; the pen changes it, the bin deletes it.
* **Comments** - **Add comment**, each one dated. Click a comment to change it; the bin deletes it.
* **Photos** - **Add photo**, each with a comment of its own. Drag a photo by the space beside its picture to put it in another place. Click a photo to view it whole: the mouse wheel zooms in and out, a drag with the right or middle mouse button moves it, and a double-click shows it whole again.

On a device you edit, links, comments and photos are saved as soon as you add, change, move or delete them - as on a worklog - so **Cancel** does not take them back. On a new device they wait in the form and are saved with **Add device**.

**Condition** is drawn as gold stars; pointing at a star lights the stars up to it before you click.

**Enter** saves the form from any box but **Description**, where it starts a new line - there **Ctrl+Enter** saves.

## My devices

On the "Workbooks" tab, "My devices" stands on the left of the window, below **Hardware** and **Board**, with how many devices it lists beside the heading. Each device's card shows its name and status, then its serial number (when it has one) with how many workbooks it has, then its hardware and board - with its condition in stars at the end of that line. Which devices are listed is the scope setting in "Configuration", and the search box above the workbooks finds devices too - by name, serial number or anything typed on them.

**Sorting of devices**, in "Configuration" under the scope setting, puts them in order: by name (the default), by hardware, board and name, or by hand - then each card has a handle at its right edge to drag it up or down. A device added while sorting by hand goes last.

Click a device to **select** it: its page opens, the workbook list shows only its workbooks (on any board), and the **Device** drop-down in the bar picks it too - it is the same choice. Click another device to select that one instead. Click a workbook that is on a device, and that device is selected. **Clear**, under the list, unselects the device and shows every workbook again.

## The device page

Click a device's card under "My devices" and its page takes the place of the workbook. The page only shows the device - everything on it is added and changed with **Edit device**:

* What it is, its status, its condition in stars, and its description
* **Workbooks associated to this device** - every workbook on the device, on any board, with the time and cost of all of them added up
* **Links of interest** and **Comments** - click a link to open it
* **Photos** - each in a panel of its own with its comment under it; click a photo to view it and zoom in

A part the device has nothing in - no links, comments or photos - is left off the page.

Buttons on the page:

* **Edit device** and **Delete device** - deleting asks first, and its workbooks stay in the workbook list, on no device
* **Export history to PDF** / **Export history to ZIP** - see below

## Repairs on a device

* **Add workbook** always asks **Associate workbook to device**: None, or one of your devices with the board shown. **Add device** beside it adds one there and then, on that board.
* Under a workbook's title in the "Workbooks" tab, **Associate workbook to device** moves an existing workbook onto a device (or off it).

## The device drop-down

The bar at the top starts with a **Device** drop-down - the same choice as the device selected under "My devices". **Show workbooks for all devices** comes first and is picked until you pick something else; then come your devices, each with its status and, last on the line, how many workbooks it has. CRT remembers which you picked. While **Show workbooks for all devices** is picked, each line of the **Workbook** drop-down starts with the name of the device the workbook is on.

**While a device is picked, you work only in its workbooks.** The **Workbook** drop-down lists only its workbooks, and **Add worklog** writes into the one picked there. With none of its workbooks on the board shown, the **Workbook** drop-down says so and **Add worklog** is hidden until **Add workbook** makes one.

Pick a device and CRT goes to its repair: the newest open one on any board, else the newest on its board - switching board first. A device with no workbook to work on has its own board shown and its page opened in the "Workbooks" tab; **Add workbook** in the bar starts one on it.

Working in a workbook of another device picks that device. A workbook on no device - or switching to a board with none of the device's workbooks - goes back to **Show workbooks for all devices**.

## Export history

**Export history to PDF** writes one document for the whole machine: the device on every page, then what it is - its condition, description, links and comments - the totals of every repair, the list of workbooks associated to the device and its photos - and after that every repair, oldest first, each drawn as its own workbook export on its own board's schematic images.

**Export history to ZIP** holds the same PDF, the device's photos under `device photos/`, and each repair's files under `workbook_<id>/worklog_<id>/`.

## Where it is stored

Beside your workbooks, one folder per device:

```
Workbooks/
├── index.json          <- also the next device number
├── device_1/
│   ├── index.json      <- the device, its links and comments
│   └── photo-1.jpg     <- its photos
└── workbook_3/
    └── index.json      <- a repair, naming its device
```

As with workbooks, a deleted device's number is never handed out again, and copying the `Workbooks` folder backs up everything.

## If it does not work

| What you see | Almost always means |
| --- | --- |
| The **Device** drop-down is greyed out | No device added yet |
| No bar at the top at all | The Workbooks feature is off in "Configuration" |
| No **Region** in the device's form | The board picked does not come in regional versions |
| A device is not offered in **Add workbook** | It is on another board - only devices with the board shown are offered |
| Picking a device opens its page instead of a workbook | It has no workbook yet, or only on a board unticked in the "Configuration" tab's list of boards shown |
| The **Device** drop-down went back to **Show workbooks for all devices** | You worked in a workbook on no device, or switched to a board with none of the device's workbooks |
| **Add worklog** is gone | The device picked has no workbook on the board shown - **Add workbook** makes one, or pick **Show workbooks for all devices** (or **Clear** under "My devices"). Or **Clear** under the workbooks was clicked - click a workbook |
| **Hardware** and **Board** are greyed out in **Edit device** | The device has workbooks - they stay on its board |
