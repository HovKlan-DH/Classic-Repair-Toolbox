[Wiki Home](Home) · [Workbooks](Workbooks-tab)

Where you browse your repairs.

---

```
┌────────────────────────────────────────────────────────────────────────────────┐
│ [Find a previous repair - words are ANDed, "use quotes" for a ...]             │
├───────────────┬────────────────────────────────────────────────────────────────┤
│ 3 workbooks   │ #3 · Dead C64  [Open]        [Edit workbook] [Delete workbook] │
│               │ No video after a storm         [Export to PDF] [Export to ZIP] │
│ ┌───────────┐ │ ▸ 3 worklogs · 2 hours · 160 DKK · 1 open                      │
│ │#3   [Open]│ ├──────────────────────────┬─────────────────────────────────────┤
│ │Dead C64   │ │ Power                    │ Video · 1 worklog                   │
│ │3 worklogs │ │ ┌──────────────────┐     │ ┌─────────────────────────────────┐ │
│ └───────────┘ │ │       [#1]       │     │ │ VIC socket [#2] [Delete worklog]│ │
│ ┌───────────┐ │ └──────────────────┘     │ │ Pin 8 lifted                    │ │
│ │#2 [Closed]│ │ Video                    │ │ [Issue] [Closed]                │ │
│ │No sound   │ │ ┏━━━━━━━━━━━━━━━━━━┓     │ │ 1 hour and 30 minutes · 160 DKK │ │
│ └───────────┘ │ ┃ [#2]             ┃     │ └─────────────────────────────────┘ │
│               │ ┗━━━━━━━━━━━━━━━━━━┛     │                                     │
└───────────────┴──────────────────────────┴─────────────────────────────────────┘
        ①                     ②                               ③
```

The search box runs across the top of the tab and filters both sides - see "Find a previous repair" below.

## ① Workbook list

One card per repair, newest first. **Click a card** to switch to that workbook - the bar, the board and "Add worklog" all follow it.

Which workbooks are listed depends on the scope setting in "Configuration": this board only (default), or all boards. With all boards, each card also names its board, and clicking one from another board switches the application to it.

## ② Board pane

Every schematic that has worklogs in the selected workbook, with the markers drawn on it. The selected schematic is outlined.

* **Click a marker** - opens that worklog in the editor
* **Click anywhere else on a schematic** - selects it, and the list on the right switches to its worklogs
* **Drag a schematic by its panel** - moves it up or down the pane

Grab it anywhere on the panel except the board image itself - the title, or the space around it. The pointer turns into an up/down arrow there; over the image it stays a hand, because clicking the image selects the schematic instead. A dashed slot shows where the schematic will land; release to drop it there.

The order is saved with that workbook, so it is the order you left it in next time you open it. Each workbook keeps its own. A schematic that gets its first worklog later joins at the bottom.

Dragging is switched off while a search is active, because the pane then shows only some of the schematics. The exported PDF does not follow this order - it lists the schematics alphabetically.

## ③ Worklog list

Headed by the selected schematic's name and how many worklogs it has (`Video · 1 worklog`). One card per worklog on that schematic: title and number, description, category and state, and a line of totals (time, cost, and how many comments/links/photos/files it holds). Time is shown in hours and minutes - `45 minutes`, `1 hour and 15 minutes` - never as a decimal.

**That line only shows what the worklog actually has.** No time logged, no cost, no photos - each is simply left off rather than listed as a zero, so a plain note reads `1 comment` instead of `0 h · 0 USD · 1 comment · 0 links · 0 photos · 0 files`.

**Click a card** to open the worklog. **The "Delete worklog" button in its corner deletes it** - see [Export and your data](Workbooks-Export-and-data).

## The header

The selected workbook's number, title, status and note, plus four buttons:

| Button | Does |
| --- | --- |
| Edit workbook | Change the description and note |
| Delete workbook | Deletes the whole repair - see [Export and your data](Workbooks-Export-and-data) |
| Export to PDF | The write-up on its own |
| Export to ZIP | That PDF plus the original photos and files |

## The totals strip

Under the header:

```
▸  7 worklogs · 12 hours and 30 minutes · 430 DKK · 4 open
```

Click it to expand it into a breakdown by category, by state, by attachment, and by component. It stays expanded or collapsed the way you left it.

Time and cost come from the Work done lines in every worklog. The time is always said in hours and minutes rather than as decimal hours, and **a workbook with no time logged, or no cost recorded, simply leaves that figure out** - a headline reading `1 worklog · 1 open` has neither yet. The count of open worklogs always shows, including `0 open`, because that one says the repair is finished.

The cost carries the currency code you picked in Configuration (`430 DKK`), so a figure read on its own still says what it is.

The expanded breakdown below is different on purpose: it keeps every zero, including the category and state pills. You open it to see the whole picture, where `0 Issue` is an answer - and a row of pills that changed width as you worked would be harder to read at a glance. The one exception is the components line, which only appears once the workbook has components in scope.

## Find a previous repair

The search box across the top of the tab filters everything - the workbook list, the board pane and the worklog list - and highlights what matched.

| You type | It finds |
| --- | --- |
| `cpu` | Anything containing "cpu" |
| `cpu socket` | A worklog containing **both** words, anywhere in it and in any order |
| `"cracked socket"` | Those two words as one run, in that order |
| `-psu` | Everything **except** worklogs containing "psu" |
| `socket -psu` | Worklogs with "socket" but **without** "psu" |
| `-"ruled out"` | Everything except worklogs containing that exact run |

A space means **and**, not or - and every word has to be found in the **same worklog**, or all of them in the workbook's own description and note. A repair whose description says "C64" and one of whose worklogs says "socket" is not found by `c64 socket`, because the two words are in different places. Each word you add narrows the result further. Quotes are what let a term contain a space: `"cracked socket"` is one term and only matches those words together, where `cracked socket` is two terms that can match anywhere in the worklog, in either order. A minus in front of a term removes anything containing it, and works on a quoted run too.

A workbook is listed when its own description and note match, or when at least one of its worklogs does. When only the workbook itself matched, all of its worklogs stay visible.

**Matching is on any part of a word, and case does not matter.** `cap` finds "capacitor" and
"Capacitor"; `410` finds "250410". This is why searching for a couple of letters can bring back
more than you expect - add another word to narrow it rather than reaching for the exact spelling.

Some worked examples, on a board you have repaired several times:

| You type | Why you would |
| --- | --- |
| `u8` | Every repair with U8 among its components in scope |
| `u8 replaced` | Only the worklogs where you wrote that you replaced it |
| `"cold solder"` | The phrase, not every worklog containing "cold" and "solder" apart |
| `ram -ruled` | RAM faults, minus the ones you ruled out |
| `attic storm` | A repair whose description or note mentions both, when that is all you remember |

Searching for a component works because the components you tick in a worklog are part of its text -
so `u8` finds the repairs that had U8 in scope, even if you never typed "U8" in the description.
Remember the substring rule though: `u8` also matches U80 and U81, so add a second word when a
board has components numbered that way.

It searches everything you have typed: the workbook's description and note, and in each worklog its title, description, comments, work done, links (headline and address), photo and file names and their comments, and component names. It also searches each worklog's category and the name of its schematic - so `issue` finds every Issue worklog, and a schematic's name finds every worklog on it.

It does **not** search numbers (hours, cost, dates, id numbers) or the words Open/Closed - those two would match nearly everything, and both already have a pill you can see.

Searching never changes which workbook you are writing into.

**The filter stays until you clear it** - use the ✕ at the end of the box. It survives changing board, so with "Show all workbooks" you can search, click a result that lives on another board, and still see the filter and its highlighting once that board has loaded.

**The highlighting follows you into the worklog itself.** Open a worklog while a search is running and the matches are marked in there too - in its comments, work done, link headlines, and the comments on its photos and files. That matters when the match is not in the title: the search can find a worklog by a comment written months ago, and this is what shows you which one it was. The title and description boxes at the top are not marked, because those are editable fields.

---

**Next:** [Export and your data](Workbooks-Export-and-data) - PDF/ZIP, where files live, deleting
