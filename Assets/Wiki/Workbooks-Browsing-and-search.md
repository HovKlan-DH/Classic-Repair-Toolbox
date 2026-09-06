[Wiki Home](Home) · [Workbooks](Workbooks-tab)

Where you browse your repair jobs.

---

```
┌───────────────┬──────────────────────────────────────────────┐
│ Find a repair │ #3 · Dead C64, no video           [Open]     │
│ [__________]  │ Reported as dead after a thunderstorm        │
│               │      [Edit workbook] [Delete workbook]       │
│ 3 workbooks   │      [Export to PDF] [Export to ZIP]         │
│               │ ▸ 3 worklogs · 2 hours · 160 · 1 open        │
│ ┌───────────┐ ├──────────────────┬───────────────────────────┤
│ │#3 Dead C64│ │ ┌──────┐ ┌─────┐ │ Worklogs on "Video"       │
│ │  Open     │ │ │Power │ │Video│ │ ┌───────────────────────┐ │
│ └───────────┘ │ │ [#1] │ │[#2] │ │ │ ② VIC socket      ✕  │ │
│ ┌───────────┐ │ └──────┘ └─────┘ │ │ Pin 8 lifted          │ │
│ │#2 Tape …  │ │                  │ │ [Issue] [Closed]      │ │
│ └───────────┘ │                  │ │ 1 hour and 30 minutes │ │
└───────────────┴──────────────────┴───────────────────────────┘
      ①                  ②                      ③
```

## ① Workbook list

One card per repair, newest first. **Click a card** to switch to that workbook - the bar, the board and "Add worklog" all follow it.

Which workbooks are listed depends on the scope setting in "Configuration": this board only (default), or all boards. With all boards, each card also names its board, and clicking one from another board switches the application to it.

## ② Board pane

Every schematic that has worklogs in the selected workbook, with the markers drawn on it.

* **Click a marker** - opens that worklog in the editor
* **Click anywhere else on a schematic** - selects it, and the list on the right switches to its worklogs
* **Drag a schematic by its panel** - moves it up or down the pane

Grab it anywhere on the panel except the board image itself - the title, or the space around it. The pointer turns into an up/down arrow there; over the image it stays a hand, because clicking the image selects the schematic instead. A dashed slot shows where the schematic will land; release to drop it there.

The order is saved with that workbook, so it is the order you left it in next time you open it. Each workbook keeps its own. A schematic that gets its first worklog later joins at the bottom.

## ③ Worklog list

One card per worklog on the selected schematic: number and title, description, category and state, and a line of totals (time, cost, and how many comments/links/photos/files it holds). Time is shown in hours and minutes - `45 minutes`, `1 hour and 15 minutes` - never as a decimal.

**That line only shows what the worklog actually has.** No time logged, no cost, no photos - each is simply left off rather than listed as a zero, so a plain note reads `1 comment` instead of `0 h · 0 USD · 1 comment · 0 links · 0 photos · 0 files`.

**Click a card** to open the worklog. **The ✕ in its corner deletes it** - see below.

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
▸  7 worklogs · 12 hours and 30 minutes · 430 · 4 open
```

Click the arrow to expand it into a breakdown by category, by state, by attachment, and by component.

Time and cost come from the Work done lines in every worklog. The time is always said in hours and minutes rather than as decimal hours, and **a workbook with no time logged, or no cost recorded, simply leaves that figure out** - a headline reading `1 worklog · 1 open` has neither yet. The count of open worklogs always shows, including `0 open`, because that one says the repair is finished.

The cost carries the currency code you picked in Configuration (`430 DKK`), so a figure read on its own still says what it is.

The expanded breakdown below is different on purpose: it keeps every zero, including the category and state pills. You open it to see the whole picture, where `0 Issue` is an answer - and a row of pills that changed width as you worked would be harder to read at a glance.

## Find a previous repair

The search box filters everything - the workbook list, the board pane and the worklog list - and highlights what matched.

| You type | It finds |
| --- | --- |
| `cpu` | Anything containing "cpu" |
| `cpu socket` | Records containing **both** words, anywhere and in any order |
| `"cracked socket"` | Those two words as one run, in that order |
| `-psu` | Everything **except** records containing "psu" |
| `socket -psu` | Records with "socket" but **without** "psu" |
| `-"ruled out"` | Everything except records containing that exact run |

A space means **and**, not or - every word you type has to be found somewhere in the record before
it is shown, so each word you add narrows the result further. Quotes are what let a term contain a
space: `"cracked socket"` is one term and only matches those words together, where `cracked socket`
is two terms that can match anywhere in the record, in either order. A minus in front of a term
removes anything containing it, and works on a quoted run too.

**Matching is on any part of a word, and case does not matter.** `cap` finds "capacitor" and
"Capacitor"; `410` finds "250410". This is why searching for a couple of letters can bring back
more than you expect - add another word to narrow it rather than reaching for the exact spelling.

Some worked examples, on a board you have repaired several times:

| You type | Why you would |
| --- | --- |
| `u8` | Every repair with U8 among its components in scope |
| `u8 replaced` | Only the ones where you wrote that you replaced it |
| `"cold solder"` | The phrase, not every record containing "cold" and "solder" apart |
| `ram -ruled` | RAM faults, minus the ones you ruled out |
| `1541 belt` | Narrowing to one repair when you remember only two details |

Searching for a component works because the components you tick in a worklog are part of its text -
so `u8` finds the repairs that had U8 in scope, even if you never typed "U8" in the description.
Remember the substring rule though: `u8` also matches U80 and U81, so add a second word when a
board has components numbered that way.

It searches everything you have typed: titles, descriptions, comments, work done, component names, link and file names.

It does **not** search numbers (hours, cost, dates, id numbers) or the words Open/Closed - those two would match nearly everything, and both already have a pill you can see.

Searching never changes which workbook you are writing into.

**The filter stays until you clear it** - use the ✕ at the end of the box. It survives changing board, so with "Show all workbooks" you can search, click a result that lives on another board, and still see the filter and its highlighting once that board has loaded.

**The highlighting follows you into the worklog itself.** Open a worklog while a search is running and the matches are marked in there too - in its comments, work done, link headlines, and the comments on its photos and files. That matters when the match is not in the title: the search can find a worklog by a comment written months ago, and this is what shows you which one it was. The title and description boxes at the top are not marked, because those are editable fields.

---

**Next:** [Export and your data](Workbooks-Export-and-data) — PDF/ZIP, where files live, deleting
