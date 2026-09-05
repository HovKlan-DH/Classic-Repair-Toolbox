# Searching

← [Back to Workbooks](Workbooks)

The **"Find a previous repair"** box at the top of the [Workbooks tab](Workbooks-The-Workbooks-tab) filters
the whole tab as you type. It exists for the question this feature is really for: *"have I seen this
fault before?"*

---

## What it filters

All three panels narrow together:

- the **workbook list** — only workbooks that matched,
- the **board pane** — only previews with matching worklogs,
- the **worklog list** — only the matching worklogs.

Matched text is **highlighted** wherever it is drawn, so you can see *why* something matched.

A workbook is shown when **its own text matches, or any worklog in it does**. If the workbook itself
matched but no individual worklog did, **all** of its worklogs stay visible — otherwise the result
would look empty.

---

## The query grammar

| You type | It means |
| --- | --- |
| `cpu` | Records containing "cpu" |
| `cpu super` | Records containing **both** "cpu" **and** "super" — a space is AND |
| `"full text"` | The quoted run as **one** term, spaces included |
| `-cpu` | Records **not** containing "cpu" |
| `-"full text"` | Exclusion works on quoted terms too |
| *(empty)* | Everything — an empty box is not a filter |

**Matching is case-insensitive substring matching throughout.** That has two consequences worth
knowing:

- A term matches **mid-word**. `p c u` finds "CPU", because each of `p`, `c` and `u` is present
  somewhere in the record. Short terms are broad.
- `"full text"` also finds `Afull textB` — it is a substring, not a whole-word match. It would *not*
  find `full-text`.

Terms are ANDed **across the whole record**, not within one field. `cpu socket` matches a worklog
whose title says "CPU" and whose comment three months later says "socket".

---

## What is searched

Everything you typed, and nothing you did not.

**On the workbook:** its description and its note.

**On each worklog:** title, description, category, schematic name, every component label in scope,
and every row of every list — link headlines and URLs, comments, work-done notes, photo and file
names, photo and file comments.

So a search reaches the comment somebody left three months ago, and the name of a photo you
remember filing.

### What is deliberately excluded

**Numbers** — ids, hours, cost, display order, dates. A search for `2` would otherwise match nearly
everything through fields you never see as text.

**Open / Closed** — both the workbook status and the worklog state. Every record carries one of two
values, and "open" is a word this domain uses constantly ("open circuit", "opened the case", "open
trace on CN2"). Including them made "open" match almost the whole database — and because terms are
ANDed, `open trace` then matched any Open workbook that mentioned "trace" anywhere.

Both values already have an always-visible pill, so they filter far better by eye than by substring.

**Category is included**, unlike state — Note, Cosmetic and Issue are descriptive words you chose,
and none of them turns up incidentally in repair notes.

---

## Behaviour worth knowing

**Searching never changes which workbook is active.** It narrows what you *see*; it does not
redirect where "Add worklog" writes. If the active workbook is filtered out, the header and the
right-hand side move to a workbook that survived — but the real active workbook is untouched, and
comes back when you clear the box.

**The query is cleared when you change board.** A board change is a change of subject, and carrying
a filter across would land you on an empty list for a board you just chose, with the reason sitting
in a box you are not looking at.

**Every other refresh keeps it.** Saving a worklog, creating a workbook, deleting one — none of these
silently drop you back to the unfiltered list.

**Typing is debounced.** Filtering reads every workbook's worklogs from disk, so the app waits a
fraction of a second after you stop typing rather than doing it on every keystroke.

**Empty states say which kind of empty they are.** A search that matched nothing says *"No workbooks
or worklogs match your search"*, not "none recorded yet" — the second reads as data loss.

---

## Examples

| Goal | Query |
| --- | --- |
| Every job involving the VIC-II | `u19` |
| Everything mentioning a cracked socket | `"cracked socket"` |
| PSU work, but not the ones about the 9V rail | `psu -9v` |
| A customer's jobs | `hansen` |
| Cosmetic work only | `cosmetic` |
| That capacitor you remember replacing | `c64 replaced` |

---

**Next:** [The summary strip](Workbooks-The-summary-strip) · [The Workbooks tab](Workbooks-The-Workbooks-tab)
