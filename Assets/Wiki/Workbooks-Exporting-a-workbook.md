# Exporting a workbook

← [Back to Workbooks](Workbooks)

A workbook can be exported as a document to hand to a customer, or as an archive that carries the
original files with it.

Both buttons sit in the [Workbooks tab](Workbooks-The-Workbooks-tab) header, on their own row under
Edit/Delete:

> **[Export to PDF]  [Export to ZIP]**

They export the **selected** workbook.

---

## Which format

| | **PDF** | **ZIP** |
| --- | --- | --- |
| The document | ✔ | ✔ (the same PDF, inside) |
| Photos | Embedded, at page resolution | **The original files** |
| Attached files | ✘ — cannot be embedded | ✔ |
| Use it for | Emailing a customer a report | Handing over everything, or archiving the job |

The PDF is the customer-facing document. The ZIP is that same PDF **plus** every photo and attached
file, one folder per worklog — because a PDF shows a photo downscaled to fit a page and cannot carry
a datasheet at all.

---

## File naming

Both exports are suggested as:

```
Workbook_3_Commodore_C64_20260904
```

That is `Workbook`, the workbook's **id**, the **hardware** and **board**, and today's date.

**The workbook's description is deliberately not in the name.** It is a sentence you typed, often
carrying a customer's own details, on a file that is about to be emailed.

You can rename it in the save dialog. The extension is set by the **button you pressed**, not by what
you type — typing `repair.pdf` into the ZIP dialog gives you `repair.zip`, not `repair.pdf.zip`. An
unrelated suffix is left alone: `board rev 2.5` becomes `board rev 2.5.zip`.

---

## What is in the PDF

**A summary section** — the same totals the [summary strip](Workbooks-The-summary-strip) shows on screen,
from the same calculation, so the document cannot disagree with the app.

**Then one section per schematic**, each starting on a **new page**:

- the schematic image at **full page width**, with a thin outline,
- every worklog's marked area **washed and outlined in its category colour**,
- worklogs with "Show marked area" off getting **no rectangle**, their pills **parked top-right** —
  exactly as all three on-screen surfaces draw them,
- then each worklog, in id order.

**Each worklog prints** its `#N` badge, title, description, category chip and status pill, its work
done with totals, its comments, its links and its photos.

Worklogs whose schematic could not be resolved are filed under **"(no schematic)"** rather than
being silently dropped.

### Details worth knowing

**The pills look like the app's own** — rounded status pills, softened category chips, and a `#N`
badge filled in the category colour with a white disc holding the state padlock. A **Closed Issue**
prints as a green padlock on a red badge, the same two-channel colouring the board uses.

**Each photo sits in its own bordered panel** with its **file name** and its comment. The border is
what makes the grouping readable — with two photos side by side, the gap between a picture and its
own caption is otherwise the same as the gap to the next one's. The file name is printed even when
there is no comment, so a recipient can find that exact photo in the ZIP.

**Web links are real, visible hyperlinks** — blue and underlined. A PDF viewer gives no hover cue of
its own, so an unstyled hyperlink is indistinguishable from prose until someone happens to click it,
and on a printed page the styling is the only cue that survives.

Which runs of your free text count as links is decided the same way as on screen, so the document
never linkifies something the app does not. **Link rows** are different — those are declared
destinations, so they are always linked, with `https` filled in when you left the scheme off.

---

## What is in the ZIP

```
Workbook_3_Commodore_C64_20260904.zip
├── Workbook_3_Commodore_C64_20260904.pdf
├── worklog_1/
│   ├── 5v-rail-ripple.png
│   └── 7805-datasheet.pdf
├── worklog_2/
│   ├── vic-socket-before.jpg
│   └── vic-socket-after.jpg
└── worklog_4/
    └── invoice.pdf
```

Folders are named **`worklog_{id}`** — the same name those attachments have in your own Workbooks
folder, so what the recipient unpacks matches what you see on your own disk.

The worklog's title is deliberately not in the folder name: it is free text often carrying customer
details, and the PDF beside it already says which worklog is which.

Two worklogs can both hold a `front.jpg`; the second gets a numeric suffix rather than overwriting
the first.

---

## Notes

**The export is not opened afterwards.** The app only opens local files from inside its own data
root, and an export is saved wherever you chose.

**Export runs off the UI thread**, so a large workbook with many photos does not freeze the window.

**If it fails, you are told.** An export that half-worked must not look like it succeeded.

---

**Next:** [Where your data is stored](Workbooks-Where-your-data-is-stored) · [The summary strip](Workbooks-The-summary-strip)
