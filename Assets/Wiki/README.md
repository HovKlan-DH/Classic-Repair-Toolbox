# Wiki mirror

This folder is the **source of truth for the GitHub Wiki**. Every `.md` file here corresponds to
one Wiki page of the **same name**, and its content is what that page should contain.

> [!IMPORTANT]
> The Wiki itself is updated **by hand**. Nothing here publishes automatically — GitHub gives no
> way to push a folder in this repository to the Wiki. Editing a file here changes what the page
> *should* say; the page only changes when the maintainer copies it across.

The published documentation lives at
[the Wiki](https://github.com/HovKlan-DH/Classic-Repair-Toolbox/wiki).

## Why this folder exists

The Wiki is a separate git repository with no pull requests, no review and no link to the code it
describes. Keeping the text here means a documentation change travels with the commit that made it
necessary, and can be reviewed like any other change — while the Wiki keeps the integrated,
easy-to-read presentation the maintainer wants for readers.

The cost of that split is one manual step, described below.

## Filenames ARE page names

A Wiki page's name comes from its filename, so the two must match **exactly**, capitals included:

| This file | Wiki page |
| --- | --- |
| `Board-Excel.md` | `.../wiki/Board-Excel` |
| `MiniPro-programmer.md` | `.../wiki/MiniPro-programmer` |
| `Home.md` | the Wiki's front page |

**Do not rename a file here** unless the Wiki page is renamed to match in the same sitting.
Renaming one without the other silently breaks every link that points at it, including the
in-app help buttons (see below).

## Links and images use WIKI syntax, not repository syntax

These files are written to be **pasted into the Wiki**, so they follow the Wiki's own conventions
rather than this repository's:

- **Internal links have no path and no extension**: `[Board JSON](Board-JSON)`, never
  `[Board JSON](./board-json.md)` and never a full `https://github.com/.../wiki/...` URL. A relative
  path breaks the moment it is pasted into a page, and an absolute one is 65 characters of noise
  that breaks if the repository ever moves.
- **Images stay as `<img>` tags pointing at their `user-attachments` URL**, exactly as the Wiki
  holds them. Text pasted into a Wiki page cannot carry an image file with it, so the URL must
  survive the paste untouched.

`images/` holds reference copies of those images so they are versioned and cannot be lost if the
upload URL ever disappears. **The Wiki does not read from this folder** — it is a backup, not a
source. Repointing a page at it would break the image.

## Page conventions

Every page opens the same way, so a reader knows where they are:

```markdown
[Wiki Home](Home) · [Parent page](Parent-Page)

One or two lines: what this page gets you.

---
```

- **No `#` H1.** The Wiki renders the page name as the title already, so an H1 in the body just
  says it twice. Sections start at `##`.
- **The breadcrumb comes first, then the summary**, then a `---` rule.
- **No "Go to Wiki Home" footers.** `_Sidebar.md` is rendered by GitHub on every page and does that
  job; a hand-maintained footer rots (it was missing from 9 of 24 pages before).
- **The app is `CRT`** — not `_CRT_`, not "the application", not "the app".
- **Bold for anything clickable** (**Add worklog**), **backticks for literal text** you type or a
  file, folder or column name (`Board label`).
- **Tabs are quoted**: the "Configuration" tab.
- **A task page ends with an "If it does not work" table.** Symptom on the left, likely cause on the
  right. This is the most-used thing on the pages that have it.
- **`> [!WARNING]` only for "you will lose work or it will silently not work"**, `> [!NOTE]` for
  everything else. Nothing else.

## Updating a page

1. Edit the `.md` file here, in the same commit as whatever code change made it necessary.
2. When ready to publish, ask for the pages that have drifted — see
   [`!sync-status.md`](./%21sync-status.md) for what has been copied across and when.
3. Open the named Wiki page, select all, paste the file's contents, save.
4. Confirm, so [`!sync-status.md`](./%21sync-status.md) can be brought up to date.

## Pages deliberately NOT mirrored here

Six Wiki pages are not in this folder. They are orphans — nothing links to them, and each is
superseded by a page that is here:

| Wiki page | Superseded by |
| --- | --- |
| `Documentation` | the individual data-file pages |
| `Data-files` | `Explanation-of-data-files` and its children |
| `Compiling-yourself` | `Compiling-yourself-from-source` |
| `Classic-Repair-Toolbox-documentation` | `Home` |
| `Contribute-data` | `Contribute-data-via-GitHub` |
| `MiniPro-programmer-how‐to` | `MiniPro-programmer` |

They still exist in the Wiki and are untouched. If any should be kept alive, mirror it here first
so it stops being edited in two places.

## The five in-app help links

These Wiki pages are opened by buttons in the shipped application:

| Page | Opened from |
| --- | --- |
| `Workbooks` | Configuration tab, "?" beside "Enable Workbooks tab" |
| `MiniPro-programmer` | Configuration tab "?", and the component popup |
| `Controlling-oscilloscope-with-keyboard` | Component popup |
| `Synchronize-oscilloscope` | Component popup |

Renaming or deleting one of those pages breaks a button in a released build, which no update can
fix for versions already installed. Grep for the page name before touching it.
