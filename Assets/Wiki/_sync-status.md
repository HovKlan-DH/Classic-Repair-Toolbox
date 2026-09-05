# Wiki sync status

What has been copied from this folder into the live Wiki, and when. See
[`README.md`](./README.md) for how the mirror works.

**How to read this:** a page is in sync when the file here has not changed since the commit named
below. Ask for a sync diff and the pages that have drifted will be listed.

**How this is maintained:** the table is updated when the maintainer confirms a page has been
pasted — not when the file changes. A row saying "in sync" means the Wiki genuinely held that
content at that commit.

Last reviewed: **2026-09-05**

## Pages to update — 9

| Page | What changed | In Wiki? |
| --- | --- | --- |
| `Commandline-parameters` | **Rewritten.** Was missing `--workbooks-root` entirely, said "two parameters", and both `--simulate-update` examples had four dashes (`----simulate-update`). Also documents the bare `--simulate-update` form and its `99.0.0` default | Exists |
| `Development-tools-used` | **Corrected.** Listed Newtonsoft.Json, which the project does not use; QuestPDF was missing. Dated model lists removed | Exists |
| `KiCad-folder` | **Written.** Was "Must be written &lt;SIGH&gt;" | Exists |
| `Scope-baseline-folder` | **Written.** Was "Must be written &lt;SIGH&gt;" | Exists |
| `Board-Excel` | 2 links fixed (see below) | Exists |
| `Main-Excel` | 2 links fixed (see below) | Exists |
| `Board-JSON` | 1 link fixed (see below) | Exists |
| `Workbooks` | **Written.** Was "Will come... soon" — now an index | Exists |

## Pages to CREATE — 4

These do not exist in the Wiki. Create each with **New Page**, named exactly as shown, or the
`Workbooks` index links will not resolve.

| Page |
| --- |
| `Workbooks-Getting-started` |
| `Workbooks-Daily-use` |
| `Workbooks-The-Workbooks-tab` |
| `Workbooks-Export-and-data` |

## Already in sync — 13

No action needed: `Home`, `Add-new-board-with-KiCad-data`, `Compiling-yourself-from-source`,
`Contribute-data-via-CRT`, `Contribute-data-via-GitHub`, `Controlling-oscilloscope-with-keyboard`,
`Explanation-of-data-files`, `Information-collected`, `Inspiration-for-creating-application`,
`Installing-application-in-Linux`, `MiniPro-programmer`, `Synchronize-oscilloscope`.

## The link fix

Seven links pointed at `wiki/Explanation-data-files`, a page that has never existed under that name
(the real one is `Explanation-of-data-files`), so they 404 in the Wiki today. Corrected in
`Board-Excel` (2), `Main-Excel` (2), `Board-JSON` (1), `KiCad-folder` (1) and
`Scope-baseline-folder` (1).

## Known content issues, not yet addressed

Not blocking a sync — recorded so they are not rediscovered each time:

- **`Board-Excel`** has a table-of-contents entry, `#technical-name-or-value`, pointing at a
  heading that no longer exists.
- **`Compiling-yourself-from-source`** lists `Command-line switches` in its contents, but the page
  has no such section.
