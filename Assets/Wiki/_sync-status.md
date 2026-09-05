# Wiki sync status

What has been copied from this folder into the live Wiki, and when. See
[`README.md`](./README.md) for how the mirror works.

**How to read this:** a page is in sync when the file here has not changed since the commit named
below. Ask for a sync diff and the pages that have drifted since their last-synced commit will be
listed, along with what changed in each.

**How this is maintained:** the table is updated when the maintainer confirms a page has been
pasted into the Wiki — not when the file changes. A row saying "in sync" means the Wiki genuinely
held that content at that commit.

Last reviewed: **2026-09-05**

| Page | Last synced | Status |
| --- | --- | --- |
| `Home` | `e60df599` | In sync |
| `Add-new-board-with-KiCad-data` | `e60df599` | In sync |
| `Board-Excel` | — | **Never synced** — carries a link fix (see below) |
| `Board-JSON` | — | **Never synced** — carries a link fix (see below) |
| `Commandline-parameters` | `e60df599` | In sync |
| `Compiling-yourself-from-source` | `e60df599` | In sync |
| `Contribute-data-via-CRT` | `e60df599` | In sync |
| `Contribute-data-via-GitHub` | `e60df599` | In sync |
| `Controlling-oscilloscope-with-keyboard` | `e60df599` | In sync |
| `Development-tools-used` | `e60df599` | In sync |
| `Explanation-of-data-files` | `e60df599` | In sync |
| `Information-collected` | `e60df599` | In sync |
| `Inspiration-for-creating-application` | `e60df599` | In sync |
| `Installing-application-in-Linux` | `e60df599` | In sync |
| `KiCad-folder` | — | **Never synced** — carries a link fix (see below) |
| `Main-Excel` | — | **Never synced** — carries a link fix (see below) |
| `MiniPro-programmer` | `e60df599` | In sync |
| `Scope-baseline-folder` | — | **Never synced** — carries a link fix (see below) |
| `Synchronize-oscilloscope` | `e60df599` | In sync |
| `Workbooks` | — | **Never synced** — replaces the 4-word stub with a real index |
| `Workbooks-Getting-started` | — | **Never synced** — new page |
| `Workbooks-Concepts-and-vocabulary` | — | **Never synced** — new page |
| `Workbooks-The-worklog-bar` | — | **Never synced** — new page |
| `Workbooks-The-Workbooks-tab` | — | **Never synced** — new page |
| `Workbooks-The-worklog-editor` | — | **Never synced** — new page |
| `Workbooks-Marking-areas-on-a-schematic` | — | **Never synced** — new page |
| `Workbooks-Components-in-scope` | — | **Never synced** — new page |
| `Workbooks-Searching` | — | **Never synced** — new page |
| `Workbooks-The-summary-strip` | — | **Never synced** — new page |
| `Workbooks-Attaching-oscilloscope-captures` | — | **Never synced** — new page |
| `Workbooks-Exporting-a-workbook` | — | **Never synced** — new page |
| `Workbooks-Where-your-data-is-stored` | — | **Never synced** — new page |
| `Workbooks-Deleting-things` | — | **Never synced** — new page |
| `Workbooks-Troubleshooting-and-FAQ` | — | **Never synced** — new page |

## Outstanding — 20 pages to paste

**15 Workbooks pages are entirely new.** `Workbooks` currently reads "Will come... soon" in the
Wiki; it becomes an index, and the other 14 are pages that do not exist there yet. They must be
created by hand in the Wiki (New Page, named exactly as the file) before the index's links resolve.

The other **five carry a fix** that is not yet in the live Wiki. Seven links pointed at
`wiki/Explanation-data-files`, a page that has never existed under that name (the real one is
`Explanation-of-data-files`), so they 404 in the Wiki today. Each is now corrected here.

| Page | Links fixed |
| --- | --- |
| `Board-Excel` | 2 |
| `Main-Excel` | 2 |
| `Board-JSON` | 1 |
| `KiCad-folder` | 1 |
| `Scope-baseline-folder` | 1 |

## Known content issues, not yet addressed

Not blocking a sync — recorded so they are not rediscovered each time:

- **`KiCad-folder` and `Scope-baseline-folder`** both read "Must be written &lt;SIGH&gt;".
- **`Board-Excel`** has a table-of-contents entry, `#technical-name-or-value`, pointing at a
  heading that no longer exists.
- **`Compiling-yourself-from-source`** lists `Command-line switches` in its contents, but the page
  has no such section.
- **`Development-tools-used`** cites specific AI models "as of March-2026" and "as of August-2026",
  which will date.
