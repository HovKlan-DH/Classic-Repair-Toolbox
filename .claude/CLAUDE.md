# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

Classic Repair Toolbox (CRT) is a cross-platform (Windows/Linux/macOS) Avalonia desktop app that helps
hardware enthusiasts diagnose and repair vintage computers. It presents schematics, component data,
oscilloscope baselines, and interactive KiCad traces for a curated set of hardware (mostly Commodore,
plus Amstrad and ZX Spectrum boards). It also drives real test equipment: SCPI oscilloscopes over TCP
and a MiniPro USB IC programmer/tester.

## Hands off CHANGELOG.md

**Never create, edit, rewrite, reformat or delete [CHANGELOG.md](../CHANGELOG.md) unless the
maintainer explicitly asks for it in that message.** It is written by hand, in the maintainer's own
words, and it is the body of every GitHub Release — an "improvement" there is not a small edit, it
is words the maintainer never wrote going out under their name.

This holds even when a change would normally warrant a changelog entry, and even when the file
already has uncommitted edits in it (those are the maintainer's, in progress). Do not touch it as a
"finishing touch" on a feature, do not tidy its formatting, and do not stage, commit, revert or
`git checkout` it. If you think an entry is needed, say so in your summary and let the maintainer
write it. "Update the changelog" from the maintainer is the only permission — and it covers that
one request, not the rest of the session.

## Documentation lives in `docs/`

The project's documentation is [docs/](../docs/) in this repository — plain Markdown, no static
site generator, no build step, no workflow. It replaced the GitHub Wiki, which is kept only as a
read-only archive; **do not add wiki links or wiki pages**, and repoint any you find.

- **Filenames are lowercase kebab-case**: `board-excel.md`, never `Board-Excel.md`.
- **Every page opens with a single `#` H1** carrying the human-readable page title.
- **Links between pages are relative and include the extension**: `[Board Excel](./board-excel.md)`.
  That form resolves in the GitHub file browser, in an editor and in a local clone alike.
- **Images live in [docs/images/](../docs/images/)** and are referenced relatively:
  `![Main window](./images/main-window.png)`. Never hotlink a `user-attachments` URL — those are
  wiki/issue upload artefacts and are not versioned with the repo.
- **[docs/README.md](../docs/README.md) is the index** GitHub renders for the folder. Every
  released page must be linked from it, grouped into sections.
- **Documentation ships in the SAME commit or PR as the code it describes.** A feature and the
  page describing it are one change, not two.

**Documentation for UNRELEASED functionality goes in [docs/upcoming/](../docs/upcoming/)**, with a
callout naming the target version:

```markdown
> [!NOTE]
> **Unreleased.** Planned for v2.6.
> This page describes functionality that is not available in the current release.
```

That folder is for **whole new topics only**. An addition to an already-released page is marked
inline with the same callout instead, so related information stays together rather than being
split across two files. Nothing in `docs/upcoming/` may be linked from `docs/README.md` as though
it were released — only from its clearly labelled "Upcoming" section. The promotion steps for when
a feature ships are in [docs/upcoming/README.md](../docs/upcoming/README.md); follow them there
rather than inventing a variation.

**The five in-app help links point into `docs/`** (`TabConfiguration.axaml.cs` and
`ComponentInfoWindow.axaml.cs`, as `blob/main/docs/<page>.md` URLs). Renaming or moving one of
those pages breaks a button in the shipped app, so grep for the filename before you rename it.

## Build, test, run, publish

There is a unit test suite covering the UI-free logic in `Handlers/`. **Always run it, always add
tests for new logic, and always update the existing tests when you change covered behaviour** — see
[Tests](#tests) below for the full rules. UI behaviour has no automated coverage, so changes to the
tabs and overlays still need the app built and run by hand (see [BUILDING.md](../BUILDING.md) for full
per-OS instructions).

- **Run the tests: `dotnet test Classic-Repair-Toolbox.slnx`** (~2s; needs no hardware and no display)
- Build (Release, matches CI): `dotnet build Classic-Repair-Toolbox.slnx -c Release`
- Build (Debug, default): VS Code task `build`, or `dotnet build Classic-Repair-Toolbox.slnx`
- Run/iterate: VS Code task `watch` (`dotnet watch run --project Classic-Repair-Toolbox.slnx`), or F5 in
  VS Code (`.vscode/launch.json`), or open `Classic-Repair-Toolbox.slnx` in Visual Studio
- Self-contained publish for a specific OS: `dotnet publish -c Release -f net10.0 -r <rid> --self-contained`
  where `<rid>` is one of `win-x64`, `linux-x64`, `osx-x64`, `osx-arm64`
- **The build configuration does not change what the app does.** A DEBUG build and a RELEASE build
  given the same arguments behave identically — same update check, same data sync, same diagnostics.
  Nothing may be gated on `#if DEBUG`; the one remaining use is `AppConfig.IsDebugBuild`, which is
  reported in the log and read by nothing else. RELEASE still matters for *timings* (DEBUG is
  JIT-only) and for warnings-as-errors, not for behaviour.
- **Command-line switches**, all parsed once at startup:
  - `--data-root=<path>` — use a different `Data` folder (default: an AppData folder that survives
    Velopack updates). See `DataManager.ResolveDataRoot`.
  - `--workbooks-root=<path>` — use a different folder for worklog workbooks (repair jobs), the same
    idea as `--data-root=` but for the purely-local "Workbooks" folder (default: also under AppData).
    See `WorklogManager.ResolveExplicitWorkbookRoot`.
  - `--simulate-update[=<version>]` — offer a fake application update (default `99.0.0`), fake the
    download and skip the restart, so the update banner can be exercised without a release. See
    [Handlers/Data/SimulationOptions.cs](../Handlers/Data/SimulationOptions.cs). Active in RELEASE
    builds too, on purpose; the startup log shouts about it and the banner says `(simulated)`.
  - All three are parsed the same way (case-insensitive, surrounding quotes stripped, first match
    wins, unrecognised arguments ignored). `--simulate-update` is set for F5 and the `watch` task in
    [.vscode/launch.json](../.vscode/launch.json) and [.vscode/tasks.json](../.vscode/tasks.json).
- **To skip the online data sync while iterating, untick "Check for new or updated data at
  application launch" in the Configuration tab.** It is a normal user setting, not a build or
  command-line concern — it short-circuits the manifest fetch, and with no manifest the board-Excel
  sync and the background image sync skip themselves too.

## Tests

`Tests/Classic-Repair-Toolbox.Tests/` (xUnit) covers the UI-free logic in `Handlers/`. It needs no
oscilloscope, no MiniPro programmer and no display, and runs in about two seconds.

**Tests are part of the change, not a follow-up. These rules are not optional:**

1. **ALWAYS write tests for logic you add.** Any new function that parses, maps, validates,
   compares, or does geometry/maths arrives *with* its tests in the same change — never "tests can
   come later". If it is pure, it goes in `Handlers/` (see `Handlers/Geometry/` for logic pulled out
   of a tab) and it gets a test file. Do not ask whether to add tests; add them.
2. **ALWAYS update the existing tests when you change covered behaviour.** If you touch a class that
   has tests, updating them is part of that same change. Never leave a red suite, never defer the
   update, and never delete a test to make a change pass.
3. **ALWAYS run `dotnet test Classic-Repair-Toolbox.slnx` before reporting any code change as done.**
   "It compiles" is not a completed change. Report the real result, including failures and counts.
   **This rule is machine-enforced, not a courtesy.** The `Stop` hook in [settings.json](settings.json)
   runs [hooks/require-green-tests.sh](hooks/require-green-tests.sh) when you try to end a turn, and a
   red suite blocks the handover with the failing tests fed back to you. It builds Release (matching
   CI) and only fires when `.cs`, `.csproj`, `.axaml` or `.slnx` files actually changed, so it is free
   on turns that touch no code. GitHub then runs the suite again on every push
   ([.github/workflows/build-and-unittest.yml](../.github/workflows/build-and-unittest.yml)), and a
   red suite blocks releases too.
4. **A failing test is a question, not an obstacle.** Decide whether the behaviour change was intended.
   If it was, update the expectation and say so explicitly in your summary. If it wasn't, fix the code.
   Never edit an assertion just to get to green, and never weaken one (e.g. loosening a tolerance or
   dropping a case) to avoid understanding a failure.
5. **When you fix a bug, first add the test that fails because of it.** Then fix the code and show
   the test going green. That test is the thing that stops the bug coming back.
6. **Never write a test that needs hardware, a network call, a display, or that starts a process.**
   Headless UI tests do not breach this - Avalonia's headless platform needs no display - but they
   must go through `UiTest.Run(...)`; see [Headless UI tests](#headless-ui-tests).
   `ExternalTargetLauncher`'s accept path calls `Process.Start`, so its tests exercise the private
   containment predicates by reflection instead; the header comment in `ExternalTargetLauncherTests.cs`
   explains the reasoning. Follow the same principle for anything else with real-world side effects.
   If logic is untestable because it is welded to a control, extract the logic (rule 1) rather than
   giving up on testing it.
7. **These are characterisation tests** — they pin down what the code does today so future changes
   cannot alter it silently. Several encode deliberate quirks (relative-tolerance value matching, "no
   vector grid + a successful summary IS a pass", micro sign vs Greek mu, the dead `region` argument
   in `BoardDataWriter`). Read the comment before assuming a test is wrong.

**Writing them:** one test file per class under test, named `<ClassName>Tests.cs`. Give each test a
sentence-shaped name saying what must hold (`A_faulty_chip_fails_and_names_the_failing_pin`), and put
a comment above anything non-obvious explaining *why* it matters — a reader should learn the rule from
the test. Use `TempWorkspace` for filesystem work. Cover the failure and edge cases, not just the
happy path: blank input, malformed contributed data, wrong region, wrong board side, locale-specific
number formats.

### What is covered

| Area | Classes |
| --- | --- |
| Oscilloscope | `ScopeValueMapper`, `ScopeCommandResolver`, `ScopeCommandPaletteDefinitions`, `ScopeFormatting`, `ScopePayloadParser`, plus `TabOscilloscope`'s command sequencing via `IScopeClient` and a fake |
| IC testing | `MiniproOutputParser`, `IcTestService` (via `MockMiniproRunner` and local test doubles) |
| Security | `ExternalTargetLauncher`, `OnlineServices`' manifest-validation predicates |
| KiCad | `KiCadRawProjectLoader`, `KiCadProjectLoader`, the `KiCadProjectData` model |
| Board data | `BoardDataReader`, `BoardDataWriter`, `BoardComponentHighlightStorage`, `ComponentListBuilder`, `ComponentImageQueries`, `OverviewHtmlBuilder`, `ContactLinkFormatter` |
| Worklog | `WorklogManager` (including `ResolveActiveWorkbook`, `AddEntryRecord`, `DeleteEntry`, the non-reusing id counters, `IsResolvedState`, `IsWorkbookStatusOpen`, `GetAllWorkbooks`), `WorklogEntryScope`, `WorklogSearchQuery`, `WorklogSearchIndex`, `WorkbookSummary`, `WorkbookExportModel`, `WorkbookPdfExporter.WriteZip` (the archive only), `WorklogAttachTargets`, `WorklogAttachmentWriter` |
| Text links | `TextLinkFinder` (which runs in a user-typed note are web links) |
| Settings / startup | `UserSettings`, `DataManager` (data-root + master workbook), `DataValidator` (smoke only), `SimulationOptions` |
| Headless UI (`Tests/.../Ui/`) | All nine tabs built headlessly, the worklog and Workbooks palettes, component highlight selection and schematics zoom, plus `Main` itself, the label editor's full edit cycle, the worklog area-marking flow, `ComponentInfoWindow`, the oscilloscope's SCPI sequencing, and the Configuration/Overview/About tabs - see [Headless UI tests](#headless-ui-tests) |
| Geometry (`Handlers/Geometry/`) | `PolygonGeometry`, `RectGeometry`, `KiCadLayerGeometry`, `KiCadPadGeometry`, `OverlayCullGeometry`, `KiCadOverlayCacheKeys`, `KiCadOverlayNetCache`, `ViewportMath`, `KiCadNetGraphBuilder`, `KiCadHoverIndex`, `HighlightRectBuilder`, `LabelEditorGeometry`, `LabelEditorSnapGeometry`, `TraceGeometry`, `KiCadCalibrationGeometry`, `WorklogBadgeLayout`, `ExportOverlayGeometry`, `WorklogDefaultAreaGeometry` |

`Handlers/` is where the real coverage is; most of the uncovered remainder is `Tabs/` and `Main/`,
Avalonia code-behind that is verified by running the app.

**No coverage percentage is recorded here on purpose** - a figure in a document goes stale silently
and then gets quoted as fact. Measure it when you actually need it:

```
dotnet test Classic-Repair-Toolbox.slnx --collect:"XPlat Code Coverage"
```

and read `lines-covered` / `lines-valid` from the `coverage.cobertura.xml` it writes under
`Tests/.../TestResults/`. **Always quote the denominator alongside the percentage, and say which
build configuration you used** - Debug and Release instrument different numbers of lines (~26.7k vs
~21.4k), so two bare percentages from different configurations are not comparable.

**CI already computes it on every push**, which is the one place a figure cannot go stale:
[build-and-unittest.yml](../.github/workflows/build-and-unittest.yml) collects coverage alongside
the test run and renders the totals into the run's GitHub job summary via
[coverage-summary.sh](../.github/workflows/coverage-summary.sh), with the raw Cobertura XML kept as
a 7-day artifact for per-file numbers. It is **reported, never enforced** - there is deliberately no
threshold that fails the build, since a floor set while the suite is growing either blocks unrelated
work or is meaningless. Read the number from a recent run rather than writing it down anywhere.

### `Handlers/Geometry/` — pure logic pulled out of the UI

This folder exists because ~2,000 lines of genuinely pure logic were trapped as `private`
members of `TabSchematics`, where no test could reach them: polygon and zone maths, KiCad layer
filtering, the 550-line net-graph builder that decides which copper belongs to which net, the
spatial hover index, highlight rect building and label-editor handle geometry.

**When you add pure logic to a tab, put it here instead.** These classes use Avalonia's
`Point`/`Rect`/`Matrix` value types but never touch a control, so they test with no display.
`KiCadRenderNodes.cs` holds the DTOs they share (formerly nested private types).

**`public` and `internal` are both fine here, and the folder deliberately uses both.** A class
extracted from a tab that nothing outside the assembly needs (`LabelEditorGeometry`,
`LabelEditorSnapGeometry`, `TraceGeometry`, `KiCadCalibrationGeometry`, `KiCadNetGraphBuilder`, `KiCadHoverIndex`, and the
`KiCadRenderNodes.cs` DTOs) stays `internal`; the tests reach it through the
`InternalsVisibleTo` entry in [Classic-Repair-Toolbox.csproj](../Classic-Repair-Toolbox.csproj), so
`internal` costs no coverage. Do not widen one to `public` for consistency's sake — a type is
`public` here only if something genuinely consumes it from outside.

The same extraction has since been done for the other tabs, into the area folder that fits rather
than into `Geometry/`: `Handlers/Oscilloscope/ScopeFormatting.cs` and `ScopePayloadParser.cs` (from
`TabOscilloscope`), and `Handlers/Data/ComponentListBuilder.cs`, `ComponentImageQueries.cs`,
`OverviewHtmlBuilder.cs` + `OverviewModels.cs`, `ContactLinkFormatter.cs` (from `Main`,
`ComponentInfoWindow`, `TabOverview` and `TabAbout`). **Match that pattern: geometry goes in
`Geometry/`, everything else goes in the `Handlers/` folder for its area.**

That sweep also removed four copies of logic that already existed in `Handlers/`
(`ComputeWheelZoomFactor`, `GetImageContentRect`, `PixelToLocalRect`, and a second
oscilloscope-title-suffix stripper). Before writing a helper in a tab, grep `Handlers/` for it —
these duplicates all arose from someone re-implementing a helper they could not see.

### Test seams — use these, do not work around them

Three classes are static singletons that would otherwise read and write the user's real files. Each
has an `internal` seam; the test project sees them via `InternalsVisibleTo` in the csproj.

- `UserSettings.LoadFrom(path)` — `Load()` resolves the AppData path and delegates here. Tests point
  it at a temp file. **Never call `UserSettings.Load()` from a test.**
- `DataManager.LoadFrom(dataRoot, workbookName)` — the local half of `InitializeAsync`, no network and
  no seeding. **Never call `DataManager.InitializeAsync()` from a test.**
- `Logger` writes nothing until `Logger.Initialize()` is called, and no test calls it — which is why
  the `Logger.*` calls inside classes under test are inert. **Never call `Logger.Initialize()` from a
  test**, or the suite starts writing to the user's real log file.
- `WorklogManager.LoadFrom(root)` — the same idea for the local "Workbooks" folder. **Never call
  `WorklogManager.Load()` from a test.** `TabWorkbooks.BoardKeyOverrideForTests` is the matching
  seam on the UI side: the Workbooks tab normally reads its board key off `Main`, which no test
  constructs, so the override lets the list be tested without standing up the main window.

Use `TempWorkspace` for anything that touches the filesystem; it creates and deletes a temp folder.
Tests that mutate `UserSettings` or `DataManager` static state live in the `"UserSettings"` and
`"DataManager"` xUnit collections so they run sequentially. `BoardDataReader` has the same
problem — its loaded boards sit in a shared static cache (a `ConcurrentDictionary`, so thread-safe,
but still one cache that tests clear and repopulate) — so `BoardDataReaderTests`
and `BoardDataWriterTests` share the `"BoardData"` collection.
**Any new test class that touches one of these statics must join its collection**, or it will
pass alone and fail intermittently in the full run.

**Collections do not run in parallel with each other.** `Tests/.../xunit.runner.json` sets
`"parallelizeTestCollections": false`. A class can only join ONE collection, and the headless UI
tests must be in `"HeadlessUi"` to share the dispatcher thread - which left `WorkbooksListTests`
racing the `"UserSettings"` collection over `UserSettings`' shared `_data` object and save path, with
unique dictionary keys giving no protection at all. Serialising every collection costs about five
seconds on a ~40s suite and removes the whole class of race. Keep that file.

### Headless UI tests

`Tests/Classic-Repair-Toolbox.Tests/Ui/` builds every tab through Avalonia's headless platform -
no display, no GPU, so it runs on CI like any other test. Two of the files are the harness:
`TestAppBuilder.cs` (a `CRT.App` subclass whose `OnFrameworkInitializationCompleted` is deliberately
empty, since the real one calls `Logger.Initialize()`, shows a splash and syncs over the network)
and `UiTest.cs` (runs a body on the UI thread). The rest are the tests themselves:

| File | Covers |
| --- | --- |
| `TabConstructionTests.cs` | Every tab constructs without throwing |
| `ConfigurationHelpIconTests.cs` | The Configuration tab's "?" help icons (Workbooks and MiniPro): that each button exists, carries the `HelpIconButton` class and the Font Awesome circle-question glyph, and shares a row with the checkbox it explains. The CLICK is deliberately not tested - it goes through `ExternalTargetLauncher`, whose accept path calls `Process.Start` (rule 6); a mis-typed `Click` handler name already fails the XAML parse |
| `ComponentHighlightSelectionTests.cs` | Selecting/deselecting in the component filter box, and the highlights that appear and vanish across the main image and every thumbnail |
| `SchematicsZoomTests.cs` | The schematic viewer's zoom limits, and zoom anchoring - that the point under the mouse pointer stays under the mouse pointer |
| `WorkbooksPaletteTests.cs` | The Workbooks tab's `Workbooks_*` theme keys: that each resolves, that both themes define all of them, and that the pin colours stay identical across themes |
| `WorkbooksListTests.cs` | The Workbooks tab's workbook list and selection: counts and their singular/plural forms, card contents, newest-first order, board scoping, default/click selection, that the status pill (list and top-line) uses the shared Open/Closed brushes, the activation fallback chain (`UserSettings.ActiveWorkbookIdByBoard` wins over newest; a stale saved id falls back to newest), that both splitters' widths are restored from `UserSettings.WorkbooksLeftPanelWidth`/`WorkbooksEntryListWidth` via `ApplySplitterWidthsForTests`, the top-line's Note text (shown/switched/collapsed-when-blank) and its Edit/Delete actions' visibility (hidden with no workbook selected, shown once one is), and that deleting a workbook via `WorklogManager.DeleteWorkbook` removes its card and a refresh lands the selection on the board's next remaining workbook (or clears the top-line entirely when it was the last one) |
| `WorkbooksBoardPreviewTests.cs` | The Workbooks tab's board pane: one preview per schematic with an entry in the selected workbook, shared previews for co-located entries, unknown-schematic and missing-image entries skipped without dropping other previews, the pane rebuilding when the selected workbook changes, (via `BuildShownTab`'s real layout pass) that a "show marked area" ON entry's badge anchors to its marker while an OFF entry's badge parks in the image's top-right corner instead, that a pill carries a Hand cursor on a hit-test-visible canvas, `RefreshBoardPreviewsForCurrentSelection` populating the pane from board data supplied after an earlier empty build without touching the workbook list, schematic selection (default-first, highlight moving on click, the entry list switching to the clicked schematic), an entry detail card's four rows read back by content (including the stats row's hours/cost/comment/link/photo/file counts, populated through `WorklogManager.UpdateEntry` rather than a bare in-memory record), the two empty-state messages on a board with no workbooks (both reading "No worklogs recorded yet for any schematics in this board", and both `VerticalAlignment.Top` - a `TextBlock` in a Grid cell defaults to Stretch, which floats a single line down the middle of an empty panel), that the card has exactly one outer border with no border wrapping any individual row, the per-card "Delete worklog" button (one on EVERY card; in the title row Grid's Auto column and Top-aligned, so it hugs the top-right corner past a title that wraps; carrying the `Button_Cancel_*` theme brushes - asserted against the KEYS, since the header button's own `DynamicResource` values have not resolved on a tab that is never attached to a window - and NOT the card's Hand cursor, compared by `Cursor.ToString()` because `Cursor` has no value equality and a fresh `new Cursor(Hand)` fails as "Expected: Hand / Actual: Hand"), that the Legend panel is gone, and `BuildWorklogEntryComponentScopeForTests` (matched components whose highlight rect the entry's area touches, `null` with no cached highlight rects at all and `null` when the entry's own schematic has no cache entry) - the computation `OnPreviewBadgePointerPressed` hands to `WorklogEntryEditorWindow.InitializeComponentScope` so the modal opened from a pill is provably the same one the Schematics tab opens |
| `WorkbooksSearchTests.cs` | The "Find a previous repair" box actually applying `WorklogSearchQuery` to the tab: the workbook list narrowing (by title, by note, and through text in one of the workbook's entries), the result count and each empty state saying "no match" rather than "none recorded", AND/quoted-phrase/`-`exclusion/case-insensitivity end to end, the entry list narrowing to matched entries while a workbook matched by its OWN text keeps all of its entries, that filtering the ACTIVE workbook out moves the top-line to a shown one WITHOUT changing `ActiveWorkbookIdByBoard`, that `ClearSearchForBoardChange` empties both the box and the filter, and the highlighting - the matched runs marked and nothing else, original casing and the full text preserved through the `Inlines` split, no marks with no search active, and marks removed again when the box is cleared |
| `WorkbooksBoardPreviewTests.cs` (cont.) | Plus the DETACH/RE-ATTACH cycle a tab switch performs: that detaching leaves no preview `Image` holding a `Source` (the disposed-bitmap crash - see the board pane's bitmap-cache note above; this test fails against the dispose-without-clearing version), that re-attaching rebuilds the pane with a LIVE bitmap rather than handing back the disposed one, and that detaching a tab which built no previews at all is harmless |
| `DeleteWorkbookWindowTests.cs` | That the delete-confirmation modal's Enter/Escape both CANCEL - including with the **Delete button focused**, the case a plain bubbling `KeyDown` handler misses entirely (the button's own Enter handling fires `Click` and confirms the delete). Asserts on the Delete button's `Click` rather than on the window closing, since it closes either way - the fix is `RoutingStrategies.Tunnel`, and this test fails against the bubbling version |
| `DeleteWorklogWindowTests.cs` | The per-entry delete confirmation, the twin of the above and for the same reason: that Enter CANCELS even with the **Delete button focused** (the Tunnel-vs-bubbling regression again, asserted on the Delete button's own `Click`), that Escape cancels, that the worklog is named by the same `#N · Title` its card and board pill show (an untitled one as `(untitled)`, not trailing off after the separator), and that the copy says WORKLOG throughout - a near-copy of the workbook dialog that still named the workbook would tell the user they are about to lose the whole job when they are deleting one line of it |
| `WorklogEditorNewEntryTests.cs` | The editor opened on a NEW entry (`InitializeForNewEntry`, the "Add worklog" flow after the quick card was removed): that it opens blank with Save disabled, that typing a title ENABLES Save - the thing `Initialize`'s own end-of-method clean state would otherwise make impossible, so a new entry could never be saved at all - that a blank or whitespace title disables it again, the drawn area's schematic carried through, the Note/Open defaults and the "Worklog created" audit comment, "Show marked area" ticked, the window titled "New worklog entry", that cancelling reports `WasSaved == false` and a null `SavedNewEntry` (a draft writes nothing, so the caller must not be told to refresh), and `InitializeComponentScope`'s `tickAll` - fully ticked for a new entry, unticked without it |
| `WorklogEditorNewEntryTests.cs` (cont.) | Also: the Save button reads "Add worklog" rather than "Update worklog" (set explicitly by `InitializeForNewEntry`, not left as the markup default), and that a blank title on a brand-new entry shows NO explanatory message - there is nothing on disk yet to disagree with, unlike a saved entry (see `WorklogEditorHeaderTests.cs`) |
| `WorkbooksSummaryAndPillsTests.cs` | The things that made this tab read as one surface: that the entry card's status pill and the top-line's have the SAME border width AND colour (the reported "the pill is not identical" — matching on only one of the two is what let them drift while each claimed to match), that a status pill's border is the STATE colour so Open and Closed differ and a category chip's is its own CATEGORY colour, both at 1px; that an entry card carries a Hand cursor marking it clickable (the click itself opens a modal a headless test cannot dismiss — what makes it provably the pill's modal is the shared `OpenEntryEditor`); that a counted pill in the summary keeps the ordinary 1px informational outline in its own colour but drops its ICON (while an uncounted one keeps it), and that the category/state counts are drawn as pills with the count LEADING each; and the summary strip's real totals (counting "worklogs", not "entries"), that only its NUMBERS are bold (asserted run by run - the finished string cannot show the difference), its collapsed-by-default state, toggling both ways, that an expanded strip SURVIVES a refresh (it rebuilds on every board change and save), the components line hidden when nothing is scoped, and the strip hidden entirely with no workbook selected. Plus that BOTH export formats have their own visible button carrying no icon (the ZIP was reported as invisible when it lived only in the save dialog's type list), and the export document built through the tab's own path — that the tab's board data reaches the exported sections, and that the suggested file name names the workbook and board but NOT the title |
| `WorkbooksSearchFocusTests.cs` | `TabWorkbooks.FocusSearchBox` - that calling it moves real keyboard focus onto `FindRepairTextBox`, that calling it twice is harmless, and that focus can still move away afterwards through ordinary interaction. Covers only what is testable without `Main` (never constructed by any test): `Main`'s own `OnMainTabControlSelectionChanged`, which calls this on tab entry and is guarded by `e.Source` against `SelectionChanged`'s bubble from a nested `ListBox`/`ComboBox` on another tab, is exercised only by running the app |
| `ThumbnailWorklogPillsTests.cs` | The thumbnail gallery's "#N" pills, via `ThumbnailWorklogPillsOverlay.LayOutPills`: that a "show marked area" ON entry's pill sits on its marked area while an OFF one is PARKED in the image's top-right corner instead - the reported bug where the thumbnail kept drawing a hidden entry's pill at its marker while the main view parked it, asserted by actual position (the marker is deliberately bottom-left) rather than by mere difference - that the same marker lands in two different places depending on the flag, that a thumbnail carrying both kinds keeps each placement, that parked pills stack without overlapping and stay inside the image however many there are, that they hug the IMAGE's edge and not the letterboxed control's, that each keeps its own id and colour, and that a zero-sized bitmap lays out nothing |
| `WorklogOverlayRefreshTests.cs` | That the Schematics tab's overlay and thumbnail pills REDRAW after an entry changes in the workbook already on screen - `RefreshWorklogEntriesListForCurrentWorkbook` re-reading from disk so an edited area moves, a ticked "Show marked area" makes the rectangle appear, an unticked one removes it, and an added entry shows up; plus that a refresh with "Show worklogs" off stays cheap and draws nothing (it is called on EVERY worklog change, including for users who never turn the overlay on). Four of the five fail against the pre-fix version, where `Main.RefreshWorklogBar` only reached the overlay inside its "the shown WORKBOOK changed" branch |
| `WorklogMarkedAreaDefaultTests.cs` | Ticking "Show marked area" on a worklog that never had one - the entry created from an oscilloscope capture, which is stored with no area and parks as a corner pill. That the tick leaves it with a rectangle that is genuinely visible, grabbable and wholly INSIDE the board (against the zero-sized rect that draws as nothing and can never be dragged into place), that the new square lands BOTTOM-right, away from the top-right parked pills, that an entry which already has a drawn area is never moved - asserted through a full untick/retick cycle, the sequence that would expose any re-placement - that unticking invents nothing, and that `SetShowMarkedAreaForNewEntry` can create a worklog parked rather than marked. Reads `WorkingEntryAreaForTests`, not the record passed in: `Initialize` clones it (see `CloneEntry`) so Cancel cannot mutate the caller's copy, and a test reading that copy would see nothing change no matter what the editor did |
| `WorklogAttachCaptureWindowTests.cs` | The modal that files a captured oscilloscope image into a worklog: that the PRESELECTED row is the ranked-first one (asserted with a component match that is NOT lowest by id, so a dialog doing no ranking at all fails it), that "Create new worklog" is always offered and is always LAST so it never displaces a real entry from the preselected slot (and is correctly the only row, and preselected, when the workbook has no entries), that the button reads "Attach to existing worklog" for an existing worklog and "Create worklog" for the new-entry row (it opens the full editor rather than attaching there and then, and a button still reading "Attach" would misdescribe that), that the target workbook is NAMED (this dialog opens from the component popup, which can be sitting over a schematic while the user has been looking at the scope), and the GROUP HEADERS - that both bands are named in the list itself ("Worklogs with U8 in scope" / "All other worklogs"), that every header is disabled so it can never be selected while the preselected row is still a real worklog, that a header is faint enough not to read as an option and is OUTDENTED with its worklogs indented under it (asserted past the Fluent theme's own 11px item padding, since a row at the default already sits right of the header - verified by removing the `ContainerPrepared` hook), that the matched heading picks the component out in BOLD inside brackets with only that run bold and the joined runs still reading "Worklogs with [U8] in scope", that the "All other worklogs" heading stays a plain string, and that no headers appear at all when nothing matches the component |
| `TextLinkRendererTests.cs` | Rendering a user-typed note with its web links clickable: that link-free text stays a plain single-`Text` block with no Hand cursor, that a linked one moves its content into `Inlines` with `Text == null` (a block carrying both renders the Text and silently ignores the Inlines), that only the link run is underlined, that re-rendering replaces the previous pass rather than layering on it, and the LINK + SEARCH-HIGHLIGHT merge - a search term landing inside a URL, one outside it, highlighting with no link present, and that the merged runs are never empty and always rebuild the original string. Plus the `LinkText` attached property the editor's DataTemplates use, including re-rendering when a recycled container is handed a different row |
| `WorklogEntryModeTests.cs` | The parked-pill canvas (separate from the anchored badge canvas, so parked pills do not pan and zoom with the board; no `Background`, since one would swallow every press across the schematic panel; below the "Netlist names" panel in z-order) and the "Add worklog" mode hint (its wording, that it starts hidden and is not hit-testable, that it covers the data-sync icon, that its text wraps inside its box rather than overflowing - a horizontal `StackPanel` measures with infinite width and would never wrap - and that it is plain text with no icon). Formerly `WorklogCreateCardTests.cs`; the quick card's own tests went with the card |

**Do NOT add the `Avalonia.Headless.XUnit` package to get `[AvaloniaFact]`.** At 12.1.1 it depends
on xunit **v3** while this suite is on xunit 2.9.3; adding it makes every `Fact` and `InlineData`
in the project ambiguous and produces ~850 build errors. `UiTest.Run(...)` drives the same public
session API directly and keeps xunit 2. Anything touching a control must go through it, or
Avalonia throws for want of a dispatcher.

**Know what these do and do not catch.** The XAML compiler already fails the build on a renamed
`x:Name` (CS1061) and on a broken `avares://` path, and a missing `StaticResource` key is silently
tolerated by Avalonia at runtime - all three were tested. So construction tests do not guard the
markup; what they add is that constructor *logic* cannot throw, plus a foundation for interaction
tests (`HeadlessWindowExtensions` gives `MouseDown`, `KeyPress`, `MouseWheel`). Prefer adding
interaction tests that assert observable state over more construction tests.

### Deliberately not covered

- **Rendering, layout and pointer interaction in `Tabs/` and `Main/`.** The tabs are now built
  headlessly (below), which proves they construct; whether the result *looks* right is still
  verified by running the app. `Main` itself is not constructed by any test.
- **I/O boundary classes**: `OnlineServices`' network half (`FetchManifestAsync`, `SyncFilesAsync`,
  `DownloadFileAsync`), `UpdateService` (real HTTP), `ScopeScpiClient` (real TCP),
  `MiniproProcessRunner` (spawns a process), and `DataManager`'s sync/seed/orphan-cleanup half.
  The abstraction below each of these (`IMiniproRunner`) is the thing to test, not the boundary.
  **`OnlineServices` is only half excluded.** The four predicates that validate a manifest entry
  before anything is written — `TryValidateManifestEntry`, `TryResolveValidatedLocalPath`,
  `TryNormalizeManifestChecksum`, `TryCreateTrustedDownloadUri` — are pure string/`Uri`/`Path` logic
  and *are* covered, by `OnlineServicesTests` via reflection (the same approach and the same
  reasoning as `ExternalTargetLauncherTests`). They decide where a downloaded file lands and which
  server it may come from, on input that arrives over the network, so they are a trust boundary
  rather than an I/O one.
- **`DataValidator`'s findings.** `ValidateAllDataAsync` returns a bare `Task` and reports everything
  through `Logger`, so the tests only prove it walks real data without throwing. Testing what it
  actually detects means changing it to return its findings — a public API change, and a decision for
  the maintainer rather than something to do in passing.

That extraction has been done — see `Handlers/Geometry/` above. What is left inside `TabSchematics`
is genuinely UI: event handlers, control updates and rendering. Two methods that look pure are not
(`GetOrCreateKiCadSchematicHoverHitTestCache` and `HitTestKiCadSchematicOverlayForHover` both read
instance caches and the view matrix), so they stayed put.

The label editor's snapping maths has since been extracted the same way: `LabelEditorSnapGeometry`
in `Handlers/Geometry/` now owns all ~950 lines of it, and `TabSchematics.LabelEditor.Snap.cs` is
the thin rim that reads the tab's controls and builds a `LabelEditorSnapContext` (working
highlights, drag mode, schematic name, the visible-pixel rect, and a selection predicate). That
context struct is the pattern to copy for anything similar: resolve the UI reads to plain values at
the call site and hand them over, rather than passing controls in. `EditableComponentHighlight` moved
to `Handlers/Geometry/` with it — it was a private nested type in `TabSchematics.Types.cs`, which a
`Handlers/` class cannot see. Extracting it also removed a wart: `ApplyNewLabelEditorRectangleSnap`
used to set and restore `thisLabelEditorDragMode` around four calls, and now drives each edge by
copying the context with `LabelEditorSnapContext.WithDragMode`. `ApplyResizeSnap` deliberately takes
the mode from that context and nowhere else — it used to accept a `dragModeOverride` argument
*alongside* the context, which meant one value with two sources that a caller could set to disagree.

**The KiCad trace calibration maths has since been extracted the same way**, into
`Handlers/Geometry/KiCadCalibrationGeometry.cs`, leaving `TabSchematics.KiCad.Calibration.cs` as the
rim that reads the tab's four edge fields, guards on mode, and refreshes the overlay. The value it
works on is `KiCadCalibrationBox`, and the one thing to understand before touching it is that
**mirroring is not a flag stored beside the edges — it IS the edge ordering**: a horizontally
flipped board is held as `Left > Right`. That is why the box is four doubles rather than a `Rect`
(a `Rect` normalises its edges and would silently discard the flip), why `IsMirroredX`/`IsMirroredY`
are derived rather than stored, and why arithmetic that needs ascending edges must come back out
through `WithNormalisedEdges` so the inversion is restored. It is also why `ApplyDrag` deliberately
does **not** clamp an edge from crossing its opposite: that crossing is how a board gets flipped by
dragging in the first place. The prize in this extraction was `RemapDragModeForFlip` — an
eight-handle by four-flip-state table that decides which stored edge a visually grabbed handle
actually controls. It was at 0% coverage, and a wrong arm in it is invisible: nothing throws,
dragging a corner of a mirrored board just resizes the wrong edge. `KiCadCalibrationGeometryTests`
asserts that table entry by entry and adds the involution property (remapping twice returns the
original handle) that a single-arm typo cannot survive.

The same sweep has now been done across the other tabs. What remains in `Tabs/` and `Main/` was
checked and is genuinely UI-bound, so **do not go looking for more to extract there** — the
candidates that look pure are not:

- **`TabSchematics.KiCad.Geometry.cs`** — the world↔local mapping reads `currentFullResBitmap` for
  the calibration offset scale.
- **Thumbnail bitmap builders** (`CreateScaledThumbnail`, `CreateHighlightedThumbnail`) — these need
  `RenderTargetBitmap` and therefore a display.
- **`TabConfiguration`'s launchers and `ComponentContribution.BuildPayload`** — `Process.Start` and
  zip/file I/O respectively, which rule 6 puts out of scope.

The test project lives inside the app's project folder, so [Classic-Repair-Toolbox.csproj](../Classic-Repair-Toolbox.csproj)
excludes `Tests/**` from its compile glob. Leave that exclusion in place.

## Code layout conventions

- **No MVVM.** UI logic lives directly in `.axaml.cs` code-behind files throughout the project, not in
  separate ViewModels.
- **Large controls are split into partial-class files** named `<Class>.<Area>.cs`, sitting beside the
  `.axaml.cs` (see the `TabSchematics.*` set). Aim to keep any one file under ~1,500 lines; when a file
  outgrows that, add another partial rather than letting it grow.
- **Every partial file opens with a header comment** stating what that part owns, and the `.axaml.cs`
  part carries the file map for the whole class. When you add, rename or move a part, update the map in
  the `.axaml.cs` header too.
- **Fields are declared in the part that owns them**, not collected at the top of the class. The parts
  are one class, so state is still shared across them.
- **Pure logic does not belong in a tab.** If a method touches no Avalonia control, put it in
  `Handlers/` (`Handlers/Geometry/` for maths and geometry) as a plain static class, not as a private
  member of a `UserControl` — that is the difference between logic that can be tested and logic that
  cannot. It then gets tests, per the [Tests](#tests) rules.
- Biggest files right now, for context budgeting: [Tabs/Oscilloscope/TabOscilloscope.axaml.cs](../Tabs/Oscilloscope/TabOscilloscope.axaml.cs)
  (~3,400 lines), [Main/Main.axaml.cs](../Main/Main.axaml.cs) (~2,700),
  [Tabs/Schematics/ComponentInfoWindow.axaml.cs](../Tabs/Schematics/ComponentInfoWindow.axaml.cs) (~1,800),
  [Handlers/Data/DataManager.cs](../Handlers/Data/DataManager.cs) (~1,600). Read the part you need rather
  than the whole file.

## Architecture

### Bootstrap (`Main/`)

- [Main/Program.cs](../Main/Program.cs) — entry point; initializes Velopack then starts Avalonia.
- [Main/App.axaml.cs](../Main/App.axaml.cs) — contains `AppConfig`, a static class holding nearly every
  tunable value in the app (file/folder names, sync URLs, timeouts, zoom limits, debug flags, version
  helpers). **Check here first before hardcoding a new constant elsewhere.** `App` itself wires up theme
  application (including JSON-defined user-preference theme colors), global exception logging, and the
  startup sequence: show `Splash` → `DataManager.InitializeAsync` (loads/syncs hardware data) → open
  `Main` window → fire-and-forget version check-in.
- [Main/Main.ModeHint.cs](../Main/Main.ModeHint.cs) — the khaki "what to do next" label in the tab-header
  row, shown while a mode (e.g. worklog area-marking) is waiting for the user to act. `ShowModeHint`/
  `HideModeHint`; it clears itself on the first pointer press.
- [Main/Main.axaml.cs](../Main/Main.axaml.cs) — the main window's code-behind (~2,700 lines). It acts as
  the central controller coordinating board selection, schematics zoom/pan/thumbnails, and cross-tab
  state. It reaches directly into `TabSchematicsControl` members (`currentThumbnails`,
  `highlightIndexBySchematic`), so changes to those ripple here.

### Tabs (`Tabs/`)

One folder per UI tab — `About`, `Configuration`, `Contribute`, `Feedback`, `Oscilloscope`, `Overview`,
`Resources`, `Schematics`, `Workbooks` — each an Avalonia `UserControl` (`.axaml`) with its logic in
the paired `.axaml.cs`. `Worklog/` is the odd one out: it holds the worklog's dialog windows, not a
tab.

Two tabs are conditional, both hidden by a Configuration checkbox and both shown from `Main.axaml.cs`:
`Oscilloscope` (`ApplyOscilloscopeTabVisibility`) and `Workbooks` (`ApplyWorklogBarVisibility`, which
drives the tab and the worklog bar together since they are one feature). Each moves selection to the
first still-visible tab when it is hidden while selected.

**`Tabs/Workbooks/` is partway from mockup to functional** — concept "C; Worklog tab" from
[Assets/UI mockups/worklog-mockup.html](../Assets/UI%20mockups/worklog-mockup.html), built as markup
so the layout can be tweaked in the running app, with real data and behaviour landing incrementally
on top of it. Split across `TabWorkbooks.axaml(.cs)`, `TabWorkbooks.BoardPreviews.cs` (the board
pane specifically), `TabWorkbooks.Summary.cs` (the collapsible workbook-summary strip) and
`TabWorkbooks.Export.cs` (PDF/ZIP export) — each has its own header explaining what it owns.

- **Real:** the left-hand workbook list (`RefreshWorkbooks`, one card per
  `WorklogManager.GetWorkbooksForBoard(boardKey)` result); clicking a card **activates** that
  workbook app-wide (`SelectWorkbook`, see "Activation" below) WITHOUT leaving this tab - the
  top-line and the board pane update in place for the newly active workbook; the board pane
  (`RefreshBoardPreviews`, in `TabWorkbooks.BoardPreviews.cs`) — every schematic image with one or
  more entries in the ACTIVE workbook, drawn with a 1px black outline so its boundary reads clearly
  against the pane. Each entry is drawn one of two ways, mirroring `TabSchematics.Worklog.cs`'s own
  `ShowMarkedArea` branch exactly: ticked gets a dashed bounds rectangle on the schematic tab's own
  `WorklogEntriesOverlay` plus its "#N" badge anchored to that area; unticked gets NO rectangle, and
  its badge is parked in the image's own top-right corner instead (`ParkedBadgeGeometry.ArrangeInTopRightBlock`,
  the same geometry the real Schematics tab's parked pills use), stacking with any other parked
  badges rather than overlapping them. Getting this backwards - anchoring every badge to its marker
  regardless of `ShowMarkedArea` - was a reported bug; `WorkbooksBoardPreviewTests` pins both
  branches down by actual on-screen position, not just presence, specifically to catch a
  regression like it. **It was then reported a second time against the SCHEMATICS TAB THUMBNAILS**,
  which had the same defect and are now fixed the same way (`ThumbnailWorklogPillsOverlay`, pinned
  by `ThumbnailWorklogPillsTests`) - so all THREE surfaces that draw these pills now share
  `ParkedBadgeGeometry.ArrangeInTopRightBlock`. Any fourth must too. **Every pill is clickable** (`OnPreviewBadgePointerPressed`), opening the
  EXACT SAME `WorklogEntryEditorWindow` the Schematics tab's own "Show worklogs" badges open,
  including the "Mark components in scope"/"Mark components completed" checklist
  (both tabs now call the one shared `WorklogEntryScope.BuildComponentsInScope` in `Handlers/Data/`,
  rather than the near-identical copy each used to carry) - an earlier version omitted that section because it needs a highlight-rect cache
  only `TabSchematics` built; reported as the two modals not actually being identical, and fixed by
  reading that SAME cache off `MainWindow.TabSchematicsControl.highlightRectsBySchematicAndLabel`
  (via the `HighlightRectsBySchematicAndLabelForPreviews`/`...OverrideForTests` seam, the same
  override-then-real-`MainWindow` pattern `CurrentBoardDataForPreviews` already used) rather than
  building a second copy - `Main.ApplyRegionFilterAsync` already keeps that cache current on every
  board load and region switch regardless of which tab is selected, so nothing new has to be kept
  in sync. **Clicking anywhere else on a preview** (not a pill - `OnPreviewBadgePointerPressed` sets
  `e.Handled` so the two clicks cannot both fire) **selects that schematic** (`SelectSchematic`): a
  highlighted border (the same IndianRed `Main_TabUnderline_Selected` accent the selected workbook
  card uses) and the entry list on the right (marker 4) switches to that schematic's entries, each
  rendered by `BuildEntryDetailCard` as ONE 1px-bordered card - not one border per field, a reported
  regression from an earlier three-separately-bordered-panel layout - holding four stacked rows:
  `"#{N} {Title}"` (the "#N" a small filled badge in the category colour, exactly like
  `WorklogEntryEditorWindow`'s `EditorIdBadge` - filled is still right here, since it names which
  workbook entry this is, not a selection state), the description, a category chip and status pill,
  then a stats row. The category chip and status pill (`BuildFilledCategoryChip`/`BuildFilledStatePill`,
  renamed from `BuildFilled*`, which described the opposite of what they build) render in the
  UNSELECTED/outlined visual `WorklogEntryEditorWindow` itself uses
  for a NOT-currently-chosen category chip / state pill (`Form_Bg` background, a 1px `Form_Border`
  outline) rather than that window's filled "selected" look - reported explicitly: this list has no
  selection concept the way a click in the full editor does, so a filled pill here would falsely
  read as "this is the chosen one." The stats row (`BuildEntryStatsRow`) sums total hours and cost
  across the entry's `WorkDoneItems` (`"{hours} h"` / the bare cost number, matching
  `WorklogEntryEditorWindow`'s own `SummaryText` formatting exactly - no currency symbol) plus how
  many comments, links, photos and files the entry carries - one number each, added because a
  workbook's worth of pills gave no sense of how much was behind each one without opening it.
  Deliberately not the smaller dot-plus-label the board pane's own pills use for category, and not
  the anchor-tag/timestamp/photo layout the mockup drew there, both gone. Defaults to the
  alphabetically-first schematic with entries when nothing is yet selected, and stays on the current
  selection across a rebuild if it is still shown (a save via a pill's editor re-enters
  `RefreshBoardPreviews` and must not reset it) - same "keep if valid, else fall back" rule
  `RefreshWorkbooks` applies to the selected workbook. All real pieces are refreshed from
  `Main.RefreshWorklogBar`, the one funnel every worklog change already passes through, so none of
  them can go stale in a case the bar handles. **The Legend panel the mockup drew under the board
  pane is gone** - removed as unneeded once every pill already named its own category and state.
- **The top-line (the highlighted bar above the board pane) is now two lines plus right-aligned
  actions.** Line 1 is the existing "#{N} · {Title}" and status pill; line 2 is the selected
  workbook's **Note** (`WorkbookHeaderNoteText`, `WorkbookRecord.Note` - the create dialog's
  optional free-text field, distinct from `Title`, which the dialog itself labels "Description"),
  its own row collapsed entirely when blank (most workbooks have no note) rather than showing an
  empty line. Kept on its own line - not folded into line 1's `WrapPanel` - because a long title
  plus a long note in one wrapping run read as one run-on line with no clear boundary between
  them. Right-aligned against both lines: **"Edit workbook"** and **"Delete workbook"** buttons
  for the selected workbook, both plain text buttons (not icons) grouped in
  `WorkbookHeaderActionsPanel`, hidden together when no workbook is selected. **Edit**
  (`OnEditWorkbookClick`) reopens `CreateWorkbookWindow`, the SAME modal "Create new workbook"
  uses (whose own submit button now reads "Create workbook", not "Create"), switched into edit
  mode via `InitializeForEdit` (pre-fills the fields, shows the real id instead of the next-id
  preview, relabels the submit button to "Update workbook") - the same add/edit-share-one-dialog
  pattern `WorklogAddLinkWindow.InitializeForEdit` already uses for a link row, so title/note
  editing has exactly one implementation to keep in sync rather than two. It calls
  `WorklogManager.UpdateWorkbook` (title/note only; id, board key, status, start date and entry
  count are untouched) and then a bare `Main.RefreshWorklogBar` - not `ActivateWorkbook` - since
  Edit only ever acts on the workbook already active. **Delete** (`OnDeleteWorkbookClick`) confirms
  first via `DeleteWorkbookWindow` (naming the workbook in its message, since several cards can be
  on screen at once; `CanMinimize="False"` alongside its existing `CanResize="False"` so the title
  bar carries only a close button, and Enter is wired to Cancel rather than the default "submit" -
  the one modal in the app where Enter must NOT confirm, since confirming here is a permanent
  delete), then calls `WorklogManager.DeleteWorkbook`, which removes the workbook's entire folder -
  entries, photos and files included, per the class's own "one folder is the whole workbook" model
  - and refreshes. Deliberately no separate "which workbook to select next" step: the deleted
  workbook's id, if it was the board's saved `ActiveWorkbookIdByBoard` entry, now names nothing on
  disk, so `WorklogManager.ResolveActiveWorkbook`'s existing stale-id fallback lands the refresh on
  the board's newest remaining workbook automatically - the same fallback that already covers a
  workbook deleted by hand outside the app.
- **A collapsible SUMMARY STRIP sits under the top-line's Note** (`TabWorkbooks.Summary.cs`).
  One always-visible headline — `7 worklogs · 12.5 h · 430 · 4 open` — with a chevron that expands a
  breakdown by category, by state, by attachment counts (comments/links/photos/files/work-done) and
  by component scope. **It says "worklogs", not "entries"**: the app calls these worklogs everywhere
  the user can see one, and "entry" is internal vocabulary that leaked out through this line once.
  **The category and state counts are drawn as the same non-selectable PILLS the rest of the app
  uses**, each carrying its count (`[ 3 Note ]`), rather than as plain text — every value
  including the zeroes, so the row does not change width as a workbook is worked on. **Every number
  in the strip is bold and the words are not**, which is why `WorkbookSummary` hands back `Stat`
  parts (prefix/number/suffix) rather than finished strings: a `TextBlock` cannot mix weights within
  one `Text`, and re-finding the digits in a formatted string would have to guess about `0.5 h`.
  Those blocks therefore carry `Inlines` with `Text == null` — a test reading only `Text` sees them
  as blank. **A COUNTED pill carries NO icon**, unlike every other informational pill: a padlock or
  category glyph between a number and its label reads as a third piece of information rather than as
  decoration. An UNCOUNTED pill keeps its glyph, because on an entry card that glyph is the only
  thing separating Open from Closed at a glance — `WorklogInfoPillBuilder` branches on `count`, and
  both halves are pinned. The numbers all come from
  [Handlers/Data/WorkbookSummary.cs](../Handlers/Data/WorkbookSummary.cs) (pure, unit tested), which
  the PDF export prints as its own opening section too — so an exported document cannot report
  different totals from the screen it was produced from. Collapsed by default and persisted in
  `UserSettings.WorkbooksSummaryExpanded` (per user, not per board), re-applied on every refresh
  because this header is rebuilt on every board change and entry save. The components line is
  hidden outright when the workbook scopes none, rather than showing a permanent zero. It is a
  `Button` styled flat (`Button.WorkbooksSummaryToggle`) rather than an `Expander`, whose border,
  background and padding would all have to be undone for a line inside an existing header row.

- **Each entry's detail card in the right-hand list is CLICKABLE**, opening the same full editor
  its pill on the board pane opens — asked for explicitly, since the card is simply the same entry
  rendered larger. Both go through the one `OpenEntryEditor`; `OnPreviewBadgePointerPressed` is now
  a wrapper around it, so the two cannot open different modals (the earlier version of exactly that
  complaint is what the shared `WorklogEntryScope.BuildComponentsInScope` already exists for). The
  card's schematic bitmap is resolved from the ENTRY's own schematic name rather than the selected
  preview, so a future list showing more than one schematic's entries cannot hand the editor the
  wrong board image. **Hovering a card outlines it in the same IndianRed accent a SELECTED schematic
  preview uses** (`ApplyEntryCardHoverBorder`) — one colour language across the tab for "the thing
  you are about to act on". Hover rather than selection, because this list has no selection: a card
  is a button. It is 2px at rest as well as hovered, for the same reason the previews are — growing
  1px to 2px would reflow the card as the pointer crossed it.

- **Each entry card carries a "Delete worklog" button in its TOP-RIGHT corner**, the per-worklog
  twin of the header's "Delete workbook" and confirmed by the same shape of modal
  ([DeleteWorklogWindow](../Tabs/Worklog/DeleteWorklogWindow.axaml.cs) — a near-copy of
  `DeleteWorkbookWindow`, kept as its own window rather than merged into a shared "confirm a
  delete" dialog: the two say different things about different objects, and the point of the copy
  is that changing one cannot silently change what the other promises about a permanent delete).
  **Enter and Escape both CANCEL there, on the Tunnel route** — same reasoning and the same fix as
  the workbook dialog, since a focused Button consumes Enter on the bubbling route and would delete
  the worklog on a reflexive keypress. It is a real `Button`, not a click region, so the press never
  reaches the card's own `PointerPressed` underneath it — a single click must not both open the
  editor and raise a delete confirmation over it. It carries the destructive `Button_Cancel_*`
  brushes (the same three keys the header's Delete names) and the default ARROW cursor rather than
  the card's Hand, which would say it does the same benign thing the rest of the card does.
  `WorklogManager.DeleteEntry` removes the entry's row from `entries.json` AND its whole
  `worklog_{id}` attachment folder — the JSON write goes FIRST, so a failed write leaves the entry
  whole rather than stranding a surviving row pointing at photos that are gone. **The remaining
  entries deliberately KEEP their ids**: those ids are what the board pills, the cards and the
  exported PDF all show, and they name the attachment folders, so a gap in the numbering is the
  correct record of a deletion rather than something to close up. The refresh goes through
  `Main.RefreshWorklogBar`, the one funnel every worklog change passes through, so the Schematics
  tab's overlay rectangles and thumbnail pills lose the deleted entry too.

- **A deleted id is never handed out again — neither a workbook's nor a worklog's.** Both used to be
  allocated as "highest currently on disk, plus one", so deleting the top one let the next record
  take its number back. That is not cosmetic: a workbook id names a PDF that has already been
  exported and very likely emailed (`Workbook_2_Commodore_C64_20260904`), and both ids name real
  folders (`2/`, `worklog_2/`) whose contents a recreated record would inherit — so a reused number
  silently makes an old document describe a different repair. Two persisted counters record the
  highest id ever HANDED OUT: `counters.json` at the workbook root for workbook ids, and each
  workbook's own `index.json` (`WorkbookRecord.LastEntryId`) for the entries inside it. Delete #2 of
  two workbooks and the next is **#3**; the gap is the correct record that #2 existed. Entry
  numbering is **per workbook** (the counter travels in the folder it numbers and is deleted with
  it), which is right because entry ids only have to be unique within their workbook and the
  workbook id itself is never reused.

  **There is deliberately NO migration** — nothing seeds a counter from existing data and nothing
  rewrites a workbook written before the counters existed. What the allocators do instead is refuse
  an id that is already taken on disk (`SkipIdsAlreadyOnDisk`, and the walk in `PeekNextEntryIdIn`):
  a counter starting from zero would otherwise hand out #1 while a `1/` folder still sits there, and
  since `CreateWorkbook`'s `Directory.CreateDirectory` succeeds silently on an existing folder, the
  new workbook's `index.json` would **overwrite the old one in place** and inherit its entries and
  attachments. That skip is a floor, not a migration — the counter is still the only thing that can
  stop a DELETED id coming back, which checking disk alone never can. The entry side exempts the
  draft's own `reservedId` from its folder check: that folder belongs to the entry being saved right
  now, and skipping past it would misnumber the entry and strand the photos just attached to it.
  One consequence worth knowing: `AddEntryRecord` can no longer be allocated an id whose attachment
  folder exists, so `MoveEntryAttachmentsFolder`'s merge is now unreachable through that path and is
  tested directly (by reflection) rather than deleted — it remains the safety net for a destination
  that appears between the allocation and the move.

- **Every workbook can be EXPORTED** (`TabWorkbooks.Export.cs`), from **two buttons** — "Export to
  PDF" and "Export to ZIP" — on their own SECOND ROW inside `WorkbookHeaderActionsPanel`, under
  Edit/Delete: four buttons across one line crowded the workbook title beside them and pushed the
  header wider than a narrow window could hold. Both rows are right-aligned so they share a right
  edge despite differing widths. **PDF** is the customer-facing document; **ZIP** is that same PDF plus
  the workbook's original photos and attached files under one folder per entry (a PDF shows a photo
  at page resolution and cannot carry an attached datasheet at all). Both go through one
  `ExportWorkbook(bool asZip)`, so the picker, the guard, the off-thread write and the error
  handling cannot drift apart. **The ZIP was originally offered only as a second file type inside
  the save dialog and was reported as missing entirely** — a format reachable only by opening a
  dropdown in a dialog the user opened for another reason is not discoverable, so the format now
  comes from the button and the extension is enforced rather than read back off the returned name.

  **Both exports are named `Workbook_{id}_{Hardware}_{Board}_{YYYYMMDD}`** — the BoardKey is
  `Hardware|Board`, so its halves become their own underscore-separated segments. The workbook
  TITLE is deliberately absent: it is a sentence, often carrying a customer's own details, on a file
  about to be emailed. Inside the ZIP, each entry's attachments sit under **`worklog_{id}`** — the
  same folder name `WorklogManager.BuildEntryAttachmentsFolderName` gives them in the local Workbook
  folder, so what a recipient unpacks matches what the repairer sees on their own disk. That helper
  is the single definition of the name; it was written out in four places before.
  **The extension comes from the BUTTON, and `WorkbookExportModel.EnsureFileExtension` REPLACES the
  other format's rather than appending to it** — typing `repair.pdf` into the ZIP dialog produced
  `repair.pdf.zip`, and the picker's overwrite prompt had been shown for a different name, so an
  existing file of that name was overwritten without asking. Only `.pdf`/`.zip` are replaced; an
  unrelated suffix (`board rev 2.5`) is kept and the real extension appended.
  **Neither button carries an icon**: the fa-regular file-pdf glyph rendered as a blank box in the
  shipped font subset.

  **`ZipArchive` in `Create` mode is write-forward only** — `GetEntry` and `Entries` both throw
  `NotSupportedException("Cannot access entries in Create mode")`. `WriteZip` therefore tracks the
  names it has written in a `HashSet` and never asks the archive. A collision check that called
  `GetEntry` shipped once and **crashed the whole application** on the first export of any workbook
  holding an attachment — a workbook with none never reached the call, which is exactly why
  `WorkbookZipExportTests` is built entirely around workbooks that HAVE attachments. That file is
  the one deliberate exception to "the PDF writer is not tested": the archive's contents and naming
  are this app's decisions, not QuestPDF's. It also pins `EnsureIconFontLoaded` being safe to call
  with no Avalonia available — the other half of the icon-font contract above.
  [WorkbookExportModel](../Handlers/Data/WorkbookExportModel.cs) decides WHAT goes in and in what
  order (grouped per schematic, entries by id, missing attachment files dropped, an entry with no
  schematic filed under `(no schematic)` rather than silently lost) and is unit tested;
  [WorkbookPdfExporter](../Handlers/Data/WorkbookPdfExporter.cs) only paints it, and its LAYOUT is
  deliberately not tested — asserting on PDF bytes tests QuestPDF rather than this app.

  **The PDF mirrors the app's own visuals**, asked for directly: outlined status pills and category
  chips with real Font Awesome icons, the filled category-coloured "#N" badges, and each schematic
  drawn with its worklog areas washed and outlined in the category colour. Every schematic starts on
  a **new page**, its image spans the **full page width** and carries the same **1px outline** the
  Workbooks board pane draws, so a marked area is large enough to locate on the board. An entry with
  `ShowMarkedArea` off gets no rectangle and its pill parks top-right, mirroring what all three
  on-screen surfaces do.

  **The pill shapes are the app's own, and QuestPDF can express all of them** — `CornerRadius` is
  available and was simply not used at first, which is why the exported pills shipped as square
  boxes and were reported. A status pill is fully rounded, a category chip only softened and a "#N"
  badge carries a real **white disc** with the state padlock inside it, matching
  `WorklogInfoPillBuilder`'s 10px/3px split and `WorklogBadgeBuilder`'s disc exactly.

  **Each photo sits in its own bordered panel** holding the picture, its FILE NAME and its comment.
  The border is what makes the grouping structural — with two photos side by side, the gap between
  a picture and its own caption is the same as the gap to the next one's, so a reader has to infer
  which belongs to which. The file name is printed even when there is no comment, so a recipient can
  find that exact photo in the ZIP's `worklog_{id}` folder.

  **Web links are real, VISIBLE hyperlinks** — blue and underlined, via QuestPDF's
  `TextDescriptor.Hyperlink`. They shipped as plain black text at first, which was reported: a PDF
  viewer gives no hover cue of its own, so an unstyled hyperlink is indistinguishable from prose
  until someone happens to click it, and on a printed page the styling is the only cue that survives
  at all. Which runs of free text count as links is decided by `TextLinkFinder` — the SAME pure
  helper the on-screen renderer uses, so the document cannot linkify things the app does not.

  **A worklog's LINK ROWS are the exception, and get `BuildLinkTarget` instead.** `TextLinkFinder`
  deliberately rejects a bare `example.com` (correct when scanning repair notes full of part numbers
  and file names), but a link row is a DECLARED destination and the add-link dialog stores whatever
  the user typed without normalising a scheme onto it. So those rows are linked whole, with `https`
  filled in for the target when the stored text lacks a scheme — **a PDF hyperlink with no scheme
  is silently ignored by every reader**, which would produce a link that looks right and does
  nothing. Pinned by `WorkbookZipExportTests` via reflection.

  **Nothing in this document may be a zero-sized container holding text** — QuestPDF answers that
  by failing the whole render. Three guards exist for it and all three are load-bearing: the
  no-icon-font fallback collapses with `Height(0)`, the badge's white disc is drawn only when
  `IconFontAvailable`, and `PillLabel` substitutes a fallback for a blank State or Category (both
  are plain strings in `entries.json`, so a hand-edited or older-build record can carry an empty
  one). Parked pills wrap into a grid via `ParkedBadgeGeometry.GetGridShape` rather than stacking in
  one unbounded column, which would grow past the bottom of the image.

  **`CategoryHexColor` is case-INSENSITIVE**, like every other category comparison in the app. It
  was a plain `switch` while the `CategoryGlyphs` dictionary beside it was `OrdinalIgnoreCase`, so
  an entry stored as `"note"` drew the right icon in the unrecognised-category grey.

  **A "#N" badge carries TWO colours, and they are different channels**: the fill is the CATEGORY
  colour and the padlock inside the white disc is the STATE colour — so a Closed Issue is a green
  padlock on a red badge. `WorklogBadgeBuilder` takes `categoryColor` and `stateColor` as separate
  arguments for exactly this reason. Colouring the glyph to match its own badge was reported: it
  made the badge report the category twice and the state not at all.

  **Every number in the summary is bold and the words are not**, as on screen. That is why the PDF
  walks `WorkbookSummary`'s `Stat` parts rather than its `Format*` helpers — the identical reason
  `TabWorkbooks.Summary.cs` does. The **state pills sit on their own line** under the categories:
  two different kinds of pill running together read as one undifferentiated list of five.

  **How the overlay is positioned — and the trap to avoid.** An entry's area is stored in the
  schematic's own PIXEL coordinates, while the page draws the image at whatever width the margins
  leave, a size QuestPDF decides during layout and never reports. So the placement is entirely
  PROPORTIONAL: [ExportOverlayGeometry](../Handlers/Geometry/ExportOverlayGeometry.cs) turns a pixel
  rect into fractions of the image, and each band is then expressed as an **aspect ratio**, which
  QuestPDF can satisfy against any width. No page dimension appears in the drawing code at all.

  **There is no percentage unit anywhere in QuestPDF** — every `Padding*`, `Width` and `Height`
  takes an absolute length (points by default). The first version computed the fractions correctly
  and then passed them to `PaddingLeft`/`PaddingTop` multiplied by 100 believing those were
  percentages, so "58% across" became "58 points across" and a marked area covering a tenth of the
  board was drawn covering most of it. Nothing threw; it was caught by holding the PDF next to the
  screen. `ExportOverlayGeometryTests` now pins the fractions against a REAL entry's stored
  coordinates, and four of them fail against that ×100 version.

  **Rows have relative sizing; Columns do not.** `Row.RelativeItem(weight)` splits the horizontal
  axis directly, but a `Column` has no equivalent, so the vertical axis is expressed as "this band
  is X wide and Y tall, i.e. this ratio" via `TryBuildBandAspectRatio`. Every empty band is OMITTED
  rather than emitted at zero — a zero-weight `RelativeItem` and a zero-ratio `AspectRatio` are both
  degenerate and QuestPDF rejects the whole layout.

  **A zero-sized container holding text fails the ENTIRE document.** QuestPDF answers it with
  `DocumentLayoutException` and abandons the render rather than clipping one element, so a single
  bad band takes down the export. Two things follow, both of which shipped broken once: a vertical
  `AlignMiddle` inside a row whose height is still being negotiated measures its child against zero
  height, and `container.Text(string.Empty)` still demands a line box — which is why the no-icon
  fallback collapses the element with `Height(0)` and the badge's white disc is only drawn when
  `IconFontAvailable`. `WorkbookZipExportTests` covers an entry WITH a shown marked area precisely
  because every other test there uses parked entries that never reach this code.

  The pixel dimensions come from `TryReadImageSize`, which parses the PNG/JPEG **header** rather
  than decoding the file — a 4220x2941 board scan would otherwise cost ~47 MB per schematic just to
  read two numbers. **The JPEG side is a marker walk, and which markers are frame headers is the
  whole subtlety**: of `0xC0`-`0xCF` only C0-C3, C5-C7 and C9-CB carry dimensions (C4/C8/CC are
  DHT/JPG/DAC and CD/CE/CF are DNL/DHP/EXP), and the standalone markers (`0x01` TEM, `0xD0`-`0xD9`)
  carry no length word, so reading two bytes after one seeks into image data. Getting either wrong
  is silent — a bogus size, no exception — and every marked area on that schematic then lands
  nowhere. Covered by `WorkbookZipExportTests`.

  **An area is CLIPPED to the image, not clamped inward.** `TryBuildAreaFractions` converts both
  edges and intersects; clamping only the origin kept the full width and slid the rectangle off the
  thing it marks (an entry at `x=-50 w=100` on a 1000px image drew 100px starting at 0).

  **The icon font is loaded on the UI thread, deliberately** (`EnsureIconFontLoaded`, called from
  the export handler before its `Task.Run`). The .otf is an `AvaloniaResource` compiled into the
  assembly — not a file on disk, and not a plain manifest resource — so only Avalonia's `AssetLoader`
  can read it, and that resolves `IAssetLoader` out of a locator **not available on a background
  thread**. Doing the load inside the writer threw `InvalidOperationException` inside a `try`, so
  every exported document silently came out with no icons at all until the PDF was inspected. Bytes
  are now read on the UI thread and cached; the QuestPDF registration, which needs no Avalonia
  service, happens lazily wherever the export runs. A failure still degrades to omitting icons
  rather than throwing — an unregistered font renders as blank boxes, which looks like a defect.
  **The export is deliberately not opened afterwards**: `ExternalTargetLauncher` admits a local path
  only inside the data root, and an export is saved wherever the user chose, so calling it would
  refuse every export and log a warning about the file it had just written.

- **The "Find a previous repair" field is now wired up** and filters the whole tab as you type.
  The query language lives in [Handlers/Data/WorklogSearchQuery.cs](../Handlers/Data/WorklogSearchQuery.cs)
  (pure, unit tested by `WorklogSearchQueryTests`): space-separated terms are ANDed, `"a phrase"`
  quotes a run containing spaces, a leading `-` excludes, matching is case-insensitive substring
  throughout (so `p c u` finds `CPU`, and `"full text"` finds `Afull textB`). An empty box is not a
  filter and matches everything. Which fields are searched is
  [WorklogSearchIndex](../Handlers/Data/WorklogSearchIndex.cs): every user-typed TEXT field on the
  workbook (title, note) and on each of its entries (title, description, category, schematic name,
  component labels, plus every link/comment/work-done/photo/file row) - **numbers
  are deliberately excluded** (ids, hours, cost, display order, dates), since a search for "2"
  would otherwise match nearly everything through fields the user never sees as text.
  **Status/State ("Open"/"Closed") are excluded too**, and for the same class of reason: every
  record carries one of two values, and "open" is a word this domain uses constantly ("open
  circuit", "opened the case"), so including them made that search match almost the whole database
  - and because terms are ANDed across a record, "open trace" then matched any Open workbook
  mentioning "trace" anywhere. Both values already have an always-visible pill, so they filter by
  eye far better than by substring. Category IS searched: its three values are descriptive and none
  turns up incidentally in repair notes.
  A workbook is shown when its own text matches OR any of its entries does; when the workbook
  itself matched but no individual entry did, ALL of its entries stay visible rather than leaving
  the result looking empty. `RefreshWorkbooks` computes the matched entry ids ONCE into
  `thisMatchedEntryIdsByWorkbookId`, and the board pane and entry list both narrow from that same
  set rather than re-running the query, so the three surfaces cannot disagree about what matched.
  Each of the three empty states says "no match" rather than "none recorded yet" when a search is
  what emptied it - the "yet" wording reads as data loss otherwise.
  **Which workbook the TAB shows follows the filtered list, while which workbook is ACTIVE does
  not.** `ResolveActiveWorkbook` still runs against the unfiltered set (typing must never redirect
  where "Add worklog" writes), but if the active workbook is filtered out, the top-line and the
  right-hand side move to a workbook that survived - otherwise the top-line named a workbook absent
  from the list, with live Edit and **Delete** buttons acting on it, above a board pane blanked
  because none of its entries matched.
  **The query is dropped on a BOARD change** (`ClearSearchForBoardChange`, called from
  `Main.OnBoardSelectionChanged` before its refreshes) - a board switch is a change of subject, and
  carrying the filter over lands the user on an empty list for a board they just chose, with the
  reason in a box they are not looking at; `OnHardwareSelectionChanged` clears
  `ComponentSearchTextBox` for the same reason. Every OTHER refresh trigger keeps it, because
  **`RefreshWorkbooks` re-reads the box itself rather than trusting the copy the `TextChanged`
  handler cached** - so an entry save or a workbook create/delete cannot silently revert to the
  unfiltered list; it is also
  what lets the headless tests drive the real path, since a tab that is never attached to a visual
  tree never raises `TextChanged`. **Typing is debounced** (`thisSearchDebounceTimer`, 200ms):
  filtering costs a `GetEntries` per workbook plus a full board-pane rebuild, synchronously on the
  UI thread, so a rebuild per keystroke made one typed word hundreds of file reads.
  `GetEntriesForThisPass` caches those reads **within** one pass (cleared at the start of every
  refresh, never across them) so the filter and the board pane do not each re-read the same file.
  Matched runs are highlighted via
  `WorklogSearchQuery.SplitIntoSegments` (the segment maths is on the `Handlers` side precisely
  because an off-by-one there drops or doubles characters on screen) rendered as `Run`s in the
  `Workbooks_SearchHit_Bg`/`_Fg` wash; with no search active the blocks carry plain `Text`, which
  is cheaper to lay out and is what `TabWorkbooks.BuildHighlightedTextBlock` falls back to. **Note
  for tests reading a card's text: a highlighted `TextBlock` has `Text == null` and its content in
  `Inlines`**, so a reader that only looks at `Text` sees a highlighted card as blank - see
  `WorkbooksSearchTests.VisibleText`.

**Board-data timing.** `Main.OnBoardSelectionChanged` used to call `RefreshWorklogBar` (which
rebuilds this whole tab) BEFORE it awaited `DataManager.LoadBoardDataAsync`, so that pass ran against
the PREVIOUS board's data, or none at all on the session's first board. Reported as "even if my
workbook has data, it does not get reflected in the tab". **That call now sits AFTER
`_currentBoardData` is assigned**, so it runs once, against the right board - the ordering fix rather
than the patch-up second call it originally got. The three early-return paths in that method
(no selection, no data file, unreadable board data) each refresh explicitly, so the tab shows the
board that IS selected rather than the previous one's contents; a SUPERSEDED load deliberately
refreshes nothing, since the newer load owns every surface.

`TabWorkbooks.RefreshBoardPreviewsForCurrentSelection` survives for a different job: it is called
from `Main.SetComponentHighlightRects` whenever the component highlight-rect cache is replaced (a
board load finishing, or a region switch), because the pane's badges are clickable before that cache
is populated and a click in that window silently dropped the editor's component checklist.

**Every worklog change redraws the Schematics tab's overlay and thumbnail pills.**
`Main.RefreshWorklogBar` calls `TabSchematics.SetShowWorklogEntriesList` when the shown WORKBOOK
changes, and `RefreshWorklogEntriesListForCurrentWorkbook` otherwise. That second branch was missing:
editing an entry inside the workbook already on screen changes neither of `SetShowWorklogEntriesList`'s
two arguments, so the overlay and the thumbnail pills kept drawing the PRE-edit record until
something else happened to rebuild them - reported as a marker not updating after adding one or
ticking "Show marked area". Editing from the **Workbooks tab** was the worst case, since its save
path funnels through `RefreshWorklogBar` alone and so nothing redrew the schematic overlay at all.
Pinned by `WorklogOverlayRefreshTests`, four of which fail against the missing-branch version.

**Activation.** Before this, every worklog-facing control (the bar, "Show worklogs", "Add worklog")
always acted on `WorklogManager.GetLatestWorkbookForBoard` - there was no way to point any of them at
an older or closed workbook. Clicking a card in the Workbooks tab now overrides that: `Main`'s
`ActivateWorkbook(boardKey, workbookId)` saves the choice to `UserSettings.ActiveWorkbookIdByBoard`
(per board, persisted) and calls `RefreshWorklogBar`. Creating a workbook activates it too
(`OnWorklogCreateWorkbookClick`); a bare refresh would leave the bar on a previously-activated older
one and write the next drawn entry into it. `ActivateWorkbook` also cancels any in-progress
entry-drawing mode, which captured the OLD workbook's id when it started.

**`WorklogManager.ResolveActiveWorkbook(workbooks, savedActiveId)` is the ONE place "which workbook"
is decided** - saved id if it still names a workbook on this board, else the newest. Pure, unit
tested, and called from both `Main.ResolveActiveWorkbookForBoard` and `TabWorkbooks.RefreshWorkbooks`,
which used to implement the same rule separately in two different shapes.

`TabWorkbooks.SelectWorkbook` has NO "already selected, nothing to do" guard, deliberately:
`RefreshWorkbooks` highlights a default card without saving anything, so the card on screen can look
active while `ActiveWorkbookIdByBoard` is empty, and an id-equality early return made clicking that
exact card a no-op. It DELIBERATELY does not switch tabs - activating a workbook is meant to be seen
on this tab. `Main.SwitchToSchematicsTab` still exists and is still used, just not from here:
`OnWorklogAddEntryClick` calls it, since drawing a new entry needs the real schematic view.

Activation goes through a settable `Action<string, int>` (`thisActivateWorkbook`, set by `Initialize`
to `Main.ActivateWorkbook`) rather than an `if (MainWindow != null)` branch. That matters for tests:
with the branch, EVERY headless test ran the no-`MainWindow` side and the shipped path - persist,
then re-derive the selection from the saved id - was pinned by nothing at all. Tests now inject
`ActivateWorkbookOverrideForTests` and drive the real path.

**Splitters.** Both of this tab's `GridSplitter`s persist their width, the same pattern
`UserSettings.LeftPanelWidth` (`Main`'s own left sidebar) and `TabSchematics.ApplySchematicsSplitterRatio`
follow: `UserSettings.WorkbooksLeftPanelWidth` for the outer one (left workbook list vs. everything
else) and `UserSettings.WorkbooksEntryListWidth` for the inner one (board pane vs. entry list).
Unlike the Schematics tab's splitter these are plain app-wide pixel widths, not a per-board ratio -
this tab's layout does not depend on which board is selected. `Initialize` applies both (via the
private `ApplySplitterWidths`, exposed to tests as `ApplySplitterWidthsForTests`) right after
`InitializeComponent`, clamped to a usable range so a width saved on a large monitor cannot restore
off-screen on a small one. Each splitter's `PointerReleased` handler
(`OnOuterSplitterPointerReleased`/`OnBoardEntrySplitterPointerReleased`) saves the new width back,
deferred via `Dispatcher.UIThread.Post` so it reads the column's `Bounds` after the drag has actually
been applied - the same reason `Main`'s own splitter handler defers.

**Both are wired with `AddHandler(..., handledEventsToo: true)`, not a `PointerReleased="..."` markup
attribute** - `GridSplitter` marks the event handled as it finishes its drag, so a plain subscription
never runs and neither width was ever saved. Same as `Main.OnMainSplitterPointerReleased` and
`TabSchematics`, which both carry the same comment.

Cards, the top-line pill and preview badges are all built in code rather than by a `DataTemplate`,
because their brushes need the two-step `Application.Current` + `ThemeVariant` lookup a template
binding cannot express — the same reason `Main` builds the worklog bar's own pill in code.

**Status pills and category chips have exactly TWO looks, and mixing them up has been reported
twice.** A pill is either SELECTABLE — only inside `WorklogEntryEditorWindow`, where clicking one
chooses it, and the chosen one is FILLED with its colour — or INFORMATIONAL, which is everywhere
else. Every informational one now comes from
[Handlers/Theme/WorklogInfoPillBuilder.cs](../Handlers/Theme/WorklogInfoPillBuilder.cs): a `Form_Bg`
fill, a **1px** border in the thing's OWN colour (the state colour for a status pill, the category
colour for a category chip), glyph and label in that same colour. Five sites used to draw these by
hand — the worklog bar and the workbook card and the top-line at 2px in the status colour, the entry
detail card's two at 1px in GREY — each under a comment asserting they all matched, which was true
of no two of them. **Do not draw one of these by hand**; the builder is the only place that visual
is decided, and `WorkbooksSummaryAndPillsTests` pins the border width and colour precisely because
those are the two axes that drifted.

**Schematic bitmaps are shared, not decoded per rebuild.** `thisSchematicBitmapsByPath` holds one
decoded `Bitmap` per image path for the life of one ATTACHMENT, disposed in
`OnDetachedFromVisualTree`. A fresh `new Bitmap(path)` per preview per pass stranded a
full-resolution decode every time (a 4220x2941 schematic is ~47 MB of BGRA) on every board change,
entry save and workbook create/close. Disposing on clear instead would be WORSE: `ShowDialog` does
not block the dispatcher, so `RefreshWorklogBar` can re-enter while a badge's editor is up, and that
editor documents that its schematic bitmap belongs to the caller - disposing under it is an
`ObjectDisposedException` on the render thread, fatal in Avalonia.

**`OnDetachedFromVisualTree` is NOT "the tab is going away" - a `TabControl` detaches the previous
tab's content on every tab SWITCH.** The preview `Image` controls keep their `Source` across a
detach, so disposing the cache without tearing the pane down first left every one of them holding a
freed Skia surface, and the next render pass over them threw `ObjectDisposedException` on the RENDER
thread - fatal, and reported as the app crashing on switching away from Workbooks. So detach now
calls `ClearBoardPreviewsBeforeDisposingBitmaps` FIRST and disposes second (nothing then references
a disposed bitmap), and `OnAttachedToVisualTree` calls `RefreshWorkbooks` to rebuild when the tab
comes back. Both halves are pinned by `WorkbooksBoardPreviewTests` - the detach test asserts no
`Image` is left holding a `Source`, and fails against the dispose-without-clearing version. **Keep
that ordering**; anything added later that renders one of these bitmaps must either live inside
`BoardPreviewPanel` or be cleared alongside it.

Its `Workbooks_*` theme keys in `App.axaml` are pinned by `WorkbooksPaletteTests` - which now covers
only the six keys the tab actually paints; sixteen further mockup-era keys were referenced by nothing
but that test and have been deleted along with it. The list, selection and activation chain are
pinned by `WorkbooksListTests`; the board pane by `WorkbooksBoardPreviewTests` (uses
`TabWorkbooks.CurrentBoardDataOverrideForTests`, alongside `BoardKeyOverrideForTests`, so a test
never has to construct `Main`); `ActiveWorkbookIdByBoard`'s persistence itself by
`UserSettingsTests`. Note the mockup's still-hardcoded entry list draws the
README's four-category vocabulary (Note/Cosmetic/Suspected/Confirmed, Pending/Fixed/Ruled out),
which is NOT the shipped `WorklogManager` model (Note/Cosmetic/Issue, Open/Closed) that every REAL
piece of this tab already uses — reconciling the two is part of making the entry list functional.

#### Schematics (`Tabs/Schematics/`)

The most complex tab: it renders schematic/PCB images with three overlay layers (component highlights,
an interactive KiCad trace/copper overlay, user-drawn polyline traces), hosts the component label
editor, and hosts the MiniPro IC-test panel.

**Adding a worklog entry goes straight to the full editor.** "Add worklog" in the top bar starts
area-marking mode; the drag that follows opens `WorklogEntryEditorWindow` on the drawn area
(`TabSchematics.Worklog.cs`'s `OpenNewWorklogEntryEditor` → `InitializeForNewEntry`). There used to
be a small "New fault" quick card in between, asking for a title, description, category, state and
the component checklist — every one of them a field the full editor also has, so reaching anything
else (links, work done, comments, photos, files) meant saving and immediately reopening the same
entry. The card was **removed outright** rather than kept as a shortcut, along with its markup, its
`AnchoredCardPlacementGeometry` corner-placement helper and the tests that pinned its fields down.

A NEW entry is held entirely in memory until Save (`thisIsDraftEntry` in the editor). That is the
one behavioural difference from editing a saved entry, and it exists so Cancel leaves nothing
behind: for a saved entry every sub-list change writes through to disk at once
(`PersistEntrySilently`), which for a half-made new entry would strand it in the workbook. Save
writes the whole record — sub-lists included — through `WorklogManager.AddEntryRecord`, which
**re-allocates the entry id at write time**: the editor reserves one up front from `PeekNextEntryId`
because its attachment folder has to be named after something, but a peek is not a reservation, so
an entry written meanwhile would otherwise give two entries the same number. When the id does move,
`AddEntryRecord` moves the draft's attachment folder with it; a cancelled draft's attachment bytes
are deleted instead (`DiscardDraftAttachments`, wired to Cancel *and* to `Closing`, since the
title-bar close does not go through Cancel).

**An oscilloscope capture can be filed straight into a worklog.** After a capture, the component
popup's existing "Saved image as [...]" banner carries an **"Attach image to worklog"** button
(`ComponentInfoWindow.OnAttachCapturedImageClick`). The banner was chosen deliberately over a mode
entered beforehand: the user is at the bench sweeping pin after pin, so a prompt per capture would
be intolerable while an ignorable button costs nothing. It is hidden entirely when
`UserSettings.EnableWorklog` is off (matching how the tab and the bar are gated) or when the board
has no workbook to file into, and it is cleared alongside the banner so it can never act on an
earlier capture's file.

**The WORKBOOK is never in question; the ENTRY is.** `ResolveActiveWorkbook` already settles which
workbook is active app-wide, so the only thing the user has to answer is which worklog - which is
why [WorklogAttachCaptureWindow](../Tabs/Worklog/WorklogAttachCaptureWindow.axaml.cs) is ONE modal
(image, workbook, worklog, comment) rather than a picker followed by the editor's Add photo dialog.
It resolves the choice and returns it; the caller performs the write, the same division
`WorklogAddPhotoWindow` keeps. The popup asks `Main.ResolveActiveWorkbookForBoard` (made `internal`
for this) rather than re-deriving "which workbook" a third time.

**The entry list is RANKED, and the first row is preselected** -
[WorklogAttachTargets](../Handlers/Data/WorklogAttachTargets.cs) (pure, unit tested). The rule is
deliberately just TWO levels: entries whose `ComponentLabels` contain the component being measured
come first, then everything else, **both bands in ascending id order**. Someone probing U8 while
working a fault on U8 gets a single Attach click. Closed entries are KEPT rather than hidden - a
board that comes back is re-measured against the entry describing the original repair.

**Ascending id, because this renders as a plain dropdown and a LIST has to look ordered.** The first
version sorted newest-first inside an open-before-closed band - defensible per criterion, and on
screen it produced `#2, #4, #3, #1`: three invisible criteria interleaving into what reads as no
order at all. Reported as exactly that. Entry ids are also what the board pills show (`#4`), so
counting order is the one ordering the user can already follow. **Open/closed is no longer a sort
level** for the same reason; both values already carry an always-visible pill, so they filter by eye.
The component match survives as the one level above id because it pays for itself - it puts the right
answer in the preselected slot.

**The two bands are NAMED by non-selectable headers inside the dropdown** - "Worklogs with U8 in
scope" and "All other worklogs" (`BuildGroupHeader`). The ordering was reported as illogical twice,
and a caption UNDER the box did not fix it: it explains an order the reader cannot see at the time,
and it is not on screen at the moment the list is open and the order actually matters. So the
grouping is stated where it is being applied. The headers are `ComboBoxItem`s with `IsEnabled` false
rather than data rows plus a `SelectionChanged` guard - a disabled item is skipped by mouse AND
keyboard, so a header can never become `SelectedItem`, whereas a guard lets the selection land on it
and then bounces, which flickers and fights arrow-key navigation. Because a header then occupies
index 0, the preselected row is set BY VALUE (`items.FirstOrDefault(item => item is EntryChoice)`),
never by index. Headers appear only when there IS a match to separate out: with one band, a lone
"All other worklogs" would name a distinction that is not being drawn.

**The matched heading names the component in bold inside brackets** - `Worklogs with [U6] in scope`
(`BuildComponentGroupHeader`), so the thing the grouping is keyed on is visible at a glance rather
than buried in a sentence. A `TextBlock` cannot mix weights within one `Text`, so it is built from
`Run`s - the identical reason `TabWorkbooks.Summary.cs` walks `WorkbookSummary`'s `Stat` parts. That
makes its `Text` null with the words in `Inlines`, so **a test reading only `Text` sees the heading
as blank**. The "All other worklogs" heading stays a plain string: nothing there needs emphasis, and
a `TextBlock` built to hold no bold run is just a heavier way to say the same thing. It also means
`Content` is a `string` for one header and a `TextBlock` for the other, so **anything distinguishing
headers from rows must match on the `EntryChoice`, not on "not a string"** - that shortcut picked the
bolded header up as a worklog row and read its outer padding as the indent.

**A header is faint (0.45 opacity) and OUTDENTED, with its worklogs indented under it.** At 0.75 it
still read as a selectable option and was reported as such - a ComboBox's rows carry no styling of
their own to contrast against. Opacity rather than a colour, because the theme defines no muted
-foreground key and dimming the inherited `Fg` stays correct in BOTH themes where a hardcoded grey
would fail one. The indent goes on the generated CONTAINER via `ContainerPrepared`: the rows are
plain records with no `Padding` of their own, and a `ComboBoxItem` style would also hit the headers,
which ARE `ComboBoxItem`s and set their own padding. **Any test for that indent must assert against
the Fluent theme's own 11px item padding, not against the header's 6px** - an un-indented row already
sits further right than the header, so the obvious comparison passes with the hook removed
entirely.

**"Create new worklog" opens the full editor on a draft with the photo already attached**
(`WorklogEntryEditorWindow.AttachCapturedPhoto`), rather than making the user save an entry and come
back for it - probing before anything is written down is how diagnosis starts. It must be called
AFTER `InitializeForNewEntry`, whose reserved id names the draft's attachment folder.

**That draft MUST be filed against a real schematic.** Both surfaces that draw worklog entries
filter by `SchematicName` - the Schematics tab's overlay (`RefreshWorklogEntriesList`) and the
Workbooks board pane - so an entry saved with a blank one is invisible on both, unreachable from the
board entirely. It shipped that way once and was reported: the worklog saved fine and then appeared
nowhere. `ComponentInfoWindow.ResolveSchematicNameForCapture` reads the schematic currently showing
on the Schematics tab (`TabSchematics.GetCurrentSchematicName`, made `internal` for this), which is
the board view the user is working against.

**The draft carries NO marked area, with `ShowMarkedArea` off**
(`WorklogEntryEditorWindow.SetShowMarkedAreaForNewEntry`) - it was born at the bench with a probe in
hand, not by dragging a rectangle. That is the supported parked-pill state: the entry shows as a
"#N" pill in the schematic panel's TOP-right corner rather than as a rectangle on the board. The
seam sets both the checkbox and the record, or the editor's own save would write back the markup
default of "ticked" and the entry would promise a rectangle it does not have.

**Ticking "Show marked area" on such an entry gives it a real, draggable square** -
[WorklogDefaultAreaGeometry](../Handlers/Geometry/WorklogDefaultAreaGeometry.cs) (pure, unit tested),
via the editor's `EnsureMarkedAreaExistsWhenShown`. Without it the tick left a ZERO-SIZED rect, which
draws as nothing or as a hairline and can never be grabbed and dragged into place - the entry looked
broken with no way to fix it from the UI. The square is placed in the board's **BOTTOM**-right
corner, the opposite corner from the parked pills, so a freshly placed area cannot be mistaken for
one of them while it is being moved. It only ever ADDS an area and never replaces one: an entry with
a drawn rectangle keeps it through any number of tick/untick cycles, since silently relocating a
user's own marked area would be worse than the bug being fixed. `IsUnset` tests a threshold rather
than comparing to zero - a rect that has been through a JSON round-trip or a hand edit can carry a
sliver that is still not grabbable.

**Both attach paths share one writer** -
[WorklogAttachmentWriter](../Handlers/Data/WorklogAttachmentWriter.cs), which the editor's own Photos
section now calls too rather than keeping its own copy. It carries four subtleties that are each
invisible when wrong and were each fixed once already: the id is allocated through
`AllocateAttachmentId` (which SKIPS ids whose bytes are already in the folder, since plain
`Max(Id) + 1` silently overwrites an orphan), the id is settled BEFORE the name that is built from
it, `DisplayOrder` is 0-based to match `ReorderAttachment`'s dense renumbering, and a failed persist
rolls the copied bytes back out. `AttachToEntry` re-reads the entry from disk rather than trusting a
caller's copy: `ShowDialog` does not block the dispatcher, so the full editor can be open on that
same entry and writing back a stale record would drop whatever it has since saved.

**Attaching refreshes through `Main.RefreshWorklogBar`**, the one funnel every worklog change
already passes through - deliberately NOT by poking `TabWorkbooks`, whose decoded schematic bitmaps
are tied to its attach/detach cycle and whose `OnDetachedFromVisualTree` comment warns about exactly
that re-entrancy. **The capture itself is written to the oscilloscope image folder before any of
this**, so cancelling the dialog, or a failed attach, costs the user the filing but never the
measurement.

**Links in user-typed text are clickable.** The workbook Note, the worklog Description, and the Work
done / Comment / Photo comment / File comment rows all render any web link in them as a clickable
run — `Handlers/Data/TextLinkFinder.cs` decides which runs are links (pure, unit tested), and
`Handlers/Theme/TextLinkRenderer.cs` turns those into styled `Run`s and opens the target through
`ExternalTargetLauncher`. Code-built blocks call `Apply`/`ApplySegments`; the editor's templated
rows use the `TextLinkRenderer.LinkText` attached property **instead of** `Text`, since a TextBlock
carrying both renders the Text and silently ignores the Inlines. Link marking and the Workbooks
tab's search highlighting are **merged in one pass** rather than applied in sequence: a search term
routinely lands inside a URL, and the two splits cut the text at independent places. Titles are
deliberately NOT linkified (`ApplyHighlightedText`'s `linkify` flag is opt-in) — a title is a
headline, not something to navigate to.

`TabSchematics` is one partial class split by area across the files below. **Find the right file here
before grepping** — the same header map is repeated in
[Tabs/Schematics/TabSchematics.axaml.cs](../Tabs/Schematics/TabSchematics.axaml.cs):

| File | Owns |
| --- | --- |
| `TabSchematics.axaml.cs` | Construction, one-time wiring in `Initialize`, fullscreen/splitter layout, the user-drawn trace colour palette, shared parse/theme helpers |
| `TabSchematics.Types.cs` | Private data types shared by the other parts (cache records, hover candidates, undo state, `EditableComponentHighlight`) |
| `TabSchematics.Viewport.cs` | Zoom, pan, the transform matrix, matrix clamping, content/viewport rects |
| `TabSchematics.Input.cs` | Pointer, wheel, gesture and keyboard handlers — these only dispatch |
| `TabSchematics.Thumbnails.cs` | Thumbnail list, selection, thumbnail bitmaps, drag-to-reorder |
| `TabSchematics.Highlights.cs` | Component highlight overlays, blink visuals, hover UI, on-schematic labels |
| `TabSchematics.LabelEditor.cs` | Label editor lifecycle, menu, apply/cancel, validation and save dialogs, search, undo/redo |
| `TabSchematics.LabelEditor.TestSeams.cs` | `...ForTests` seams letting headless tests drive the editor (see its header) |
| `TabSchematics.Worklog.TestSeams.cs` | `...ForTests` seams for the worklog area-drawing flow (see its header) |
| `TabSchematics.LabelEditor.Interaction.cs` | Label editor selection, resize handles, drawing, dragging, coordinate conversions |
| `TabSchematics.LabelEditor.Snap.cs` | Builds the snap context from tab state; the maths is `Handlers/Geometry/LabelEditorSnapGeometry` |
| `TabSchematics.KiCad.cs` | KiCad project load, board-label→net/reference mapping, selection sets, runtime cache scopes |
| `TabSchematics.KiCad.Panels.cs` | The "Important signals" and "Net connections" side panels |
| `TabSchematics.KiCad.Render.cs` | Draws the KiCad overlay, refresh scheduling, pin-1 marking |
| `TabSchematics.KiCad.RenderCache.cs` | Builds/caches per-net PCB render nodes and connected-segment chains |
| `TabSchematics.KiCad.Geometry.cs` | KiCad world ↔ screen mapping, world bounds, curve sampling, zone polygon geometry |
| `TabSchematics.KiCad.HitTest.cs` | Hover hit-testing, hit-test caches, hover throttling, trace hover mode UI |
| `TabSchematics.KiCad.Calibration.cs` | Interactive KiCad trace calibration mode; the maths is `Handlers/Geometry/KiCadCalibrationGeometry` |
| `TabSchematics.Settings.cs` | Board-level and global setting rows, and restoring them per board |

Supporting classes in the same folder are ordinary (non-partial) types: `KiCadOverlayRenderControl`,
`SchematicHighlightsOverlay`, `PolylineManagement` (user-drawn traces; reaches into
`TabSchematics.schematicsMatrix`), `HighlightSpatialIndex`, `SchematicThumbnail`, `ComponentInfoWindow`,
`ComponentLabelEditorOverlay`, and `IcTestPanel`.

### Data layer (`Handlers/Data/`)

- `DataManager` (static) — resolves the data root (default location or `--data-root=`), loads the master
  Excel workbook, and drives sync against the online checksum manifest.
- `BoardData` / `BoardDataReader` / `BoardDataWriter` — the schema (in [Handlers/Data/BoardData.cs](../Handlers/Data/BoardData.cs))
  and read/write logic (via EPPlus) for per-board Excel files: schematics, components, component
  images/highlights, local files, links, credits, and KiCad signal mappings.
- `DataValidator` — validates board/contribution data.
- `UserSettings` — JSON-persisted user preferences (theme, window placement, MiniPro path override, etc.).
- `KiCadProjectData` / `KiCadProjectLoader` / `KiCadRawProjectLoader` — parse raw KiCad PCB/schematic
  files into a normalized bundle so the Schematics tab can highlight matching copper and wire geometry.
- `Logger` — writes to the app's log file in the AppData folder alongside settings.
- `ComponentListBuilder` — the main window's component list: region filter, category filter, search,
  and the `ComponentListItem` rows it produces. Also owns `IsSupportedKiCadRawFile`.
- `ComponentImageQueries` — which component images and entries the popup shows for a region, and
  which of them carry an oscilloscope baseline.
- `OverviewHtmlBuilder` / `OverviewModels` — bill-of-materials grouping and the printable HTML for
  the Overview tab, plus the `OverviewRow`/`OverviewLink` models it renders. **The Overview AXAML
  binds these through an `xmlns:data="clr-namespace:Handlers.DataHandling"` mapping** — if you move
  or rename them, update `TabOverview.axaml` too.
- `ContactLinkFormatter` — classifies a contributor's contact string and builds its href.
- `TextLinkFinder` — which runs inside a user-typed free-text field are web links, so the UI can
  render them as clickable. Deliberately conservative: only `http://`, `https://` and `www.` at a
  word boundary count. A bare `example.com` does NOT, because that is the shape repair prose
  collides with — `74LS08.pin3`, `5.0V`, `notes.txt` — and a false link is worse than none. Pure
  string work, unit tested; the Avalonia half is `Handlers/Theme/TextLinkRenderer.cs`.

### Content (`Assets/Data/`)

All hardware reference content is data, not code: organized as `<Manufacturer>/<Hardware>/<Board>/...`
(e.g. `Commodore/C64/250407/`), with per-manufacturer `Shared files` folders and a top-level
`Generic shared files` folder for cross-manufacturer component images. A master workbook
(`Classic-Repair-Toolbox.xlsx`, sheets `Hardware & Board` and `Oscilloscope`) lists all available
hardware/boards and oscilloscope SCPI command sets; each board also has its own Excel file matching the
`BoardData` schema. This content is what `DataManager` syncs from `classic-repair-toolbox.dk` at runtime
using SHA-256 checksums, independent of app releases — adding new hardware is primarily a data
contribution, not a code change.

### External integrations (`Handlers/`)

- `Online/OnlineServices` — fetches the checksum manifest and syncs changed data files.
- `Online/UpdateService` — checks/applies application updates via Velopack against GitHub Releases
  (`GitHubOwner`/`GitHubRepo` in `AppConfig`).
- `Oscilloscope/IScopeClient` — the "talk to the scope" seam (`SendAsync`/`QueryLineAsync`/
  `QueryBinaryBlockAsync`), the same idea as `IMiniproRunner`. The Oscilloscope tab's sequencing
  takes this interface, so it can be tested against a fake with no scope on the network; the real
  `ScopeScpiClient` below it stays an untested I/O boundary.
- `Oscilloscope/ScopeScpiClient` — raw SCPI-over-TCP client implementing `IScopeClient`; `ScopeCommandPalette`/`ScopeCommandResolver`/
  `ScopeValueMapper` translate the data-driven `OscilloscopeEntry` command strings (per brand/model, from
  the master workbook) into actual scope interactions for baseline capture.
- `MiniPro/` — integration with the MiniPro USB IC programmer for in-app IC testing. `IMiniproRunner` is
  the abstraction; `MiniproProcessRunner` spawns the bundled `minipro.exe` (streamed, cancellable output),
  `MockMiniproRunner` simulates it without hardware attached. `IcTestCatalogue`/`IcTestModel`/`IcTestService`
  manage the IC test definitions. Only Windows (`win-x64`) currently bundles the `minipro` binary — see the
  conditional `ItemGroup` in [Classic-Repair-Toolbox.csproj](../Classic-Repair-Toolbox.csproj).
- `Security/ExternalTargetLauncher` — the only sanctioned way to open an external link or local file from
  the UI; it restricts targets to HTTP/HTTPS/mailto URIs or local paths that resolve inside the current
  data root *and* carry a document/image/data file extension from its allowlist (it hands files to the
  OS shell, which would run a `.exe`/`.bat`/`.lnk` instead of displaying it), rejecting anything else.
  Use this rather than shelling out directly when opening user/data-supplied links or files.

**QuestPDF** (the workbook PDF export) is the app's one non-Avalonia UI dependency worth knowing
about. Two things about it:

- **Its licence is a condition on this project, not just a package reference.** The Community
  licence QuestPDF is used under is free for individuals and for organisations under $1M USD annual
  revenue — which this project is. An organisation above that threshold shipping a fork would need
  its own commercial licence from QuestPDF.
- **`QuestPDF.Settings.License` must be set before it generates anything**, or the first export
  throws. It is set once in `App.OnFrameworkInitializationCompleted`, not at the export call site,
  so a missing line fails at launch in development rather than in a user's hands.

### Contribution webserver (`Assets/Webserver/`)

A working copy of the PHP deployed at `classic-repair-toolbox.dk/app-contribution/` — the server
side of the Contribute tab, split into two entities: `api/index.php` receives the uploads the app
posts to `AppConfig.ContributionUploadUrl`, and `review/index.php` + `review/functions.php` are the
review queue (admin-only, IP-restricted via `review/.htaccess`) that diffs each submission against
the live server data and merges or rejects it. `api/index.php` requires `review/functions.php` for
its shared helpers. It is deployed by hand and never ships with the app (Assets are whitelisted
per-file in the csproj).

**The payload is a two-sided contract.** `ComponentContributionPayload` in
[Tabs/Contribute/ComponentContribution.axaml.cs](../Tabs/Contribute/ComponentContribution.axaml.cs)
is what `review/functions.php` parses, and the PHP deliberately mirrors the app's Excel-reading
rules (case-insensitive headers, exact board-file resolution, `.json`-beside-the-Excel highlights,
the `# Revision date:` marker). When you change either side, change the other in the same
sitting, and run the PHP suite too: `php Assets/Webserver/tests/run-tests.php`
(~1s, no webserver needed — PHP is installed on this machine; the tests sit outside
`app-contribution/` so that folder stays an exact mirror of what is deployed). **When the payload contract
changes, also bump `$minimumContributionVersion` in `api/index.php`** to the first released
version carrying the new contract — older apps are rejected at upload with an update-required
message. Full details, file inventory (including which files are legacy) and known caveats:
[Assets/Webserver/app-contribution/README.md](../Assets/Webserver/app-contribution/README.md).

## Release process

Versioning lives in [Classic-Repair-Toolbox.csproj](../Classic-Repair-Toolbox.csproj)
(`AssemblyVersion`/`InformationalVersion`) — bump `InformationalVersion` there before releasing, since
that is the only place a release version is entered. Releases are made by hand from the GitHub Actions
tab — run [.github/workflows/build-and-release.yml](../.github/workflows/build-and-release.yml) with no
inputs — **never by pushing a tag**; that trigger was deliberately removed. Its first job reads
`InformationalVersion` straight from the csproj (via `dotnet msbuild -getProperty:InformationalVersion`)
and derives pre-release status from it (a version containing `-`, e.g. `2.5.0-beta.1`, is a pre-release;
a bare `X.Y.Z` is not) — every other job consumes that one job's output rather than re-deriving it. It
then runs the test suite and stops there if it is red, then a CodeQL scan, then builds/signs/packages
(Velopack) self-contained builds for win-x64, linux-x64, osx-x64 and osx-arm64, and finally publishes a
GitHub Release using [CHANGELOG.md](../CHANGELOG.md) as the release body. The tag is created by that
last step, so a failed run leaves nothing behind to clean up and the same version number can simply be
re-run once the fix is pushed. That release body is written by hand and is off-limits to Claude —
see [Hands off CHANGELOG.md](#hands-off-changelogmd).
