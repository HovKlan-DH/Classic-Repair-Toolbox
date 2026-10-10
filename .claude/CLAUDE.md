# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

It holds the RULES and the TRAPS - what must stay true, and what broke before. The history of how
each came about is in git; a dated "(owner decision, 2026-...)" marks something the project owner
settled, which is not to be re-argued without asking.

## What this is

Classic Repair Toolbox (CRT) is a cross-platform (Windows/Linux/macOS) Avalonia desktop app that helps
hardware enthusiasts diagnose and repair vintage computers. It presents schematics, component data,
oscilloscope baselines, and interactive KiCad traces for a curated set of hardware (mostly Commodore,
plus Amstrad and ZX Spectrum boards). It also drives real test equipment: SCPI oscilloscopes over TCP
and a MiniPro USB IC programmer/tester. Users improve the data from inside CRT (Contribute and Drafts
tabs); maintainers review it in CRT's Maintainer tab, and the contribution service (CRT.Server)
publishes it to the BETA data and then to the stable data.

## One change, every side of it

**This is one system in four parts, and a change to shared behaviour is not done until every
part that shares it has been changed in the SAME session.** The parts are the CRT desktop app
(`src/CRT.App/`, including the Maintainer tab and, in `src/CRT.App/Controls/`, the controls the
Drafts and Maintainer tabs share), the contribution service (`src/CRT.Server/`), the shared library
they both reference (`src/CRT.Data/`), and the board DATA itself (`Assets/Data/`, and the published
trees the server writes).

**"Board", never "system"** (owner decision, 2026-10-09: "rename everything, code too"). One entry
of the drop-down lists - "Commodore / C64 / 250407" - is a BOARD, in every word a person reads, in
code, routes (`/api/review/boards/...`), JSON fields (`boardId`), the database (`boards`, `board_id` -
migration 0020) and the Wiki. "System" is left only where it means something else (the operating
system, a file system, systemd, a video system, SCPI's "system error") and where renaming would lose
data: names ALREADY STORED keep their old spelling on purpose - a draft marker's `SystemKey` and
`NewSystem`, a receipt's `SystemId`, the settings' `maintainerLast...SystemId`, `BoardView.SystemId`
(older CRTs still send it), the retired `system.json`, and every applied migration. Never "fix" those.
A name the DATABASE stores inside JSON is the exception that moved: 0020 rewrites
`submission_changes.changes_json`'s `IsNewSystem` to `IsNewBoard` - renaming a member of a type the
server keeps as JSON needs such a rewrite in the same migration (`BoardNamesMigrationTests`).

**The compiler covers less of this than it looks.** Both applications reference `CRT.Data`, so a
changed signature fails the build everywhere at once. The dangerous surfaces are the ones it CANNOT see:

- **HTTP routes and JSON field names.** A renamed endpoint or JSON property compiles on both sides
  and fails at runtime, in the user's hands. **The review API's bodies are CRT.Data's
  `ReviewApiContract` records**: a request is one record both ends use, an answer is a record the
  server returns and the Maintainer tab parses, and CRT.App.Tests' `ReviewWireContractTests` puts each
  answer through the server's JSON settings and the real parser. A new route gets its records there
  and a case in that test.
- **The workbook schema.** `BoardWorkbookSchema` names the columns the reader reads and the writer
  writes. A column added on one side only is silently blank on the other.
- **On-disk formats.** The data tree's layout, the JSON sidecar, the draft folder, the sync manifest.
- **Vocabulary the user reads.** A state name, a status word or a date format that appears in more
  than one place must change everywhere at once (CRT.Data holds the shared wordings:
  `MaintainerScreenWording`, `ConfigurationWording`, `OneSubmissionInBeta`, `StablePublishing`, ...).

**So before finishing a change to anything shared, ask which of the four parts also speak it, and
change them together.** If one genuinely cannot be changed yet, say so in the turn summary.
**A shared change needs a test that would fail if only one side moved** - a test per side, each
asserting its own half, passes happily while the halves disagree.

**The website's own pages are not part of this project** (owner decision, 2026-10-05): the old PHP
pages on classic-repair-toolbox.dk are retired and are not mentioned in code comments or documents.
The service still answers the old ADDRESSES that CRT 2.x posts to (see "Installed CRTs keep
working"). `Assets/Webserver/` is a git-ignored local mirror of the site - never delete a file there
to "retire" it, it has no git copy.

## Hands off CHANGELOG.md

**Never create, edit, rewrite, reformat or delete [CHANGELOG.md](../CHANGELOG.md) unless the
project owner explicitly asks for it in that message.** It is written by hand, in the project owner's own
words, and it is the body of every GitHub Release - an "improvement" there is not a small edit, it
is words the project owner never wrote going out under their name.

This holds even when a change would normally warrant a changelog entry, and even when the file
already has uncommitted edits in it (those are the project owner's, in progress). Do not touch it as a
"finishing touch" on a feature, do not tidy its formatting, and do not stage, commit, revert or
`git checkout` it. If you think an entry is needed, say so in your summary and let the project owner
write it. "Update the changelog" from the project owner is the only permission - and it covers that
one request, not the rest of the session.

## Code reviews: every finding carries a severity

**Every finding in a code review starts with its severity**, so the project owner can see at a
glance whether a review is still finding serious problems or only tidying (owner request,
2026-09-29). This covers `/code-review` at every effort level and any review asked for in words.
`/security-review` keeps its own exploitability grading, and `/simplify` is not a review (it applies
its fixes) and carries no severities.

- **The prefix.** The first thing in EVERY summary field the report has is the level in brackets:
  `[Critical]`, `[High]`, `[Medium]` or `[Low]`, for example `[High] A rejected submission's mail
  is never sent`. When findings go through `ReportFindings`, that is both `short_summary` and
  `summary`; its `short_summary` is capped at 60 characters INCLUDING the prefix (`[Critical] ` is
  11), so keep the claim short rather than truncating it. When findings are returned as a JSON
  array, it is each `summary`.
- **The tally.** After the findings - in the reply text below the JSON block, or after the
  `ReportFindings` call, never inside either - one line counts the findings that survived
  verification: `Severity: 0 Critical, 3 High, 4 Medium, 4 Low`. With no findings it reads
  `Severity: none found`.
- **The stop signal**, on its own line under the tally. When the worst finding is Medium or Low, or
  there is none: `Worst finding is Medium - another full review of this change is unlikely to pay
  off.` (with the real level, or `No findings - ...`). When there is a Critical or High finding: `Worst
  finding is High - review again once it is fixed.`

| Severity | Means, in this project |
| --- | --- |
| **Critical** | Data loss or corruption (a user's drafts or workbooks, the published data trees, the database), a security hole someone can actually reach and exploit, or a crash on a path people use every day. Fix before anything is released or deployed. |
| **High** | Wrong behaviour a user, contributor or maintainer WILL meet: a stuck state that never clears, a publish or submission that fails, a mail not sent, a crash on a rarer path. |
| **Medium** | Wrong only in an edge case, a real performance cost, text that misleads, a latent bug that needs another change to go off, or a hardening gap with no known way to exploit it. |
| **Low** | Duplication, logic in the wrong place, naming, a missing test, a log level - nothing anyone would notice in use. |

**Judge by what would ACTUALLY happen, not by what the kind of code sounds like**: a thread-safety
defect nobody can reach is Low, and a one-line wording error on a refusal the maintainer depends on
can be High. **When unsure between two levels, pick the LOWER one** unless the failure scenario
shows concretely how the higher one happens - rounding up on doubt would make every review look
serious and the stop signal would never fire. Rank findings most severe first.

## Documentation: the Wiki, mirrored in `Assets/Wiki/`

The published documentation is the **GitHub Wiki**. Its source of truth is
[Assets/Wiki/](../Assets/Wiki/) - one `.md` file per Wiki page, named **exactly** as the page is.
**The Wiki must be fully up to date with the app at every release** (owner request, 2026-10-05).

### Know who is reading: hobbyists, not professionals

**CRT's users are hobby hardware enthusiasts repairing their own vintage computers.** Some are
deeply technical; most are ad-hoc "I want to see what this is" users. As a rule of thumb **nobody
is doing this for a living and nobody earns significant income from it.**

So documentation and user-visible strings must never assume a job, a workplace, customers, clients,
billing or a trade. Write "the country you live in", "the repair", "share the write-up", "export it".
Where a page has to name the thing a workbook represents, it is **a repair**, not a job. If a
phrasing is genuinely ambiguous, ask rather than guess. This does not touch the app's own domain
nouns (a workbook holds worklogs), ordinary English "job" meaning a task, or code comments.

### How the mirror works

**Nothing publishes automatically.** The project owner copies changed files to the Wiki by hand.
Editing a file here changes what the page *should* say; the page itself changes only when they paste
it. Never tell the user a Wiki page has been updated - say which files changed and that they are
ready to be pasted.

- **Filenames ARE page names**, capitals included. Never rename one without the Wiki page being
  renamed to match.
- **These files use WIKI syntax, not repository syntax**, because they are pasted: internal links
  carry no path and no extension (`[Board JSON](Board-JSON)`), and images stay as `<img>` tags
  pointing at their `user-attachments` URL. `Assets/Wiki/images/` holds reference copies only.
- **[Assets/Wiki/!sync-status.md](../Assets/Wiki/%21sync-status.md) is GENERATED - never hand-edit
  any of it or add anything to it.** The `Stop` hook [hooks/wiki-sync-status.sh](hooks/wiki-sync-status.sh)
  rewrites it every turn: ONE table of the pages waiting to be pasted, exactly two columns - the FILE,
  and where that page lives in the live Wiki - and nothing else (owner instruction: no "changed
  since" columns, no "what changed" prose; explain changes in your turn summary instead). The Wiki
  location is DERIVED from [_Sidebar.md](../Assets/Wiki/_Sidebar.md) at hook run time (a page listed
  twice keeps its first trail; one not listed falls back to `Home`) - after restructuring the sidebar,
  re-run the hook and glance at the trails. The comparison is on CONTENT against the blob hashes in
  [.claude/wiki-synced.tsv](wiki-synced.tsv), not on commits - pages are pasted while uncommitted.
- **When the project owner CONFIRMS a paste**, run [hooks/wiki-mark-synced.sh](hooks/wiki-mark-synced.sh)
  with those page names (or `--all`). Never stamp a page they have not confirmed.
- **When asked for a sync diff**, the table IS the answer; say what changed in each page in your reply.

**Documentation changes ship in the same commit as the code they describe.** The `Stop` hook
[hooks/remind-wiki-mirror.sh](hooks/remind-wiki-mirror.sh) maps code paths to the pages that
document them and names any whose code changed while the page did not. It WARNS (whether a change
is user-visible is a judgement call). When a page starts documenting new code, add it to the `MAP`.

**Six Wiki pages are opened by buttons in the shipped app**, across seven buttons:
`Workbooks-tab`, `MiniPro-programmer` (two buttons), `Controlling-oscilloscope-with-keyboard`,
`Synchronize-oscilloscope`, `Maintainer-tab` and `View-boards-from-online-source`. They are named
once, in `AppConfig` (`WikiPage*` plus `WikiPageUrl`). Renaming or deleting one breaks a button in
builds already installed, which no update can fix, so `WikiHelpPageNamesTests` asserts each name
still matches a file in `Assets/Wiki/`. Add an entry there for any new help button.

**Some old Wiki pages are deliberately not mirrored** (orphans superseded by newer pages) - see
[Assets/Wiki/README.md](../Assets/Wiki/README.md). Do not add them back without being asked.

## Build, test, run, publish

- **Run the tests: `dotnet test Classic-Repair-Toolbox.slnx`** (about 2 minutes in Release; needs
  no hardware, network or display)
- Build (Release, matches CI): `dotnet build Classic-Repair-Toolbox.slnx -c Release`
- Run/iterate: VS Code task `watch`, F5 in VS Code (`.vscode/launch.json`), or Visual Studio
- Self-contained publish: `dotnet publish -c Release -f net10.0 -r <rid> --self-contained` with
  `<rid>` one of `win-x64`, `linux-x64`, `osx-x64`, `osx-arm64`
- **The build configuration does not change what the app does.** Nothing may be gated on `#if DEBUG`;
  the one use left is `AppConfig.IsDebugBuild`, which is only logged. RELEASE matters for timings and
  warnings-as-errors, not behaviour.
- **Command-line switches**, all parsed once at startup the same way (case-insensitive, surrounding
  quotes stripped, first match wins, unknown arguments ignored):
  - `--data-root=<path>` - another `Data` folder (default: an AppData folder that survives Velopack
    updates). `DataManager.ResolveDataRoot`.
  - `--workbooks-root=<path>` - another folder for worklog workbooks. `WorklogManager.ResolveExplicitWorkbookRoot`.
  - `--drafts-root=<path>` - another folder for drafts. `DraftManager`.
  - `--simulate-update[=<version>]` - offer a fake update (default `99.0.0`), fake the download and
    skip the restart. `Handlers/Data/SimulationOptions.cs`. Active in RELEASE too, on purpose; the log
    shouts about it and the banner says `(simulated)`. Set for F5 and `watch`.
- **To skip the online data sync while iterating, untick "Check for new or updated data at
  application launch" in the Configuration tab.** It is a normal user setting.
- **macOS data folder**: .NET 8+ maps `LocalApplicationData` to `~/Library/Application Support` on
  macOS (Linux: `~/.local/share`). Documentation must say so.

## Tests

**THREE test projects**, all xUnit and all run by the one `dotnet test Classic-Repair-Toolbox.slnx`:
[tests/CRT.App.Tests/](../tests/CRT.App.Tests/) (the app's `Handlers/`, the Maintainer tab's logic in
`Maintainer/`, and the headless UI tests in `Ui/`), [tests/CRT.Data.Tests/](../tests/CRT.Data.Tests/)
and [tests/CRT.Server.Tests/](../tests/CRT.Server.Tests/). None needs an oscilloscope, a MiniPro, a
display, a database or a network. Read real counts off a run; do not quote them from here.

**The framework is xunit v3 (4.x), as the `xunit.v3.mtp-off` package, on VSTest** - so `dotnet test`,
CI, the release gate, the Stop hook and `coverlet.collector` coverage all work as on xunit 2. Moving
to Microsoft Testing Platform is a change to all of those at once, not a package bump. **xUnit1051**
(pass `TestContext.Current.CancellationToken`) is in `NoWarn` in all three projects.

**v3 RANDOMISES test order on every run**, so a test leaking state into whichever test follows fails
(or hangs) at random. The run header prints the seed, and running the test executable directly
replays that order and names a hang:
`tests/CRT.App.Tests/bin/Release/net10.0/CRT.App.Tests.exe :N -reporter verbose -longRunning 30`.

**Performance, measured - keep both:** every `ExcelPackage` is made by CRT.Data's
`EpplusLicense.NewPackage`/`OpenPackage`, which turn off EPPlus's forced GC on dispose (it was ~85 %
of a 10-minute suite, and cost CRT on the UI thread per board read) - `EpplusPackagesTests` fails on
`new ExcelPackage(` elsewhere in `src/`. A test that WRITES a board workbook as set-up goes through
`CachedWorkbooks.Write` (saves under test stay real; `CachedWorkbooksTests` holds the key complete).
For the next slowdown, measure before guessing: run a class alone against its time in the full run,
or take a sampled trace (`dotnet-trace collect --profile dotnet-common,dotnet-sampled-thread-time --
<absolute path to the test exe>`) and count busy samples per method.

**Tests are part of the change, not a follow-up. These rules are not optional** (they are cited by
number elsewhere - keep the numbers):

1. **ALWAYS write tests for logic you add.** Any new function that parses, maps, validates,
   compares, or does geometry/maths arrives *with* its tests in the same change - never "tests can
   come later". If it is pure, it goes in `Handlers/` (see `Handlers/Geometry/` for logic pulled out
   of a tab) and it gets a test file. Do not ask whether to add tests; add them.
2. **ALWAYS update the existing tests when you change covered behaviour.** If you touch a class that
   has tests, updating them is part of that same change. Never leave a red suite, never defer the
   update, and never delete a test to make a change pass.
3. **ALWAYS run `dotnet test Classic-Repair-Toolbox.slnx` before reporting any code change as done.**
   "It compiles" is not a completed change. Report the real result, including failures and counts.
   **This rule is machine-enforced:** the `Stop` hook in [settings.json](settings.json) runs
   [hooks/require-green-tests.sh](hooks/require-green-tests.sh) when you try to end a turn, and a
   red suite blocks the handover with the failing tests fed back to you. It builds Release and only
   fires when `.cs`, `.csproj`, `.axaml` or `.slnx` files changed. GitHub runs the suite again on
   every push ([.github/workflows/build-and-unittest.yml](../.github/workflows/build-and-unittest.yml)),
   and a red suite blocks releases too.
4. **A failing test is a question, not an obstacle.** Decide whether the behaviour change was intended.
   If it was, update the expectation and say so explicitly in your summary. If it wasn't, fix the code.
   Never edit an assertion just to get to green, and never weaken one (e.g. loosening a tolerance or
   dropping a case) to avoid understanding a failure.
5. **When you fix a bug, first add the test that fails because of it.** Then fix the code and show
   the test going green. That test is the thing that stops the bug coming back.
6. **Never write a test that needs hardware, a network call, a display, or that starts a process.**
   Headless UI tests do not breach this - Avalonia's headless platform needs no display - but they
   must go through `UiTest.Run(...)`. `ExternalTargetLauncher`'s accept path calls `Process.Start`,
   so its tests exercise the private containment predicates by reflection instead. Follow the same
   principle for anything else with real-world side effects. If logic is untestable because it is
   welded to a control, extract the logic (rule 1) rather than giving up on testing it.
7. **These are characterisation tests** - they pin down what the code does today so future changes
   cannot alter it silently. Several encode deliberate quirks (relative-tolerance value matching, "no
   vector grid + a successful summary IS a pass", micro sign vs Greek mu, the dead `region` argument
   in `BoardDataWriter`). Read the comment before assuming a test is wrong.
8. **NEVER hard-code a Windows path in a test** (`@"C:\Drafts\..."`, `"D:\\data"`). GitHub runs the
   suite on Linux, where a backslash is not a separator, so such a test passes here and fails after
   the push. Build paths with `Path.Combine` (from `Path.GetTempPath()` when one must be rooted). A
   literal that is safe on purpose carries a `// windows-path-literal: <why>` comment on its line or
   up to three lines above. **Machine-enforced:** `TestPathLiteralTests` fails on any unexplained one.

**Tests do not read Markdown documents** (owner decision, 2026-10-05: "only code should have test
cases"). The two exceptions the owner kept: `ServerVersionTests` (VERSION.md's newest row equals the
server's version) and `WikiHelpPageNamesTests` (the help buttons' pages exist). The systemd unit is a
shipped file, `src/CRT.Server/crt-server.service`, precisely so `ServerExitCodesTests` can read it.

**Writing them:** one test file per class under test, named `<ClassName>Tests.cs`. Give each test a
sentence-shaped name saying what must hold (`A_faulty_chip_fails_and_names_the_failing_pin`), and put
a comment above anything non-obvious explaining *why* it matters. Use `TempWorkspace` for filesystem
work. Cover the failure and edge cases, not just the happy path: blank input, malformed contributed
data, wrong region, wrong board side, locale-specific number formats.

**No case lists before the code** (owner decision, 2026-10-03). Do not send numbered lists of test
cases to confirm before writing code. Implement the code and its tests in one go, and ASK when
something is genuinely the project owner's to decide - a few pointed questions, with the answer you
would propose.

**Coverage**: no percentage is recorded here on purpose. CI computes it on every push and renders the
merged total into the job summary ([coverage-summary.sh](../.github/workflows/coverage-summary.sh)) -
reported, never enforced. Locally: `dotnet test Classic-Repair-Toolbox.slnx --collect:"XPlat Code
Coverage"` writes ONE report per test project; merge them with ReportGenerator before reading a
total, and always quote the denominator and the build configuration.

### `Handlers/Geometry/` and the other extractions - pure logic out of the UI

**When you add pure logic to a tab, put it in `Handlers/` instead** - `Handlers/Geometry/` for maths
and geometry, otherwise the `Handlers/` folder for its area. Those classes may use Avalonia's
`Point`/`Rect`/`Matrix` value types but never touch a control, so they test with no display.
`public` and `internal` are both fine (tests reach `internal` through `InternalsVisibleTo`); do not
widen a type for consistency's sake. **Before writing a helper in a tab, grep `Handlers/` for it** -
several duplicates arose from re-implementing helpers that already existed.

The sweep is done: what remains in `Tabs/` and `Main/` was checked and is genuinely UI-bound, so do
not go looking for more to extract there. Two traps worth knowing in extracted code:

- **KiCad calibration (`KiCadCalibrationGeometry`): mirroring is not a flag - it IS the edge
  ordering** (a horizontally flipped board is `Left > Right`). That is why the box is four doubles,
  not a `Rect` (which would normalise the flip away), why arithmetic that needs ascending edges goes
  through `WithNormalisedEdges`, and why `ApplyDrag` does not stop an edge crossing its opposite.
  `RemapDragModeForFlip` is asserted entry by entry, plus the involution property.
- **Label editor snapping (`LabelEditorSnapGeometry`)** takes a `LabelEditorSnapContext` built at the
  call site (`TabSchematics.LabelEditor.Snap.cs`) - resolve UI reads to plain values and hand them
  over, the pattern to copy. The drag mode comes from that context and nowhere else.

### Test seams - use these, do not work around them

- `UserSettings.LoadFrom(path)` - tests point it at a temp file. **Never call `UserSettings.Load()`.**
  `LoadFrom` on a MISSING file keeps the previous settings, so write `{}` first when a test needs a
  clean slate.
- `DataManager.LoadFrom(dataRoot, workbookName)` - **never call `DataManager.InitializeAsync()`.**
- `Logger` writes nothing until `Logger.Initialize()` - **never call it from a test.**
- `WorklogManager.LoadFrom(root)` - **never call `WorklogManager.Load()`.** `TabWorkbooks.BoardKeyOverrideForTests`
  and `CurrentBoardDataOverrideForTests` stand in for `Main`.
- `ReviewSessionStore.InitialiseAt(path)` - **never call `ReviewSessionStore.Initialise()`** (it
  resolves the real file, which may hold a live maintainer token); point it back at nothing
  (`InitialiseAt(string.Empty)`) when done. Its tests share the `"ReviewSessionStore"` collection.
- `TabMaintainer.UseRememberedChoices(...)`, `UseRememberedSelections(...)`, `UseTabBadge(...)` - how
  `Main` hands the Maintainer tab its settings; a tab a test builds never touches `UserSettings`.

Tests that mutate shared statics live in their collections - `"UserSettings"`, `"DataManager"`,
`"BoardData"` (`BoardDataReader`'s static cache), `"ReviewSessionStore"`, `"HeadlessUi"`. **Any new
test class touching one of these statics must join its collection.** Collections do NOT run in
parallel with each other (`xunit.runner.json`, `"parallelizeTestCollections": false`) - keep that,
it removed a whole class of race.

### Headless UI tests

`tests/CRT.App.Tests/Ui/` builds the tabs, windows and `Main` itself through Avalonia's headless
platform - no display, so it runs on CI. `TestAppBuilder.cs` is a `CRT.App` subclass whose
`OnFrameworkInitializationCompleted` is deliberately empty; `UiTest.cs` runs a body on the UI thread.
**Anything touching a control must go through `UiTest.Run(...)`/`RunAsync(...)`.**

- **`UiTest.RunAsync` hands the test back OFF the dispatcher thread, explicitly**
  (`ContinueWith(..., TaskScheduler.Default)`). A plain `await` continued on the dispatcher thread, and
  the next synchronous `UiTest.Run` then queued its body to the thread it was blocking - the whole run
  hung. `UiTestTests` pins it. Do not "simplify" it.
- **The Application is built ONCE per assembly** (`AvaloniaTestIsolationLevel.PerAssembly`). The
  default `PerTest` rebuild threw intermittently before a test's body ran; `PerAssembly` made it much
  rarer but not gone - an empty "Error Message" on a headless UI test is that; re-run with
  `--logger trx` to confirm. **Its condition: no test may mutate the shared `Application`** - no
  `RequestedThemeVariant` assignment, no `Styles.Add`; read theme resources with an explicit variant.
  `UiTest.Run` fails a body that changes the theme; a control that applies a theme takes a seam
  (`TabConfiguration.ApplyThemeOverrideForTests`).
- **Do NOT add `Avalonia.Headless.XUnit`** to get `[AvaloniaFact]` - its adapter is built against an
  older xunit v3 extensibility API, untried against the 4.x this suite runs. `UiTest` drives the same
  session API directly.
- **Know what these catch.** The XAML compiler already fails on a renamed `x:Name` and a broken
  `avares://` path, while a missing `StaticResource` is silently tolerated at runtime - so construction
  tests guard constructor logic, not markup. Prefer interaction tests that assert observable state
  (`HeadlessWindowExtensions`: `MouseDown`, `KeyPress`, `MouseWheel`); a test reading drawn colours
  needs a SHOWN window.
- **A `TextBlock` with mixed runs has `Text == null`** and its content in `Inlines` - a test reading
  only `Text` sees it as blank.

**Deliberately not covered**: rendering and layout correctness (verified by running the app); I/O
boundaries (`OnlineServices`' network half, `UpdateService`, `ScopeScpiClient`, `MiniproProcessRunner`,
`DataManager`'s sync half, the `MySql*` stores) - test the abstraction below each. Two halves ARE
covered because they are trust boundaries: `OnlineServices`' manifest-validation predicates (by
reflection) and `UpdateChannelFilter` (which release stages a user opted into).

## Code layout conventions

- **No MVVM.** UI logic lives in `.axaml.cs` code-behind.
- **Large controls are split into partial-class files** named `<Class>.<Area>.cs`; keep any one file
  under ~1,500 lines. **Every partial file opens with a header comment** saying what it owns, and the
  `.axaml.cs` part carries the file map for the whole class - update it when you add or move a part.
  **Fields are declared in the part that owns them.**
- **Pure logic does not belong in a tab** - see `Handlers/Geometry/` above.
- **`Controls/` holds the controls more than one tab or window uses** - the table editor
  (`Controls/BoardTable/`), `BusyOverlay`, `ListRowDrag`. **They touch nothing of CRT's
  orchestration - no `Main`, `DataManager`, `UserSettings` or `Logger`**; a host hands them what they
  need, and CRT.Data's `CrtLog` is their logging seam. `SharedControlsIndependenceTests` enforces it.
  A control used by one tab stays in that tab's folder.
- **No button label has "..."** (owner rule, 2026-10-03). A button says what it does in plain words,
  whether or not it opens a window. Status lines and placeholders are not buttons. **Machine-checked:**
  `ButtonLabelTests`. A Wiki page naming a button uses the same label.
- **Everything only the administrator can use carries Font Awesome's padlock** (owner rule,
  2026-10-04): `<TextBlock Classes="AdminLock" Text="&#xf023;" />` right after its name, NO tooltip,
  and the entry hidden from everybody else. Today that is the administrator's entries on the
  Maintainer tab's Account screen; `TabMaintainerModesTests` lists them by name, so a new entry has to
  be decided one way or the other.
- **Every tooltip opens at its control's EDGE, never under the pointer** (`Handlers/Theme/ToolTipPlacement`,
  registered in `App`'s static constructor). Avalonia's default could land a tooltip over the pointer,
  and CRT's tooltips are popup windows, so the next click went to the tooltip (Approve needed three
  clicks). **Do not set `ToolTip.Placement="Pointer"` or a positive `VerticalOffset` on a control.**
  `ToolTipPlacementTests` reproduces the lost click.
- **Every wait the user watches runs under `BusyOverlay`** (owner decision, 2026-09-28). One per
  window, the last child of its root Grid (`Main`, and each dialog that waits); the Maintainer tab
  has none of its own. Find it with `BusyOverlay.For`/`RunAsync(this, ...)`. It blocks clicks AND
  keys, dims after 300 ms, and gives up after `WaitLimit.Maximum` (2 minutes) **with nothing
  happening** - `WaitContext.Report` restarts the clock. **A timeout is never "it failed"**: a call
  that changed something on the server looks again and says what it found (`WaitWording.AfterTimeout`;
  `ServerWait` turns it into `ReviewApiFailure.TimedOut`); local work uses `RunLocalAsync`. Several
  steps are one wait through `HoldAsync`. `BusyOverlayHostsTests` guards the hosts. **Not under it**:
  background work nobody pressed for, board switching, the Submit dialog's upload (own progress), the
  MiniPro test, the scope's live output, and the synchronous table LOAD. The table's SAVE is under it
  (`BoardTableEditor.SaveAsync`: rows taken on the UI thread, file work on the pool).
- **User-visible punctuation is plain ASCII** (hyphens, straight quotes) unless an app label already
  has a character such as "·".
- Biggest files, for context budgeting: `Tabs/Contribute/ComponentContribution.axaml.cs`,
  `Tabs/Schematics/ComponentInfoWindow.axaml.cs`, `Handlers/Data/UserSettings.cs`,
  `Handlers/Data/WorklogManager.cs`, `Tabs/Schematics/TabSchematics.Worklog.cs`,
  `Handlers/Data/DataManager.cs` (~1,700-2,200 lines each). Read the part you need.

## Architecture

### Bootstrap (`Main/`)

- `Main/Program.cs` - entry point; initialises Velopack, then Avalonia.
- `Main/App.axaml.cs` - holds `AppConfig`, nearly every tunable value (file/folder names, URLs,
  timeouts, zoom limits, version helpers, the Wiki page names). **Check here before hardcoding a
  constant.** `App` wires theme application, global exception logging and the startup sequence:
  `Splash` -> `DataManager.InitializeAsync` -> `Main` -> fire-and-forget check-in.
- `Main/Main.axaml.cs` - the main window, the central controller for board selection, schematics,
  and cross-tab state. **Split into partials by area** (`Main.BoardSelection.cs`, `Main.Worklog.cs`,
  `Main.Maintainer.cs`, `Main.SubmissionChecks.cs`, `Main.ModeHint.cs`, ...); its `.axaml.cs` header is
  the file map. It reaches into `TabSchematicsControl` members, so changes there ripple here.

### Tabs (`Tabs/`)

One folder per tab - `About`, `Configuration`, `Contribute`, `Drafts`, `Feedback`, `Maintainer`,
`Oscilloscope`, `Overview`, `Resources`, `Schematics`, `Workbooks` - each a `UserControl` with its
logic in code-behind. `Worklog/` holds the worklog's dialog windows. **Conditional tabs:**
Oscilloscope (`ApplyOscilloscopeTabVisibility`), Workbooks with the worklog bar
(`ApplyWorklogBarVisibility`), Drafts (shown once a draft exists), Maintainer (`EnableMaintainerTab`,
off by default). Hiding a selected tab moves the selection to the first visible one.

#### Drafts (`Tabs/Drafts/`) - local-first contributing

A contributor edits a board locally (the Contribute tab's component editor, the label editor, its
"Edit board as draft" for the whole board in table form, or "Create board" for a new board), sees how
the draft differs from the published copy, and submits.

- **"Edit board as draft"** (Contribute tab, between "Add new component" and "Add a new board"; owner
  request 2026-10-09; `Main.EditBoardAsDraft.cs`): no draft yet - seeded from the published FILE
  (`DraftSeeder.SeedFromPublishedFile`) and opened in the Drafts tab's table; a draft already - nothing
  written, `DraftExistsWindow` says a board has ONE draft and offers it.
The folder also holds that flow's dialogs (`NewBoardWindow`, `NewBoardMaintainerWindow`,
`SubmitDraftWindow`, `DraftDriftWindow`, `MySubmissionsWindow`, `BoardFilesWindow`,
`DiscardDraftWindow`, `ForgetSubmissionWindow`).

- **The badge** (`Main.SubmissionChecks.cs`): `SubmissionReceiptStore.UnreadCommentCount`, the same
  number as "My submissions". Sent submissions are asked about at launch and every minute while the
  window is not minimised and something is open; each check is followed by the launch path (retire
  drafts, then `ApplyDraftsTabVisibility`) on `onFinished`, never `onChanged`.
- **What was just sent cannot be sent again**: the receipt keeps CRT.Data's `DraftFingerprint` (the
  workbook's VALUES, the sidecar's bytes, every draft file by path and content). While the draft
  matches its LATEST submission, Submit is greyed out with the reason - unless that submission never
  finished sending. An empty fingerprint never blocks. `TabDrafts.SubmitAsync` checks again after
  settling the table's edits.
- **A submission the server no longer knows reads "No longer on the server"** (404 after a board
  deletion or the data reset; `SubmissionReceipt.NotFoundUtc`): neutral colour, never blocks Submit,
  no BETA try, not reported on a discard.
- **"Save to draft" in the component editor lands on the Drafts tab**, only after a successful save.
- **A NEW DRAFT IS SEEDED FROM THE PUBLISHED FILE AS IT IS ON DISK** (`DraftSeeder.SeedFromPublishedFile`),
  never from `BoardDataReader`'s cache - an Excel edit made while CRT ran would be lost.
- **No other editor writes a draft under a table with unsaved edits**: the Contribute window and the
  label editor ask `TabDrafts.HasUnsavedTableEditsFor(excelDataFile)` and, if so, save NOTHING and show
  the `SavingElsewhere` notice (Cancel alone; the owner rejected "save the table first" from there).
- **Each draft's row counts problems too** (`DraftStatusReader.CountProblemsCached`), and the list is
  read again when the tab is shown or CRT's window is activated - when an Excel edit arrives.
- **Retiring a published draft** (`PublishedDraftRetirer`, CRT.Data's `DraftRetirement`): a draft whose
  work reached the STABLE source is deleted (BETA alone is not enough - a push-back can undo it), after
  every launch status check and when "My submissions" closes; never one with unsaved table edits.
  **A receipt names its board by ID; the drafts flow is keyed by WORKBOOK path** -
  `DraftStatusReader.ResolveForBoard` maps one to the other. A new board's draft is retired once its
  published board is on the machine. A user on the BETA source is told once, when their work reaches
  stable, to switch back (`Main.SourceSwitchNotice.cs`).

#### The board table editor (`Controls/BoardTable/`, shown by the Drafts and Maintainer tabs)

"Edit in table format" opens a draft's workbook as editable sheets below its row
(`TabDrafts.Table.cs` owns table mode; `BoardTableEditor` the control; `UnsavedTableEditsWindow` the
prompt; colours in `BoardTableColors.axaml`). Colours are the owner's: green added, orange the changed
CELL (published value in its tooltip), red + strike-through deleted, shown where the row used to be.

- **The control only paints.** Every rule is in CRT.Data (`BoardTableDocument`/`BoardTableSheet`,
  `DraftTableSession`, `BoardTableClipboard`, `BoardTableHistory`, `BoardTableRowDrag`,
  `BoardTableSearch`, `BoardTableRowFilter`), because the Maintainer tab shows the same table.
- **Pairing is `BoardDataDiffer`'s rule** (`BoardDraftNaturalKeys`; keys case-insensitive, values
  trimmed + ordinal, first row per key wins), so the tabs' counts equal the draft row's "N rows
  changed". **A key edit with nothing else changed is ONE row modified** (`BoardDataDiffer.PairRenamedRows`,
  owner decision 2026-10-04 - used by the table, the differ and the server's `ReviewSummary`); a key
  AND another cell changed is an add plus a deleted ghost. **A component's identity is its label
  PLUS its region**; **an important signal's is its display name PLUS its KiCad net**.
- **A save replaces all nine sheets, so it may only land on the file it was read from**
  (`DraftWorkbookStore.EditIfUnchanged`, a SHA-256 taken before the read). Never route it through plain
  `Edit`.
- **While on screen it watches its draft file** (`BoardTableEditor.FileWatch.cs`, every 2 s, judged
  against the session's CURRENT fingerprint): changed and nothing unsaved - reload (raising
  `ReloadedFromOutside`, answered like `Saved`); changed with unsaved edits - the `ChangedOnDiskBar`
  and Save off. Never reloads under a cell being typed in or a row dragged. `OpenElsewhereBar` shows
  while Excel's or LibreOffice's lock file exists; Save is off only while the workbook is REALLY held
  open (`DraftTableSession.IsHeldOpen`), since a crashed Excel leaves its lock file behind.
- **The unsaved-edits prompt has wordings** (`UnsavedTableEditsPrompt`): Leaving (Save/Discard/Cancel),
  Reloading, DraftChangedOnDisk and DraftOpenElsewhere (no Save - it would be refused), SavingElsewhere,
  LeavingBoard (the Maintainer tab's Boards screen).
- **The grid is ProDataGrid** (MIT fork of Avalonia's DataGrid; theme included as
  `avares://Avalonia.Controls.DataGrid/Themes/Fluent.v2.xaml`). Traps, each found the hard way:
  - Cell colours come from a per-column `CellTheme` whose `Background` BINDS to the cell's state -
    never set a cell's colour in code (containers are recycled).
  - Data columns set `IsReadOnly = false` explicitly, or the grid makes them read-only.
  - `EditTriggers` is set explicitly (double-click, typing, F2) - the default edits on a single click.
  - `CanUserDeleteRows="False"` - the Delete key removed rows behind the model's back.
  - `FilteringModel.OwnsViewFilter = false` - otherwise the grid wipes the view's filter on every
    re-attach. The user's PICK (`FilterWanted`) is kept apart from the filter applied.
  - A value the TEMPLATE sets outranks a plain style: style it through a class trigger
    (`HeaderStyled`, `RowsDraggable`, `SeveralSelected`, `CellsWrap`).
  - Style selectors match exact types: cell text is `BoardTableCellText` (a TextBlock subclass), so
    use `:is(TextBlock)`.
  - A heading's template keeps a 32px column for sort/filter icons the table never shows - the name's
    presenter spans it (`Grid.ColumnSpan`, through the `HeaderStyled` trigger). Headings wrap, the
    heading row sizes itself (no `ColumnHeaderHeight`), and a fit to a heading measures the NAME
    (`ColumnAutoFitGeometry.HeadingWidth`: the space before it repeated after it).
  - The table must never sit inside a ScrollViewer (it would realise every row) - `DraftsBodyGrid`
    holds list and table as siblings; `ApplyTableMode` changes row `Height`s in place (assigning a new
    `RowDefinitions` drew the open row at zero height).
  - Every key the table's markup names must resolve in CRT's App.axaml (`SharedTableColourKeysTests`);
    a missing `DynamicResource` draws nothing, silently.
- **The table owns the keyboard while open**: the component filter backs off
  (`Main.ShouldReturnFocusToComponentSearch`), and `TabDrafts.FocusTableIfOpen` puts focus in the grid.
  Tab/Shift+Tab move along and wrap (tunnel route); Enter moves down. Buttons hand focus back to the
  grid (`ThenFocusGrid`) - a disabled focused button drops focus to nothing.
- **Undo/redo (Ctrl+Z / Ctrl+Y) live in the model** (`BoardTableHistory`): each step is a snapshot of
  one sheet's live rows taken just before the change, back to the last save. **A new kind of change is
  undoable only if it calls `History.Record` before it mutates.** Inside a cell being edited, Ctrl+Z
  is the text box's. Undoing back to the saved state clears "unsaved".
- **Rows move by the editor's own drag** (`BoardTableEditor.RowDrag.cs`; ProDataGrid's row drag is
  off): the row moves LIVE with a dashed placeholder, the target read off a FROZEN layout
  (`CaptureRowSlots`) - hit-testing the live grid made the row run away near the bottom edge. A red
  ghost is never a target; a whole drag is one undo step. Alt+Up/Down remain; no move buttons.
- **The colour key IS the filter**: five pills - Added, Modified, Deleted, Errors, Warnings - each a
  toggle; picked ones show rows of any picked kind, and **count the WHOLE DRAFT** (owner decision).
  A pick also hides the tabs of sheets it shows nothing of (`SheetsShown`/`SheetToShow`). Moving rows
  is off while anything is picked. Three looks: counting nothing (outlined, 0.4 opacity, not
  pickable), counting rows (filled), picked (2px outline). A sheet tab says its name and change count
  only.
- **The checks are in the table** (`BoardDataChecks`, one rule set with the server): **an error in the
  table is a refusal at submit** (`Every_error_the_table_finds_is_one_the_server_refuses`); a new local
  rule is a warning. A problem cell gets a corner triangle drawn as ONE background brush; a problem in
  no row is a line above the table. Duplicates and rows a save drops are WARNINGS (owner decision
  2026-10-03), marked on every row of the set. Submit on a draft with errors sends nothing and opens
  its table on them.
- **The search box** (`BoardTableSearch`, the Workbooks tab's `WorklogSearchQuery` grammar, every
  cell): decided ONCE when typed; a search finding nothing anywhere keeps only the current sheet with
  `NothingFoundLine`. **The marks are the table's own, NOT ProDataGrid's**: every data column is a
  `BoardTableTextColumn`, whose cells show `BoardTableCellText` - the row's text bound to `CellText`,
  the search handed down once from the grid (`Marks`, inherited). ProDataGrid's search text block built
  its runs from the previous row's result in a recycled cell, so a narrowed search showed rows with
  ANOTHER row's text (2026-10-09). Do not go back to the grid's search model.
- **Several rows deleted in one go** (`BoardTableEditor.Selection.cs`, `SelectionMode="Extended"`):
  "Delete row" calls `BoardTableSheet.DeleteRows` - one undo step, and a component's rows on other
  sheets go with it (`BoardTableDocument.DeleteRowsOfComponent`; highlights at save). A RENAME keeps
  them. `SelectCell` makes its cell the whole selection. ONE "Insert row" button.
- **Row order is what the user sees, and new components are PLACED** (`ComponentPlacement`: a new
  component goes into its category in natural label order, a new category last, an edited one stays
  put; a regional twin straight after its twin). **Order is NOT a change `BoardDataDiffer` counts** -
  making it one is an owner decision not yet taken.
- **Hover cards** (`BoardTableEditor.FilePreview.cs`): pointing at a FILE cell opens a card at once -
  a `Popup` in the window's overlay layer, light dismiss OFF, beside the cell. Every text tooltip is
  instant (`CellToolTipDelay`) and NOT hit-testable, so the pointer reaches the cell below. The bytes
  come from the host's `IBoardTableFileSource`.
- **Every cell wraps; a double-click on a heading's edge fits the column to every row**
  (`BoardTableEditor.TextWrap.cs`, `ColumnAutoFitGeometry`; the grid's own fit measured only built rows).
- **The current cell has no fill and a 2px dashed red frame** (the template's `CurrencyVisual`
  restyled; the grid's focus visual, selection outline and fill handle hidden).
- No "Restore row"/"Revert cell" buttons (owner request, once undo existed); the model keeps them for
  the Maintainer tab.

#### Workbooks (`Tabs/Workbooks/`) and the worklog

`TabWorkbooks.axaml(.cs)`, `.BoardPreviews.cs` (the board pane), `.Summary.cs` (the summary strip),
`.Export.cs` (PDF/ZIP), `.PreviewReorder.cs`. A workbook is a repair on one board; it holds worklogs.

- **`WorklogManager.ResolveActiveWorkbook(workbooks, savedActiveId)` is the ONE place "which workbook"
  is decided** - saved id if still on this board, else the newest. Activation goes through
  `Main.ActivateWorkbook` (`UserSettings.ActiveWorkbookIdByBoard`, persisted per board) via a settable
  action, so tests drive the real path. `SelectWorkbook` has deliberately no "already selected" guard.
- **Every worklog change funnels through `Main.RefreshWorklogBar`**, which refreshes the bar, this tab
  and the Schematics overlay (`SetShowWorklogEntriesList` when the shown workbook changed,
  `RefreshWorklogEntriesListForCurrentWorkbook` otherwise - the second branch was missing once, and
  edits did not redraw). It runs AFTER `_currentBoardData` is assigned on a board change.
- **"Show marked area" ON anchors a pill to its area; OFF parks it in the image's top-right corner**,
  identically on the Schematics view, its thumbnails and this tab's board pane - all three use
  `ParkedBadgeGeometry.ArrangeInTopRightBlock`. A fourth surface must too. Ticking it on an entry with
  no area gives a real, draggable square in the BOTTOM-right (`WorklogDefaultAreaGeometry`); an
  existing area is never moved.
- **Pills and cards open the SAME editor** (`OpenEntryEditor`), with the component checklist from the
  shared `WorklogEntryScope.BuildComponentsInScope` and the Schematics tab's own highlight-rect cache.
- **Status pills and category chips have exactly TWO looks**: SELECTABLE (only in the editor; the
  chosen one filled) and INFORMATIONAL (everywhere else) - every informational one comes from
  `Handlers/Theme/WorklogInfoPillBuilder.cs` (1px border in its own colour). Do not draw one by hand.
  A counted pill drops its icon. A "#N" badge's fill is the CATEGORY colour and its padlock the STATE
  colour.
- **Schematic bitmaps are shared per attachment** and disposed in `OnDetachedFromVisualTree` - which a
  `TabControl` calls on every tab SWITCH. Detach clears the pane FIRST, then disposes; attach rebuilds.
  Anything new rendering one of these bitmaps must be cleared alongside. Disposing under an open editor
  crashes the render thread.
- **Storage: one folder per workbook (`workbook_{id}/`), one per worklog inside it (`worklog_{id}/`),
  each holding its own `index.json`.** **A deleted id is never handed out again** - workbook ids (the
  root's own `index.json`) and worklog ids (`WorkbookRecord.LastWorklogId`) - because ids name
  exported files and attachment folders. The
  allocators also skip ids already on disk. **No migration** of old data.
- **A NEW entry is held in memory until Save** (`thisIsDraftEntry`); Save re-allocates the id
  (`AddEntryRecord`) and moves the draft's attachment folder; Cancel and the title-bar close delete a
  draft's attachments. An existing entry's sub-list changes save at once (`PersistEntrySilently`).
- **The search** (`WorklogSearchQuery` in CRT.Data, `WorklogSearchIndex`): words ANDed, quotes, -minus,
  case-insensitive substring; text fields only (no numbers, no Open/Closed). The matched ids are
  computed once per refresh and shared by list, pane and entry list. Typing is debounced 200 ms;
  `RefreshWorkbooks` re-reads the box itself. Which workbook the TAB shows follows the filter; which is
  ACTIVE does not.
- **Both splitters persist their width**, wired with `AddHandler(..., handledEventsToo: true)` -
  `GridSplitter` marks its release handled.
- **Export** (`WorkbookExportModel` decides what goes in, tested; `WorkbookPdfExporter` paints, layout
  untested): two buttons, the extension from the button (`EnsureFileExtension` REPLACES the other
  format's); files named `Workbook_{id}_{Hardware}_{Board}_{YYYYMMDD}`, never the title; attachments
  under `worklog_{id}` (`WorklogManager.BuildEntryAttachmentsFolderName`, the one definition).
  QuestPDF traps: **`ZipArchive` in Create mode cannot `GetEntry`** (it crashed CRT - `WriteZip` tracks
  names itself); **there is no percentage unit** (placement is proportional via
  `ExportOverlayGeometry` and aspect ratios); **a zero-sized container holding text fails the whole
  document** (three guards); **the icon font is read on the UI thread** (`EnsureIconFontLoaded`); a link
  row needs a scheme or PDF readers ignore it (`BuildLinkTarget`); image sizes come from the PNG/JPEG
  header (`TryReadImageSize` - which JPEG markers carry a size is the subtlety); areas are CLIPPED to the
  image. The export is not opened afterwards (`ExternalTargetLauncher` would refuse its path).
- **Links in user-typed text are clickable** (`TextLinkFinder` in CRT.Data decides - only `http://`,
  `https://`, `www.`; `Handlers/Theme/TextLinkRenderer.cs` renders), merged with search highlighting in
  one pass. Templated rows use the `TextLinkRenderer.LinkText` attached property INSTEAD of `Text`.
  Titles are not linkified.
- **An oscilloscope capture can be filed into a worklog** ("Attach image to worklog" on the popup's
  saved-image banner, hidden without the worklog feature or a workbook): ONE modal
  (`WorklogAttachCaptureWindow`), entries ranked by `WorklogAttachTargets` (component match first,
  then ascending id), the bands named by disabled `ComboBoxItem` headers (preselection by VALUE).
  "Create new worklog" opens the full editor with the photo attached
  (`AttachCapturedPhoto` after `InitializeForNewEntry`), filed against the CURRENT schematic, parked.
  Both attach paths share `WorklogAttachmentWriter` (id allocation skipping existing bytes, rollback on
  a failed persist, the entry re-read from disk).

#### Schematics (`Tabs/Schematics/`)

The most complex tab: schematic/PCB images with three overlay layers (component highlights, the
interactive KiCad overlay, user-drawn traces), the label editor, and the MiniPro IC-test panel.
"Add worklog" starts area-marking mode; the drag opens the full editor on the drawn area (there is no
quick card). **Find the right partial here before grepping** (the map is repeated in
`TabSchematics.axaml.cs`):

| File | Owns |
| --- | --- |
| `TabSchematics.axaml.cs` | Construction, `Initialize`, fullscreen/splitter layout, the trace colour palette, shared helpers |
| `TabSchematics.Types.cs` | Private data types shared by the parts |
| `TabSchematics.Viewport.cs` | Zoom, pan, the transform matrix |
| `TabSchematics.Input.cs` | Pointer, wheel, gesture and keyboard handlers - these only dispatch |
| `TabSchematics.Thumbnails.cs` / `.ThumbnailsDetach.cs` | Thumbnail list and the detached thumbnails window mode |
| `TabSchematics.Highlights.cs` | Component highlight overlays, blink, hover UI, labels |
| `TabSchematics.LabelEditor*.cs` | Label editor lifecycle, interaction, snap context, test seams |
| `TabSchematics.Worklog.cs` / `.Worklog.TestSeams.cs` | Worklog overlay, area drawing and moving |
| `TabSchematics.KiCad*.cs` | KiCad load, panels, rendering, render cache, geometry, hit-testing, calibration |
| `TabSchematics.Settings.cs` | Board-level and global setting rows |

#### Maintainer (`Tabs/Maintainer/`, logic in `Handlers/Maintainer/`)

**"Maintainer" is ONLY the role** - a person in a board's pool; the person who owns this project is
"the project owner" in every document and comment. The ACTIVITY is "review" (`ReviewEndpoints`,
`/api/review/...`, `ReviewSummary` keep their names). The tab was a separate application until
2026-09-29.

- **Shown only when "Enable Maintainer tab" is ticked** (`UserSettings.EnableMaintainerTab`, default
  off), between Drafts and Configuration. `Main.Maintainer.cs` owns visibility and layout: while
  selected, the sidebar and the worklog bar collapse; `LeftPanelWidth` is never saved as 0. "Hide
  the Maintainer tab while no work is waiting for me" shows it only while its badge counts something
  (`MaintainerModes.TabIsShown`) - never taken away while it is on screen, holds unsaved edits, or
  before both lists have been read.
- **The remembered sign-in is restored at launch, quietly** (`RestoreInBackgroundAsync`, never in the
  constructor) for the tab's badge. The token is DPAPI-protected on Windows and not stored elsewhere
  (`ReviewSessionProtection` - its entropy string must not change). **The API's `refreshToken` IS the
  bearer token**; there is no access-token exchange.
- **A sign-in on the tab is the sign-in for the whole of CRT** (`SignedInChanged` ->
  `Main.ShareMaintainerSignIn`): the Feedback tab and the Submit dialog use the account's address,
  read only, and submissions go with the account's token. `ContactAddress` is the one rule. It stays
  shared whatever the Configuration tab says, until signing out.
- **The tab's badge** (`MaintainerModes.TabAttention`): the two queues' attention counts ADDED UP. No
  tooltips on it. A 401 clears what it counts.
- **The minute check** (`QueueRefreshRules.MinuteCheck`): everything while on screen with CRT in
  front; otherwise only the queue and the BETA list while the badge can be seen. Every request slides
  the 30-day session (accepted with the request). Nothing open is reloaded unless its row changed.
- **Four screens as a tab strip** - Boards, "Queue: Contributor submissions", "Queue: Awaiting push
  from BETA to stable", Account. **Their labels are CRT.Data's `MaintainerScreenWording`** - never write
  a screen's name as a literal. Switching screen HIDES, never closes. With nothing waiting, the tab
  opens on Boards (`MaintainerModes.ScreenOnOpening`); the queues open on the last-looked-at entry
  (`MaintainerModes.EntryToOpen`).
- **Account**: every maintainer has "My account" (name, address, password, Sign out -
  `MyAccountView`; the session alone is enough, no current password - an accepted risk) and "Server
  version". The administrator's own entries (Maintainers, Order of boards, Unused files, Rebuild
  checksum manifests, Delete a board, API usage, Reset contribution data) carry the padlock and are
  taken OUT of the list for anybody else (`ShowAdministratorEntries` - not `IsVisible = false`, which a
  `VirtualizingStackPanel` undoes).
- **The submission view IS the table** (`BoardTableEditor.Open(BoardTableDocument)`, document mode;
  Save raises `SaveRequested` and the host saves it as an AMENDMENT, decided by the server's
  `AmendSubmissionFlow`). A submission has three views (Board data, Files, Contributor); switching hides.
  A NEW board's table is compared with the submission itself, so it starts white. What the table
  cannot show (highlight and calibration changes, check findings) is listed above it (`ReviewNotInTable`).
  **A file replaced under its own path colours nothing - the Files button's COUNT says it**
  (`SubmissionViews.ChangingFiles`). The queue is grouped by board with DISABLED heading items, so it is
  read by ITEM, never by index. Approve is green, the others red; Approve stays last.
- **A board has six views** (Board data, Files, Contributor, Maintainer, History, Statistics). Board
  data and Files have a **Data source: BETA | Stable** switch; stable is read-only. **A Board data edit goes STRAIGHT TO BETA** (save -> `edit/check` -> a reason in
  `PublishBoardChangeWindow` -> `boards/edit`), read-only for a board the account does not maintain
  or one waiting under BETA > Stable. The table and files are read again when BETA moves
  (`BetaContentHash`), never under unsaved edits.
- **A board's email addresses go only to its own maintainers and the administrator** (owner request,
  2026-10-05; `BoardDetailAnswer.AddressesHidden`): anybody else sees names, and the Contributor and
  Maintainer views say so (`BoardsDisplay.AddressesHiddenLine`).
- **A new board is PLACED in the drop-down lists before it can be approved** (`BoardPlacementView`,
  the same drag as `ListRowDrag`). Two boards may never share a hardware name + board name
  (`MasterListing.NamesTakenBy`).
- **File trees** (`FileTreeView`, model in `Handlers/FileTree`): changed files in the table's colours,
  sizes on every row, hover card, double-click to open (BETA/stable files from their PUBLIC addresses -
  the token is never sent there). Never inside a ScrollViewer.
- **The Contributor view is the contributor's record**; no text a user reads calls a contributor
  "their", "them" or "they" (owner request) - the wording classes' tests assert it.
- **Every person the tab names is BOLD, by the name on the account** (owner request, 2026-10-09):
  maintainers and contributors go through `PersonRuns` (drawn by `TabMaintainer.ShowCounts`); the
  server's history (`BoardHistoryRules`) names whoever did a thing and whoever a pool change names by
  account name, the address only without one. A pool change reads "Maintainer added: **Anna**" /
  "by **Dennis**". A line with bold runs has `Text == null` - tests read it through `TabMaintainer.TextOf`.
- **No stage line under a board's name** (owner request, 2026-10-09: the three numbered cards -
  Submitted, BETA, Stable - and their "Now:" sentence were removed as confusing, "and then decide
  later"; `BoardStagesDisplay` is in git history). **Why BETA's table cannot be changed is an amber
  panel of its own** (`ReadOnlyNotice`), and **so is why Approve is off on the queue**
  (`BeforeApprovingNotice`; owner request, 2026-10-09: "The UI should have a uniform and consistent
  look"). Both sit DIRECTLY ABOVE their view switch, said whichever view is open, and both are
  App.axaml's `Border.Notice` with a `TextBlock.NoticeIcon` and `TextBlock.NoticeText` -
  a new "why this cannot be done now" panel uses those classes, never its own colours. The read-only
  reason comes from the board's DETAIL (`BoardDetailAnswer.MayEdit`/`MayNotEditReason`, the table's
  own rule on the server) or BETA's table, whichever was read last - so it is there before Board data
  is read. The line above BETA's table says what happened and what the table is, never why it is
  read-only (`BoardSections.TableNote`; `ReadOnlyReason` is the panel's).
- **"Compare sources"** (owner request, 2026-10-09; `BoardDetailView.Compare.cs`), right after the BETA |
  Stable buttons, Board data only, remembered (`UserSettings.MaintainerCompareSources`, handed in by
  `TabMaintainer.UseRememberedComparison`): BETA's table is BUILT against the stable source's board and
  the stable table against BETA's (`BoardTableDocument.Create(otherSource, thisSource, ...)`), the
  tooltip naming "Stable source value" / "BETA source value"; file cards read the other side from the
  other source's address (`PublishedTableFileSource.Baseline`). Both tables are read when compared.
  **It is off while BETA's table holds an unsaved change** (the editor's `UnsavedChangesChanged`):
  comparing rebuilds the table. A cell still in its editor counts - the box commits it first
  (`BoardTableEditor.CommitCellEdit`), and nothing rebuilds BETA's table under `IsEditingCell`. A save
  still sends BETA's live rows applied to BETA as read - never the other source's red rows. **Every
  table read or opened goes through `ShowTables`**: answers taken first, each table built ONCE against
  the other as it now is, the other rebuilt only when what it is compared with moved
  (`ApplyComparison`) - never build a table any other way. Ticked with the stable table unreadable,
  BETA's line says it is not compared and why (`NotComparedReason`, part of what it is built against).
- **The Boards list says how each board's submissions went** ("19 submissions in total; 2 rejected, 1 in
  BETA, 17 in stable" - `BoardSubmissionCounts`, `BoardOverviewFlow.SubmissionCounts`). Account's
  Maintainers and "Delete a board" list boards in the Boards list's order (`BoardsDisplay.InBoardsListOrder`).
- **Statistics has a graph of views per day** (`BoardViewsChart`, every number `ViewsChartGeometry`;
  owner's choice: drawn, no charting library): 30 / 90 / 365 days, a day with none drawn at 0, the day
  pointed at written above the plot (a drawn readout, not a tooltip - see the tooltip rule). Colours
  `Chart_*` in both themes, validated against CRT's grounds.
- **Quitting CRT asks about unsaved table edits** of both tables (`HasUnsavedTableEdits`), after the
  Drafts tab's, in one continuation. The tab has no `BusyOverlay` of its own. The component filter
  never takes focus over it. The server address is `AppConfig.CrtServerRootUrl` (no "/api");
  `CrtServerBaseUrl` adds "/api" for `SubmissionClient`.

### Data layer

`src/CRT.Data/` (Avalonia-free `net10.0`, namespace `Handlers.DataHandling`) holds what the app, the
server and the Maintainer tab share: the board schema and readers/writers (`BoardData`,
`BoardDataReader`, `BoardDataWriter`, `BoardWorkbookSchema` - both directions, round-tripped by
reflection in its tests - `BoardWorkbookWriter`/`BoardWorkbookStyle`), the table editor's model, the
checks (`BoardDataChecks`), placement, search, the contracts and wordings, and the logging seam
`CrtLog` (null sink by default; CRT installs `AppLoggerAdapter`). **`CRT.Data` must never depend on
Avalonia.** `src/CRT.App/Handlers/Data/` keeps what orchestrates the app: `DataManager`,
`DataValidator` (its rules are `BoardDataChecks`), `UserSettings`, `Logger`, `WorklogManager`, the
KiCad loaders, `PublishedDraftRetirer`. Check which side a class belongs on before moving it - a hidden
dependency makes a "move" a rewrite. `HighlightRectBuilder`, `PolygonGeometry` and `RectGeometry` stay
in CRT.App (Avalonia types, 21 call sites).

- **`BoardWorkbookSchema.AllSheets` is the order of every workbook written and of the table's tabs,
  ending in "Credits".** Readers find sheets by name.
- **How every written workbook looks is read off the shipped references** (C64 250407, C128 310378):
  the title band is BLACK with white text, the header row's height differs per sheet, and every sheet
  has the same preamble - hardware, board, revision date, a blank line, the title band, headers on row
  6, panes frozen under them. Check a change against the references' XML, not memory.
- **The Overview AXAML binds `OverviewRow`/`OverviewLink` through an `xmlns:data` mapping** - update
  `TabOverview.axaml` if you move them.

### Content (`Assets/Data/`)

Hardware reference content is data, not code: `<Manufacturer>/<Hardware>/<Board>/...`, with
per-manufacturer `Shared files` and a top-level `Generic shared files`. A master workbook
(`Classic-Repair-Toolbox.v<version>.xlsx`, sheets `Hardware & Board` and `Oscilloscope`) lists all
hardware/boards and oscilloscope SCPI command sets; each board has its own workbook and a JSON sidecar.
CRT syncs it from `classic-repair-toolbox.dk` by SHA-256 manifest, independent of app releases. **There
are two generations**: each master names its own generation of board files, and `DataManager` picks
the master by version (highest at or below the app's own); the server writes only the newest.

### External integrations (`Handlers/`)

- `Online/OnlineServices` - the manifest and the data sync; its four manifest-validation predicates
  are a trust boundary and are tested. Every request names its CRT (`User-Agent: CRT <version>`).
- `Online/UpdateService` - Velopack against GitHub Releases; `UpdateChannelFilter` decides which
  release stages the ALPHA/BETA checkboxes admit, applied BEFORE Velopack ranks the feed.
- `Oscilloscope/IScopeClient` - the seam the Oscilloscope tab's sequencing takes; `ScopeScpiClient` is
  the untested TCP boundary. `ScopeCommandResolver`/`ScopeValueMapper` turn the master workbook's
  command strings into scope interactions.
- `MiniPro/` - `IMiniproRunner` (`MiniproProcessRunner` spawns the bundled `minipro.exe`, Windows
  only; `MockMiniproRunner` simulates).
- `Security/ExternalTargetLauncher` - **the only sanctioned way to open an external link or local
  file**: HTTP/HTTPS/mailto, or a local path inside the data root with an allowlisted extension.
- **QuestPDF**: used under its Community licence (free under $1M revenue - a condition on any fork).
  `WorkbookPdfExporter.ConfigureQuestPdf()` (called at launch, and by the tests) sets the licence and
  flips 2026.9.0's font defaults back (`UseSystemFonts` on, no throwing on missing glyphs or families,
  no detailed layout errors). **Check QuestPDF's release notes on every bump.** The icon font is
  referenced by its family ("Font Awesome 7 Free") at weight 900.

## The contribution service (`src/CRT.Server/`)

An ASP.NET Core minimal-API service on the project owner's AlmaLinux box, reached through Apache at
`/api/`. **Installing and running it: [src/CRT.Server/INSTALLING.md](../src/CRT.Server/INSTALLING.md)**
(the systemd unit ships as `crt-server.service`). The agent builds; the project owner deploys by hand.

**The shape to keep: endpoints are a RIM, flows hold the decisions.** `*Endpoints.cs` reads a request
into plain values, calls one `*Flows` method, and maps the verdict onto a status code - nothing else.
Every rule lives in a pure class taking its stores and mailer as arguments, tested against fakes.

Load-bearing, and easy to undo by accident:

- **The status codes are the security design.** Registration and password reset always answer 202;
  login answers 401 with no reason. A 409 for a taken address is an enumeration oracle.
- **The rate limiter runs BEFORE the Argon2 hasher** (~128 MiB per verification).
- **A request body's size limit is written on its ROUTE** (`.WithBodyLimit(...)`); everything else
  gets 64 KB. `RequestBodyLimitsTests` builds the real route table and fails on a body-reading route
  missing from its list.
- **Only token HASHES are stored** - sessions, verification and reset links, submission tokens.
- **Every mail is HTML with a plain-text copy, written ONCE as `MailBody` blocks** that HTML-encode
  every run - no template writes markup. Each greets by the account's name and ends with the contact
  line. The BETA check box is named through `ConfigurationWording`.
- **`SubmissionPathRules` is the single path-containment rule** (resolve, then check containment).
  Every write path and the maintainer's file read go through it.
- **Two operations are irreversible**: a publish (`ApprovePublishFlow` - no revision history is
  kept; the order of its checks is the design) and deleting a board (`BoardDeletionFlow` - held to a
  fingerprint; refused for a board an older master lists or with a file another board uses; the
  record is deleted LAST so pressing Delete again finishes).
- **Two roles, authority PER BOARD** (owner decision 2026-09-25): an ADMINISTRATOR
  (`accounts.is_administrator`, granted only by hand in SQL) reviews and publishes everything and
  assigns maintainers; a MAINTAINER is an account in a board's pool. The question is always
  "administrator, or in THIS board's pool" - `ReviewAuthority`, given a `ReviewAccess` loaded per
  request, so removal and locking bite on the next request. **An administrator may also be in a pool**
  (owner request 2026-10-05) - to be named as a board's maintainer; the row never counts as the
  maintainer half of a two-person approval (`CanGiveMaintainerApproval`) and gets that board's mail
  once (`SubmissionRouting`, deduplicated by address).
- **A submission REPLACING a shared file needs two approvals** - a maintainer of the board AND the
  administrator, for BETA and again for stable; adding a new shared file needs one. `ApprovalRules`
  (CRT.Data) is the rule and its `ApprovalStatus` travels to the Maintainer tab, so the button cannot
  promise a publish the server would not do. It is decided from the tree as it is NOW, not the stored flag.
- **Publishing is TWO stages**: Approve writes BETA; "Publish to stable" copies a BOARD from BETA
  (`ProductionPromotionPlan`/`ProductionPromoter`/`ProductionPromotionFlow`), only bytes already in
  BETA, refused if BETA moved since the maintainer looked. Both stages take the one `PublishLock`. Off
  until the three `Production*` settings are set. **For now ONLY ADMINISTRATORS publish to stable**
  (owner request 2026-10-05; `ProductionPublishingAdministratorsOnly`, true unless set, asked through
  `ReviewAuthority.CanPublishToProduction`): a maintainer's publish is refused with CRT.Data's
  `StablePublishing` sentence, their plan carries it (so CRT greys the button out), their rows never
  `AwaitsYou`, and a shared-file replacement needs the administrator alone. Push back and Reject stay
  theirs. **Do not remove the setting or flip its default without being asked.**
- **One submission in BETA per board** (owner decision): an approval is refused while the board
  waits for stable (`OneSubmissionInBeta`), on only when stable publishing is configured. A board
  copied to stable BY HAND is recognised and recorded (`RecordIfProductionAlreadyHoldsAsync`).
- **"Push back to queue" rolls BETA back to stable's state** (`BetaRollbackFlow`), per board - it
  returns EVERY submission merged since the last promotion to the queue, so its confirmation names
  them all; "Reject" on that screen is the same rollback. It works because merged submissions keep
  their blobs. "Never promoted" is the board's RECORD, never an empty folder. Recorded as `returned`
  (`submission_beta_returns`). Paths are ORDINAL.
- **A maintainer's edit of a BOARD goes straight to BETA through the ordinary approval**
  (`BoardEditFlow`: a submission from the account, then `ApprovePublishFlow.ApproveAsync`). **Do not
  give it a path that writes BETA by itself** - the approval keeps every rule.
- **Automatic removal stays inside the board's own folder** (`AutomaticRemovalScope`); a removal is
  always from a list the approver was shown (`FileRemovalPreview`, 409 if it differs). `DataTreeUsage`
  is the one rule for what a tree uses - a change may only make it keep MORE; a folder CRT reads by
  name must be in `FoldersReadByName`; a board workbook is read at its spelling ON DISK.
- **A served file is REPLACED, never opened for writing, and every folder is checked FIRST**
  (`FileReplacer`, `VerifiedFileCopy`, `TreeWriteAccess`) - the project owner copies data in as root.
  The highlight file is written only when its CONTENT changes. Every blob is re-verified before a
  publish writes anything.
- **The revision date and the "# Hardware:"/"# Board:" caption are the SERVER's** at both stages
  (`PublishMerge`); `BoardDataCaptionTests` fails on a `new BoardData { ... }` naming neither. Do not
  make the client stamp a date.
- **WHICH files a submission may carry is `SubmissionFileRules`**, run at create AND in `PublishPlan`;
  `SubmissionRulesShippedDataTests` runs every rule over all of `Assets/Data` - do not loosen a rule
  to pass it. KiCad data travels by `SubmissionKiCadFiles`. "Already held" includes the published tree
  at the same path (imported at create).
- **A newer submission from the same contributor replaces the older one** (`withdrawn`, only if still
  pending and never amended).
- **The Boards screen is for every maintainer, every board** (`CanReviewAnything`) - do not narrow
  WHICH boards without asking. **The ADDRESSES in it are narrowed** (`CanSeeAddressesOf`): a new field
  carrying a person must follow `AddressesHidden`; the flow's test serialises the whole answer and fails
  on any "@".
- **A board's history** comes from its submissions and the audit rows naming it - a new kind of
  event needs its audit subject to be the board id and a place in `BoardHistoryRules.ShownActions`.
  What a submission changed in BETA is recorded AT the publish (`submission_changes`).
- **Usage data**: board views (`crt_board_views`, one row per view, names from the published master,
  no address or identifier; local-network views counted while `CountLocalNetworkBoardViews`), API
  usage (`crt_api_calls`, per day/route pattern/version, counted before the version gate), and the
  launch check-in (`crt_update`, a table no migration creates, written as CRT 2.x always had it -
  including the sender's address). Feedback is saved under `FeedbackRoot` (group-writable, capped by
  `FeedbackMaxStoredBytes`) and mailed from the service with Reply-To the sender.
- **The data reset** (Account > "Reset contribution data", `DataResetFlow`): one transaction, DELETE
  never TRUNCATE, never the data trees, `crt_update` or feedback; only while `AllowDataReset` is true.
- **Database migrations are append-only code** (`Migrations/*.sql`, copied with the publish): the
  runner refuses to start on a changed applied file (even a comment), a gap, an out-of-order number or
  a missing applied file. **Never edit an applied migration** - not even to reword a comment.

## Decisions on record, and accepted risks

Settled by the project owner; do not reverse or "fix" without asking.

- **No two-factor sign-in** (TOTP deferred 2026-09-25). A stolen maintainer account can publish to
  that maintainer's boards (BETA, and stable if they may) with a password as the only factor.
  Mitigations: per-board authority, only bytes already in BETA reach stable, a shared-file
  replacement needs the administrator too, a mail to the administrator on a maintainer's stable
  publish, the audit trail.
- **A stolen session token can change the account's address and password** (2026-10-03: no current
  password asked, "as I see it as you are already logged in"). Mitigations: the old address is mailed,
  other sessions are signed out, the token is DPAPI-protected on Windows.
- **No publish history is kept.** A bad publish is recovered by a corrective publish or the project
  owner's backups. Do not build a revision store without asking.
- **Images are signature-checked, not decoded** on the server (decoding needs an imaging library).
- **Blobs of merged submissions are kept for ever**, so the next edit uploads only what changed.
- **No administrator feed or anomaly alerts** - the audit rows exist; nothing renders them yet.
- **EPPlus licensing**: "Use EPPlus for now, and I WILL take this later" - do not swap the library.
- **Open, not decided**: a fact affecting several boards is several submissions (no cross-board
  submission); two contributors editing the same component - the later approval wins, with nothing
  more until it demonstrably hurts.
- **Backups**: the design assumes both data trees and the database are backed up and restorable.

`Assets/NewContributeStrategy.md` (the phase plan this pipeline was built from) was deleted on
2026-10-05 as finished. About 110 code comments still cite it ("Phase 4, task 3", "session 2c") -
those refer to its last version in git history (`git log --all -- Assets/NewContributeStrategy.md`).

## Installed CRTs keep working

**Every CRT ever released stays installed, and must keep working** (owner request, 2026-10-04). The
server is always the newest party, so the SERVER bends: there is one API, no `/api/v1` beside
`/api/v2`, and it only ever grows. Where a CRT truly cannot be served any more, it is TOLD so in words.

**It protects what a RELEASED CRT uses - never a pre-release** (owner decision, 2026-10-04). Anything
only `-alpha.N`/`-beta.N` builds use is changed IN PLACE, with CRT's side changed in the same session.
**So the first question about any API change is "does a RELEASED CRT use this?"** - is it in a
`crt-*.txt`, or one of the addresses CRT 2.x posts to (the check-in, feedback, the old contribution
upload - held by their own contract tests)? No: change it freely. Yes:

- **Add, never rename or remove.** New optional fields, new answer fields, new routes; a missing new
  field means the old behaviour.
- **A new meaning is a new route.**
- **A new REQUIRED request field, a narrower rule, a new member of an enum an answer carries** - each
  is a break.
- **The server ships before the CRT release that needs it.**
- **Readers ignore what they do not know** (`ReviewApiContract.WireSettings` skips unknown fields;
  never `UnmappedMemberHandling.Disallow`).
- **The data trees are an API too**: add sheets and columns, never rename or drop what an older
  reader looks for.

**The machinery:**

- Every request names its CRT (`User-Agent: CRT <version>`); CRT.Data's `CrtVersion` orders versions.
- `ClientVersionPolicy` can answer a CRT older than a minimum on `/api/submissions` or the Maintainer
  tab's routes with **HTTP 426 and `ClientOutdatedAnswer`**. **No minimum is set.** The check-in,
  feedback, board views, health and the old contribution address are **never** refused
  (`ClientVersionPolicyTests` fails on an unclassified new top-level route).
- **The API REVISION tells a CRT it is too old**: `ClientVersionContract.ApiRevision`, sent as
  `X-CRT-Api-Revision`; a LOWER revision on the gated routes gets the same 426. **`ApiCompatibilityTests`
  raises it for you**: `api-revision-<N>.txt` holds what a CRT built for revision N may use (written
  the first time, grown as the API grows), and a change that breaks it fails until `ApiRevision` goes
  up. Raise it by hand for what the surface cannot see (a new meaning, a newly required field). Each
  raise is a MAJOR server bump. `GET /api/health` and Account > "Server version" show both revisions.
- **CRT shows the server's words** (`ApiRefusal.Read`; `ReviewApiClient.WithServersWords`).
- **An outdated CRT gets "CRT has to be updated" over the Drafts and Maintainer tabs, which cannot
  be closed** (owner request, 2026-10-09; `Main.UpdateRequired.cs`, `AppUpdateRequirement`,
  `UpdateRequiredOverlay` in `Controls/`). Once either tab is in use - the Drafts tab shown or the
  Maintainer tab on (`AppUpdateRequirement.AsksServer`), at launch or later - CRT asks `/api/health`'s
  `ApiRevision` once: a higher one covers BOTH tabs. Nobody using neither is asked - but "Edit board
  as draft" and "Add a new board" ask it FIRST, awaited, when nothing has yet
  (`AskApiRevisionBeforeFirstDraftAsync`), or a machine's first draft opens under the cover that
  follows. **Whether a tab is covered has ONE record each**: the Maintainer tab reads its overlay
  (`IsUpdateRequiredShown`), and the minute submission check stops on `AppUpdateRequirement`'s Drafts
  reason - never a flag of their own beside it. Any 426 either
  client meets raises `ApiOutdatedSignal` (in `SubmissionClient.SignalIfOutdated` - every refusal it
  reads, the discard notice's too - and `ReviewApiClient.StatusFailureAsync`) and covers only THAT
  area's tab - a minimum version is per area - then the revision is asked again. Never lifted while
  CRT runs; the overlay's one button installs the update found or opens the releases page. The
  signal is a static event only `Main.StartAsync` subscribes to, so a test driving a client into a
  426 leaks nothing. It is not a second `BusyOverlay`: it covers one tab on purpose, though both
  fade and take keys through `OverlayCover`. **Installing an update asks about unsaved table edits
  first** (`Main.ConfirmLeavingTablesAsync`, shared with quitting): Velopack's restart exits without a
  Closing event. A covered Maintainer tab asks without a Save (`LeavingUpdateRequired`); a covered
  Drafts tab gets nothing new - "Edit board as draft" and "Add a new board" make nothing - and another
  editor held back by its table is told what settles it (`TabDrafts.SavingBlockedFor`).
- **When CRT.App's `InformationalVersion` is a release (no `-`), the surface must be frozen as
  `crt-<version>.txt`**: the first test run writes it and fails once - commit it with the release.
  **Never edit a `crt-*.txt`**; put the field or route back instead. A break the owner decided on goes
  in `allowed-breaks.txt` with its reason. A `crt-<version>.txt` for a version not yet PUBLISHED may be
  rewritten.
- **Every answer with data in it is a named CRT.Data record** (`*Answer`/`*Response`, or listed in
  `ApiSurface.OtherAnswers`) - an endpoint may write an anonymous object only to carry `message`,
  `error` or `errors`.
- **CRT 2.x's contribution upload is answered, not 404'd**: `POST /api/legacy/contribution` gives the
  426 `OUTDATED_VERSION 3.0.0 - ...` text 2.5.0 and later understand. Apache forwards the old address
  (`/app-contribution/api/`) to it, as it does `/app-checkin/` and `/app-feedback/` (INSTALLING.md).
- **The server counts which CRT version calls which route** (`ApiUsageCounter`, Account > "API usage").

**Retiring a route a released CRT uses** (owner decision, 2026-10-04; nothing before 3.0.0's release):
API usage is the evidence; the owner decides per route, never for the Forever routes; the route stays
MAPPED and answers 426 (`ClientVersionContract.Outdated`), never 404; what it carried goes in
`allowed-breaks.txt` with the reason; a MAJOR server bump names it.

## CRT.Server's version is YOURS to bump

**The project owner bumps CRT by hand. They do NOT bump CRT.Server** - they asked (2026-09-26) for it
to be handled for them, because it only has to be visibly versioned so a deployment can be identified.
When you change the server, you decide the bump and make it. The policy, the judgement calls and the
history are in [src/CRT.Server/VERSION.md](../src/CRT.Server/VERSION.md); read it before bumping.

SemVer 2.0.0, judged on **the HTTP API an installed client depends on**: MAJOR when a client that
worked can now fail (a route or field renamed or removed, a field made required, a rule that starts
refusing what it accepted); MINOR for new behaviour that breaks nobody; PATCH for a defect fix or
anything no caller can observe. **A change only to `CRT.Data` still bumps the server** when the
server's behaviour moves with it. **The numbering RESTARTED at `1.0.0` on 2026-10-05** (owner request,
at go-live) **and again on 2026-10-09** (owner request, with the contribution data reset again -
nothing had been released to anybody else), the API version back at 1 with it (`api-revision-1.txt`
then holds that day's API). The earlier numberings (development up to `5.1.0`, then `1.0.0` to
`2.1.0`) are archived at the bottom of VERSION.md without backticks, so `ServerVersionTests` reads
only the live table. A server version named in an older comment ("server 4.3.0") is one of those
series; every version since the second restart holds all of it. The next breaking change is `2.0.0`.

Two edits per bump: `InformationalVersion` in [CRT.Server.csproj](../src/CRT.Server/CRT.Server.csproj),
and a row in VERSION.md's history saying what a caller would notice. `GET /api/health` reports it,
without the `+<commit>` metadata (the endpoint is public). **Machine-checked:** the `Stop` hook
[hooks/server-version-bump.sh](hooks/server-version-bump.sh) warns when a file the service is built
from changed while the version stood still - "nothing a caller can observe changed" is sometimes the
right answer.

**Say the API version whenever the API changed** (owner request, 2026-10-09: "whenever you change
anything for the API, then do comment what is the API version on SERVER and APPLICATION, so I always
can track this"). End that turn's summary with one line, `API version: server N, application N`, and
say whether it moved. Both are `ClientVersionContract.ApiRevision` - what Account > "Server version"
shows as "API server version" and "API application version". VERSION.md's history carries it per
server version (its **API** column). **Machine-checked:** the `Stop` hook
[hooks/api-version-report.sh](hooks/api-version-report.sh) prints the line when a contract, an endpoint,
CRT's routes or parser, or the recorded API surface changed against the last commit.

## Release process

Versioning lives in [CRT.App.csproj](../src/CRT.App/CRT.App.csproj) (`AssemblyVersion`/`InformationalVersion`)
- bump `InformationalVersion` before releasing; it is the only place a release version is entered.
Releases are made by hand from the GitHub Actions tab - run
[.github/workflows/build-and-release.yml](../.github/workflows/build-and-release.yml) with no inputs -
**never by pushing a tag** (that trigger was removed). Its first job reads `InformationalVersion` and
derives pre-release status (a `-` means pre-release); every other job uses that output. It runs the
test suite and stops if it is red, then a CodeQL scan, then builds, signs and packages (Velopack)
self-contained builds for win-x64, linux-x64, osx-x64 and osx-arm64, and publishes a GitHub Release
with [CHANGELOG.md](../CHANGELOG.md) as the body. The tag is created by that last step, so a failed run
leaves nothing to clean up. The release body is written by hand and is off-limits - see
[Hands off CHANGELOG.md](#hands-off-changelogmd).

**Windows signing is done INSIDE `vpk pack`** (`VPK_SIGN_TEMPLATE` running jsign against the YubiKey),
followed by a step that fails the release unless every `.exe`/`.dll` in the full `.nupkg`, and
`Setup.exe`, carries a valid signature. Do not go back to separate jsign steps around the pack - they
could not reach the `Squirrel.exe` and launcher stub Velopack adds inside the package. The template
must stay sequential (a YubiKey locks after three wrong PINs). macOS and Linux builds are not signed.
There is no separate maintainer application to release any more.
