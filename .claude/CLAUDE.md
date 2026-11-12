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
and a MiniPro USB IC programmer/tester. Users improve the data from inside CRT (the Contribute
tab); maintainers review it in CRT's Maintainer tab, and the contribution service (CRT.Server)
publishes it to the BETA data and then to the stable data.

## One change, every side of it

**This is one system in four parts, and a change to shared behaviour is not done until every
part that shares it has been changed in the SAME session.** The parts are the CRT desktop app
(`src/CRT.App/`, including the Maintainer tab and, in `src/CRT.App/Controls/`, the controls the
Contribute and Maintainer tabs share), the contribution service (`src/CRT.Server/`), the shared library
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
working"). The site's files are edited where they are served, through the Z: drive -
`Z:\mydir\http\classic-repair-toolbox.dk`, never the UNC path (owner request, 2026-10-12; the local
mirror `Assets/Webserver/` was deleted that day). That folder is a git repository of its own with its
own CLAUDE.md - read it before changing anything there. **One page there reads this project's
database directly**: the Fun facts page (`public_html/funfacts/`) queries `crt_board_views`,
`crt_update`, `crt_statistics` and `crt_versions`, so a migration that renames or drops one of their
columns changes those queries in the same session (0020's `systemId` -> `boardId` ended the page with
a fatal error until 2026-10-12).

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

## Code reviews: what each severity means here

**The review format is the global one in `~/.claude/CLAUDE.md`, the same in every project** (owner
request, 2026-10-15): every finding rated with a **Severity** and a **Criticality** (Must fix,
Should fix, Nice to have), the findings as a table `# | Finding | Severity | Criticality | Fix`
worst first, then a tally and a one-line verdict on whether another run is worthwhile; in
`ReportFindings`, `summary` starts with `[Severity: High | Criticality: Must fix]` and
`short_summary` with `[High/Must fix]`. It replaced this project's own `[High]` prefix, `Severity:
0 Critical, ...` tally and stop signal (2026-09-29) - a global hook refuses findings without both
ratings, so the old prefix alone would now be sent back. Every code review is a security review
too, its findings in the same table marked "Security:", and `/security-review` uses the same format.
This section only says what the levels MEAN in this project. `/simplify` is not a review (it
applies its fixes) and carries no ratings.

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
serious and the verdict would never say stop. Rank findings most severe first.

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

**Seven Wiki pages are opened by buttons in the shipped app**, across eight buttons:
`Workbooks-tab`, `MiniPro-programmer` (two buttons), `Controlling-oscilloscope-with-keyboard`,
`Synchronize-oscilloscope`, `Maintainer-tab`, `View-boards-from-online-source` and `Table-view` (the table's
"Learn more", owner request 2026-11-11: "we need to reference a Wiki, so the user can see what options there is -
e.g. how to copy/paste" - a 4.0 page, pasted on release day with the others). They are named
once, in `AppConfig` (`WikiPage*` plus `WikiPageUrl`). Renaming or deleting one breaks a button in
builds already installed, which no update can fix, so `WikiHelpPageNamesTests` asserts each name
still matches a file in `Assets/Wiki/`. Add an entry there for any new help button. **Since CRT 4.0
such a button reads "Learn more"** (App.axaml's `LearnMore` class - a link, underlined; the plan:
"The remaining help buttons read 'Learn more'"); the "?" icon (`HelpIconButton`) is left only where it
shows a tooltip and opens nothing.

**The Wiki is SHORT since CRT 4.0** (the plan: "The Wiki shrinks to installing, the reference pages,
and the six pages installed CRTs open"; 2026-10-29) - the screens explain themselves, so a tab or a
contributing flow gets no page of its own. Fourteen pages left the mirror then (`Schematics-tab`,
`Contribute-tab`, `Drafts-tab`, `Main-Excel`, ... - listed in [Assets/Wiki/README.md](../Assets/Wiki/README.md)),
to be deleted from the live Wiki on release day; do not add them back without being asked. **The
mirror describes 4.0 from that day, so its pages are pasted on release day**, not before - the live
Wiki describes what everybody runs.

**Some old Wiki pages are deliberately not mirrored** (orphans superseded by newer pages) - see
[Assets/Wiki/README.md](../Assets/Wiki/README.md). Do not add them back without being asked.

## Build, test, run, publish

- **Run the tests: `dotnet test Classic-Repair-Toolbox.slnx`** (about 3 minutes in Release, build
  included; needs no hardware, network or display)
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
    shouts about it and the update notice says `(simulated)`. Set for F5 and `watch`.
  - `--demo-notices` - play every notice once after launch, narrated by the mode hint, with example
    words (`Main.NoticeDemo.cs`). Leaves a notice showing for a real reason alone. Not set for F5.
  - `--show-tours[=<version>]` - each tab's introduction tour the first time the tab is shown in the run,
    whatever was seen, remembering nothing; with a version, the tour of what came since it
    (`Main.Tours.cs`). Not set for F5 - its own launch profile, "(introduction tours)", in VS and VS Code.
- **To skip the online data sync while iterating, untick "Check for new or updated data at
  application launch" in the Configuration tab.** It is a normal user setting.
- **macOS data folder**: .NET 8+ maps `LocalApplicationData` to `~/Library/Application Support` on
  macOS (Linux: `~/.local/share`). Documentation must say so.

## Tests

**FIVE test projects**, all xUnit and all run by the one `dotnet test Classic-Repair-Toolbox.slnx`, each in a
process of its own and all at once: [tests/CRT.App.Tests/](../tests/CRT.App.Tests/) (the app's `Handlers/` and the
Maintainer tab's logic in `Maintainer/`), the headless UI tests in two -
[tests/CRT.App.ContributeUiTests/](../tests/CRT.App.ContributeUiTests/) (the Contribute tab: the component pane, the
table editor, the drafts, the Submit dialog and that flow's windows) and [tests/CRT.App.UiTests/](../tests/CRT.App.UiTests/)
(every other tab and window, and `Main`) - [tests/CRT.Data.Tests/](../tests/CRT.Data.Tests/) and
[tests/CRT.Server.Tests/](../tests/CRT.Server.Tests/). None needs an oscilloscope, a MiniPro, a display, a database or a
network. Read real counts off a run; do not quote them from here. **The UI tests are two projects because one Avalonia
session serves a process on ONE dispatcher thread** (owner request, 2026-11-08: "yes, do this", to running the App tests
in parallel): in one project with the rest they were 3 min 58 s of a suite nothing else could shorten, and as one
project of their own still 3 min 31 s; as two, the whole suite took 161 s instead of some 240 s. **A new UI test goes
where its tab's are, and the two stay about as long as each other** - split them again by tab, never at random, if one
grows well past the other. Helpers both App test projects use live in `CRT.App.Tests/Shared/`, the headless harness in
`CRT.App.UiTests/Shared/`, each compiled into the projects that need it (their csproj) - never a library, which would
need CRT's internals opened to it too.

**The framework is xunit v3 (4.x), as the `xunit.v3.mtp-off` package, on VSTest** - so `dotnet test`,
CI, the release gate, the Stop hook and `coverlet.collector` coverage all work as on xunit 2. Moving
to Microsoft Testing Platform is a change to all of those at once, not a package bump. **xUnit1051**
(pass `TestContext.Current.CancellationToken`) is in `NoWarn` in all five projects.

**v3 RANDOMISES test order on every run**, so a test leaking state into whichever test follows fails
(or hangs) at random. The run header prints the seed, and running the test executable directly
replays that order and names a hang:
`tests/<project>/bin/Release/net10.0/<project>.exe :N -reporter verbose -longRunning 30`.

**A test PROCESS that dies is `[FATAL ERROR] Xunit.Sdk.TestPipelineException`, and its project still says "Passed!"**,
counting only the tests that ran; `dotnet test` exits 1 (v3 runs each assembly as a child process). Windows' Application
log names the cause (".NET Runtime", event 1023). **Since 2026-11-02 this machine's .NET 10.0.14 processes die now and
then with an internal error, 0x80131506** - the test processes of every project and the compiler server (VBCSCompiler)
alike, so no test's doing: a garbage collection finding its heap damaged (a dump of one that hung instead: a GC thread
looping in `plan_phase`/`find_next_marked`, the frames of dotnet/runtime#122528, whose reporter traced it to faulty RAM).
Run the suite again; the Stop hook says so when it sees one.

**Performance, measured - keep all four:** every `ExcelPackage` is made by CRT.Data's
`EpplusLicense.NewPackage`/`OpenPackage`, which turn off EPPlus's forced GC on dispose (it was ~85 %
of a 10-minute suite, and cost CRT on the UI thread per board read) - `EpplusPackagesTests` fails on
`new ExcelPackage(` elsewhere in `src/`. A test that WRITES a board workbook as set-up goes through
`CachedWorkbooks.Write` (saves under test stay real; `CachedWorkbooksTests` holds the key complete).
**Every test project runs with `TieredPGO` off and the server garbage collector** (their csproj; owner request,
2026-11-07: "please optimize this" - the app tests 6 min 5 s, then 3 min 58 s): a short test process never earns back
tiered PGO's instrumented code (a headless UI class 43 s -> 33 s), and the server collector collects on every core
(-> 29 s); CRT itself keeps the defaults. Measured again with the five projects side by side (2026-11-08): the
workstation collector 176-188 s against 161 s, and **never cap the server collector's heaps** - `GCHeapCount=2` was
faster and crashed or hung the runtime in 3 of 6 runs (see the dying processes above). **A sequential chip's every-order check follows each (state, events done)
once** (`IcTestCatalogueShippedDataTests.EveryOrder` - after any set of a vector's pin changes the socket is the same
whatever their order, so the state alone tells two orders apart; 8! orders replayed one by one took the 74LS373 eleven
seconds), and the simulated chip finds a pin by name in a table and works out a good chip's response once. Windows a
test leaves open were measured too: closing them all saved nothing. For the next slowdown, measure before guessing: run a class alone against its time in the full run,
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
- `Logger` writes nothing until `Logger.Initialize()` - **never call it from a test.** A test about the LEVEL a line is
  logged at reads it through `LoggedLines.Watch()` (CRT.App.Tests; `Logger.SeenForTests`), in a `using`.
- `WorklogManager.LoadFrom(root)` - **never call `WorklogManager.Load()`.** `TabWorkbooks.BoardKeyOverrideForTests`
  and `CurrentBoardDataOverrideForTests` stand in for `Main`.
- `ReviewSessionStore.InitialiseAt(path)` - **never call `ReviewSessionStore.Initialise()`** (it
  resolves the real file, which may hold a live maintainer token); point it back at nothing
  (`InitialiseAt(string.Empty)`) when done. Its tests share the `"ReviewSessionStore"` collection.
- **`Main` counts no board view until `StartAsync` has run** (`StartCountingBoardViews`), so a test loading real
  published boards through `Main` sends nothing to the live service - the view timers outlive the window on the
  tests' shared dispatcher. Never open that gate anywhere else; `MainBoardSwitchTests` asserts it stays shut.
- `TabMaintainer.UseRememberedChoices(...)`, `UseRememberedSelections(...)`, `UseTabBadge(...)` - how
  `Main` hands the Maintainer tab its settings; a tab a test builds never touches `UserSettings`.

Tests that mutate shared statics live in their collections - `"UserSettings"`, `"DataManager"`,
`"BoardData"` (`BoardDataReader`'s static cache), `"ReviewSessionStore"`, `"HeadlessUi"`. **Any new
test class touching one of these statics must join its collection.** Collections do NOT run in
parallel with each other (`xunit.runner.json`, `"parallelizeTestCollections": false`, in every project) - keep that,
it removed a whole class of race. It costs nothing measurable since the projects run side by side (2026-11-08): the two
UI projects decide the wall clock, and CRT.App.Tests and CRT.Data.Tests finish under them. **The projects share no
static - each is a process - but they share the DISK**: a test writes only under its own `TempWorkspace` (or a folder
of `Path.GetTempPath()` with a GUID in its name), never a fixed path another project's test could be writing at the
same moment.

### Headless UI tests

`tests/CRT.App.UiTests/` and `tests/CRT.App.ContributeUiTests/` build the tabs, windows and `Main` itself through
Avalonia's headless platform - no display, so they run on CI. `TestAppBuilder.cs` (in `CRT.App.UiTests/Shared/`, as
`UiTest.cs`) is a `CRT.App` subclass whose
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
- **A seam handing out the task a handler ALSO awaits does not wait for the handler** (2026-11-01: three tests failed at
  random, under load). `LastModelSaveForTests`, `SaveInFlightForTests` and the like are the inner save: already done, the
  test's `await` carries on before the handler's continuation - posted to the dispatcher, perhaps not yet - has run, and
  `RunJobs` does not close that gap. Poll for what the continuation does. Never prove "off this thread" by comparing thread ids
  (a pool thread free at its await runs the work itself - 7,258 of 20,000 when measured), and wait for a background read
  before deleting its file (Windows refuses). **A wait that gives up says so** (the tabs' `Until` asserts): half a second and
  silence failed later as something else.
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
  (`Controls/BoardTable/`), `BusyOverlay`, `ListRowDrag`, `PinoutDrawing`, `ColorDropperButton`. **They touch nothing of CRT's
  orchestration - no `Main`, `DataManager`, `UserSettings` or `Logger`, nor `ExternalTargetLauncher`, which logs**; a
  host hands them what they need (`BoardTableEditor.OpensWikiPage`, set in `App`'s static constructor), and CRT.Data's
  `CrtLog` is their logging seam. `SharedControlsIndependenceTests` enforces it.
  A control used by one tab stays in that tab's folder.
- **A color picker's list has a dropper in it** (owner requests, 2026-11-03, "a general thing": "I really miss the 'color picker'
  option", and then "merge it with the other color panel next to it" - beside the picker, as large as it, it took the slider's room;
  Avalonia's has none - `ColorDropperButton`, `ColorPickerDropper` puts it into every `ColorPicker`'s flyout as its template is
  applied, registered in `App`'s static constructor, no markup at any picker; the "Draw trace" palette, no `ColorPicker`, has its
  own beside its colors; the pure part `ScreenPixels`): pressed, it closes the flyout (a window of its own - `FlyoutToClose`,
  `Anchor`) and a click takes the color under the
  pointer from a picture of CRT's OWN window (`RenderTargetBitmap`, 5 ms for 1920x1080 - no common way to read the rest of the
  screen on the three platforms). Esc, another button or the window losing the front let go. **While it is on the keys are its
  own**: a handler taking a key whatever is handled asks `ColorDropperButton.IsAnyPicking`, or its Esc ends "Draw trace" too.
  **Size an overlay by the window, never `OverlayLayer.Bounds`**: a layer nothing has used yet has no size until laid out.
  **A style for an icon goes by a class of the icon's own,
  never "this control's TextBlocks"**: a tooltip is a child of its control, and read as Font Awesome's symbols.
- **A sentence says its own bold pieces as `<b>...</b>`** (2026-11-03; `BoldRuns.ShowMarked`, the one parser - `Marked`,
  `Plain` for a test) wherever a wording class holds no control: Board status's explanations, the picture tools' help (the keys in
  bold), the KiCad summary. A block showing one has `Text == null` - read it with `BoldRuns.TextOf`.
- **No button label has "..."** (owner rule, 2026-10-03). A button says what it does in plain words,
  whether or not it opens a window. Status lines and placeholders are not buttons. **Machine-checked:**
  `ButtonLabelTests`. A Wiki page naming a button uses the same label.
- **On the Contribute tab every entry that moves up and down stands in a frame of its own** (owner request, 2026-11-08:
  "everything that can be moved up/down must have its own panel with a border"): `Border.MovableRow` around the row of
  every list with a drag handle - a component's files and links, the board's files, links, credits, important signals and
  categories - and `Grid.MovableRowHeadings` on the headings over it, inset by the frame's edge and padding (9px) so they stay
  over their boxes. **Machine-checked:** `ContributeMovableRowsTests` (every `RowDragHandle` in `Tabs/Drafts/*.axaml`).
- **A focus moved by a posted handler asks `IsKeyboardFocusWithin` first** (owner report, 2026-11-08: the component popup's
  "Edit component" needed two clicks): Avalonia's `Button` lets go of its press when it loses the focus, and the popup's
  `Activated` handler - raised by the very press that brought it to the front - moved the focus after the press put it on
  the button. `ComponentInfoWindowTests.Edit_component_clicks_on_the_press_that_brings_the_popup_to_the_front`.
- **Buttons side by side are 6px apart, everywhere** (owner request, 2026-10-30: "the same horizontal
  spacing ... everywhere in UI"): a row's `Spacing`/`ItemSpacing` is 6, a row spaced by margins has a
  right margin of 6, two buttons in a Grid stand either side of a 6px column; code uses
  `Handlers/Theme/ButtonRow.Gap`. A gap that SETS A GROUP APART (the header's "UI" and "Table") is wider on
  purpose. Segmented switches, screen tabs, link buttons and a note box's row of format icons (`FormatTool`) are not such rows. **Machine-checked:**
  `ButtonSpacingTests`.
- **A button's colour says what it does, and only two do** (owner decision, 2026-10-13: "Throw-away
  and danger equals this IndianRed, while Accept/save/update is a green'ish button"). App.axaml's
  `Accept` class (green, `Button_Ok_*`) for a button that goes ahead - saves, submits, sends, adds,
  creates, applies, approves, publishes, installs; `Danger` (IndianRed, `Button_Cancel_*`) for one that
  throws something away or turns it down - deletes, removes, discards, rejects, withdraws, pushes back,
  resets. Everything else is plain (Cancel, Close, Browse, Open, Clear a selection, Sign in/out). The
  colour is the CLASS, never brushes on the button. A card's or a floating button's "stays green under
  the pointer" style names `.Accept` in its selector. **Machine-checked:** `ButtonColourTests` (markup);
  a button labelled in code carries the class and its own test checks it.
- **Everything only the administrator can use carries Font Awesome's padlock** (owner rule,
  2026-10-04): `<TextBlock Classes="AdminLock" Text="&#xf023;" />` right after its name, NO tooltip,
  and the entry hidden from everybody else. Today that is the Maintainer tab's "Management" screen, whose
  TAB carries it - its entries are the administrator's alone, so none has a padlock of its own (owner
  request, 2026-11-07: "no need to have a 'lock' icon on each individual menu item");
  `TabMaintainerModesTests` lists the Account and Management entries by name, so a new entry has to be
  put on one or the other.
- **Shift+Delete cuts in every text box** (owner report, 2026-11-07: it deleted without copying; Avalonia
  12.1.4 lists only Ctrl+X as Cut): `Handlers/Theme/TextEditingKeys.UseShiftDeleteCut`, added to the
  platform's keys once in `App.Initialize` - not on macOS, which has no such key.
- **A note people write at length is a `RichNoteTextBox`** (owner request, 2026-11-07: "a very simple
  WYSIWYG ... Bold, Italic, Underscore, Strikethrough, Bulletlist (nested) ... nor should it feel 'heavy'
  and slow"; `Controls/RichNote/`): a workbook's note, a worklog's description, work done, a comment, a
  photo's or file's comment, a device's description, and a component picture's and baseline's note. It
  IS a TextBox - typing, caret, selection, clipboard and IME are Avalonia's - whose presenter
  (`RichNoteTextPresenter`, `CreateTextLayout` overridden) draws each character's style in the very layout
  the caret is measured against, under Fluent's own template copied whole (`RichNoteTheme.axaml`, with a
  row of five `FormatTool` icon buttons over the text; compare it with Avalonia's TextBox.xaml when
  Avalonia is updated). **Code reads and writes `Markup`, never `Text`** (the words without their
  formatting), **and trims it with `RichNote.TrimMarkup`, never `Trim()`** (code review, 2026-11-08: a first line's
  indent says whether it is a bullet and how deep - "  - sub" saved trimmed reopened a level up) - where it is saved,
  compared (a component picture's and baseline's note: `ComponentPaneEdits.NormalisedMarkupNote`) and shown alike. **A board
  row's note is marked `[RichNoteMarkup]`, its column a sheet's `NoteColumns`** (owner decision, 2026-11-08: an indent-only
  change "is a change like any other change so it should count"): `BoardDataDiffer.ComparableText` compares it so, the table's
  cells keep and compare it so, and `BoardWorkbookSchemaTests` holds the attribute and the columns together. **What is stored is readable text, CRT.Data's `RichNote`**: `**bold**`, `*italic*`,
  `__underline__`, `~~strike~~`, a line starting "- " a bullet, two spaces a level (at no other indent - one
  space or twelve is the dash it is, `IsBulletIndent`; code review, 2026-11-08), "\" before a mark
  character to keep it - written with the fewest marks, and a note typed before reads as it was typed
  (measured 2026-11-07: none of the 69,839 shipped texts nor the owner's worklogs holds anything read
  otherwise). **A "\" is an escape only where it changes the reading** (`RichNote.IsEscapeAt`; code review,
  2026-11-07: "\\NAS\c64" read as "\NAS\c64", "D:\_old" as "D:_old", and the first edit saved them so) - before a
  star, a doubled "_" or "~", or a run of backslashes ending at one - and Serialize writes by that very rule
  (`NeedsEscape`); **nothing inside a web address is a mark** (`LinkParts`: "...#object.__init__" stays the
  link it is), only marks at its very end may close one opened before it, and a style covers an address whole
  or not at all (`Tidy`). **Line ends are "\n"**: the box sets its `NewLine` (a TextBox's is `Environment.NewLine`,
  "\r\n" on Windows, and every note of two lines was stored with a "\" before each literal star - code review,
  2026-11-07), and a "\r\n" pasted in is written as "\n" (`Tidy`). **Text typed in a PLAIN box is shown through
  `TextLinkRenderer.PlainLinkText`, never `LinkText`** (the "Resources" tab's notes): `LinkText` reads formatting. Every rule of an edit is `RichNoteEditing` (typing takes the style before it on its line;
  Tab/Shift+Tab nest a bullet, Enter continues or ends the list, Backspace at its words takes it away);
  **a change's styles follow from where it BEGAN** (`RichNoteEditing.ChangeOf`; code review, 2026-11-08: compared as text
  alone, "b" typed after a bold "a" in "ab" took the plain "b"'s style) - the selection's start as `Text` changes, which
  Avalonia 12.1.4 sets before it moves the caret, a Backspace noted in `OnKeyDown` as it takes out what stands before;
  the box keeps its OWN undo, as the TextBox's undid the text and left its styles behind. Shown by
  `TextLinkRenderer` (styles merged with links and search marks), the PDF export
  (`ComposeLinkedText`), searched and listed as `RichNote.PlainOf`. **A TextBox lets go of its
  selection as its `CaretIndex` is set** (Avalonia's `OnCaretIndexChanged`) - set the caret first, then
  the selection.
- **Every tooltip opens at its control's EDGE, never under the pointer** (`Handlers/Theme/ToolTipPlacement`,
  registered in `App`'s static constructor). Avalonia's default could land a tooltip over the pointer,
  and CRT's tooltips are popup windows, so the next click went to the tooltip (Approve needed three
  clicks). **Do not set `ToolTip.Placement="Pointer"` or a positive `VerticalOffset` on a control.**
  `ToolTipPlacementTests` reproduces the lost click.
- **EVERY WARNING AND ERROR REACHES THE SCREEN, NOT THE LOG ALONE** (owner request, 2026-11-07: "I think we must show
  every ERROR and WARNING in here in the UI"; "implement what you believe is the most optimal solution as seen from a user
  perspective"). Logger hands every WARN and CRIT line to `ProblemsLog` (the run's, `Current`; only once `Logger.Initialize`
  ran, so a test fills nothing). **A CRIT line is an ERROR, which Main's "Something went wrong" card says at once**
  (`Main.Problems.cs` - in plain words, `ProblemsLog.ReadableOf`: the first line, no exception type names; how many since it
  was closed; until closed), **unless the place it happened says it itself: `Logger.Critical(message, alreadyShown: true)`**
  (the data check's manifest, the update install, the Submit dialog's outcomes) - so a new failure a person must know is
  logged CRIT, and one a person needs to do nothing about WARN, which is only listed (`AtomicJsonFile.Write(...,
  needsNothingDone: true)` for a rebuildable cache). **CRT.Data's `CrtLog.Critical` raises the card too** (through
  `AppLoggerAdapter`), so a class moved there keeps the level it had - KiCad files that cannot be read are WARN (code
  review, 2026-11-08: CRIT since the loader moved, for data nobody at the keyboard could mend). **A file read again at every refresh is said ONCE per state of the
  file** (`UnreadableFiles.Report` - a damaged workbook, worklog or device; code review, 2026-11-07: logged at every read, the
  card came back as soon as it was closed; CRT.Data's since 2026-11-08, so a damaged board file too - `BoardDataReader`,
  `BoardFiles`). A test of CRT.App.Tests watching the lines logged (`LoggedLines`) sees CRT.Data's at CRT's own levels. Errors logged before the window was there are said as it opens, **taken with what
  comes after in ONE step** (`ProblemsLog.Watch` - listening, then reading, counted an error logged in between twice).
  **"Show errors and warnings" is the FIRST of the "Configuration" tab's buttons, with a badge per kind while the run has any** - errors
  in the danger color, warnings in the warning color (owner request, 2026-11-09; `TabConfiguration.WatchProblems`, from `Main.WatchProblems`
  - NavBadge's "error" and "warning"); none of those five buttons has a tooltip (owner request, same day).
  **"Errors and warnings"** (`ProblemsWindow`, from the card's "Show details" and the
  "Configuration" tab's "Show errors and warnings") lists them newest first, with "Copy all" for the "Feedback" tab. **A fatal
  crash is said at the NEXT launch, once** ("CRT closed unexpectedly" - `CrashReports.LatestFatal` over the crash file,
  `UserSettings.LastCrashReportSaid`), its report listed - **and never with nothing stored**: the crash file is never emptied,
  so the first launch of a CRT that keeps the setting stores its own time and says nothing (code review, 2026-11-07: it said
  September's crash as "the last time it ran"). **The data check never says "up to date" over files it could not
  download** (`DataCheckNotice.After(..., failedCount)`, counted by `OnlineServices.FailedDownloads` across the check - **the
  launch check's from the splash on**, `DataManager.FailedDownloadsAtLaunch`, as the main file and the board files sync there).
- **Main's notices FLOAT over the tabs' bottom-right corner, never in a row of their own** (owner
  report, 2026-10-11; `NoticeStack` in Main.axaml, `Main.Notices.cs`). As rows above the tabs, every
  notice that came or went moved the window, and RESIZED the schematic, which is fitted to its
  container. The right panel has two rows, worklog bar and tabs (`MainWindowTests` holds the tabs still
  while each card comes and goes). A new notice is a card of its OWN in the stack (no more sharing
  one card's text between flows), added to `NoticeNames` and to the demo (`--demo-notices`; its test
  fails on a card it never shows). **A card stays when there is something to do or somebody asked,
  and otherwise goes by itself** (`HideNoticeLater`): the data check's "all up to date" goes, its
  "please refresh board" stays (`DataCheckNotice`). Two exceptions, for the same reasons: "your settings could
  not be read" (2026-10-13) and "the list of what you have sent could not be read" (2026-11-02) stay until
  closed (`Main.SettingsNotice.cs` says why). The launch check MUST say something - silence
  was reported as a bug. The stack is at most half the tabs' width (`NoticeStackGeometry`), and a
  card that has just appeared lets clicks through for half a second, the tooltip lesson again.
  **The data check's card says the file downloaded on a line of its own, never wrapped**
  (`PathSegmentEllipsis` - the middle folders go, never the file name), **and keeps its widest for
  `Main.DataCheckWidthHeldFor` (3 s) after it last grew** (owner report, 2026-11-01: it flickered as
  every file's path made it shrink and grow; `WireDataCheckWidthHold`, the notices' own timers).
- **Every wait the user watches runs under `BusyOverlay`** (owner decision, 2026-09-28). One per
  window, the last child of its root Grid (`Main`, and each dialog that waits); the Maintainer tab
  has none of its own. Find it with `BusyOverlay.For`/`RunAsync(this, ...)`. It blocks clicks AND
  keys, dims after 300 ms, and gives up after `WaitLimit.Maximum` (2 minutes) **with nothing
  happening** - `WaitContext.Report` restarts the clock. **A timeout is never "it failed"**: a call
  that changed something on the server looks again and says what it found (`WaitWording.AfterTimeout`;
  `ServerWait` turns it into `ReviewApiFailure.TimedOut`); local work uses `RunLocalAsync`. Several
  steps are one wait through `HoldAsync`. `BusyOverlayHostsTests` guards the hosts. **Not under it**:
  background work nobody pressed for, board switching, the Submit dialog's upload (own progress), the
  MiniPro test, the scope's live output, the synchronous table LOAD, and a draft saving itself (the
  Drafts table's session, a second after a change). Waiting for that save because the contributor
  pressed something - closing the table, submitting, quitting - IS under it (`TabDrafts.SaveTableNowAsync`).
- **EVERY TAB BUT "Resources", "About" AND "Maintainer" has an INTRODUCTION TOUR, the first time it is shown** (owner
  request, 2026-11-11: "I do not like Wiki documentation (too much to read), nor creating YouTube videos ... the tool itself
  provide basic introduction per tab"; decisions: "Everyone should see it", "this is a view-demo only"; "Schematics",
  "Contribute" and "Workbooks" first, the rest the same day - "OK, now do the rest" - and 2026-11-12: "no need to have an
  introduction for the "About" and "Resources" tabs", the Maintainer tab "much more limited and specialized";
  `Main.Tours.cs`, `Main/TourOverlay`, every step `Handlers/Tours/TourCatalogue`): the window dimmed but for an amber
  frame round what a step is about, a card beside it, nothing under it clickable. **A new tab gets a tour or goes on the
  list of tabs without** (`MainToursTests.Every_tab_of_the_window_has_a_tour_or_is_left_without_on_purpose`,
  `Main.TourTabItems`, `TabItemsWithoutTour`), and the first introduction taken ends on a card saying the other tabs have
  one, framing them (`OtherToursCard`, only while no other tour has been seen). **A tour grows with CRT, and a change
  to a tab's UI is checked against its tour in the same change**: each step carries the RELEASE version it came with
  (`Since`; `TourCatalogueTests`), CRT remembers per tab the newest version seen (`UserSettings.IntroductionToursSeen`),
  and somebody who saw an older one gets "Tour of what's new" - every step newer than it, of EVERY version skipped, in the
  tab's order (`TourPlan`). So: a feature worth showing gets a step with the version it ships in (`InIntroduction: false`
  when it is for those who update only - a first look stays at most seven steps of its own, `TourCatalogueTests` - the first introduction taken ends on one card
  more, on the other tabs, which is no step of the tab); a step about something that CHANGED gets
  the new version, so it counts as new again; a step about something gone goes. A step names its controls by x:Name
  ("Tab/Name" inside a tab) - `MainToursTests` fails on a name the window no longer has; one not on screen as the tour
  starts is left out of that showing, a paragraph about a control not on screen too (`TourPart.Shows`) - **and what a
  showing left out, or said only in part, is remembered by its headline and comes back WHOLE the first time the tab is
  shown with all of it on screen** (`UserSettings.IntroductionTourStepsLeftOut`, `TourPlan.LeftOutBefore`, the kind
  `LeftOut`; code review, 2026-11-12: "Export to PDF" never reached a new user with no workbook). So no two steps of a
  tour share a headline (`TourCatalogueTests`). A tour waits while a tool of the schematic image is on (a picture tool,
  a worklog's area, "Draw trace"). A step may set the window up for itself (`TourStage` - the "Schematics" tour's
  panels shown expanded, owner request 2026-11-11), put back however the step is left and NEVER stored (the panels'
  "...PanelExpanded" settings are not written). A step's body may mark bold pieces `<b>...</b>`. "Take the introduction tours again" (the "Configuration" tab) forgets them all.
  Tours start only once `StartAsync` ran (`LetToursStartForTests` in a test), so no other test of `Main` meets one.
  **Kept in step with the UI two ways, as the Wiki is** (owner question, 2026-11-12: "Isn't Wiki a hook that never can be
  forgotten/missed, and should this also then be?"): **every bold name in a tour is the text of a control** of CRT's
  window, the worklog editor or "Errors and warnings" (`TourLabelsTests` - a button renamed fails it; keys and typed
  examples are its `NoControls` list), and the `Stop` hook [hooks/remind-tours.sh](hooks/remind-tours.sh) names the tours
  whose tab's code changed while THEIR part of `TourCatalogue.cs` did not - per tour, as the Wiki's goes per page - and it
  WARNS, as the Wiki's does. A tab given a tour goes in its `MAP`.
- **A picture is moved ONE way everywhere, the Schematics tab's** (owner request, 2026-11-06: "I do think we need to unify
  how pan works, and then have the same pan functionality as present in the 'Schematics' tab"; `Handlers/Geometry/PanGesture`):
  a drag with the RIGHT or the MIDDLE button pans, the wheel zooms at the pointer, and the left button is left to what it does
  there - the Schematics tab in all its tools (aligning KiCad data and marking components were its `TryBeginPan` already), the
  component popup's picture (the right button alone until then), the Contribute tab's previews of the schematic images
  (`PicturePreview`, a left drag until then - Ctrl+wheel still zooms, the wheel scrolling the cards) and the photo viewer
  (`PhotoZoomView`). A new picture that pans asks `PanGesture.IsPanPress`, and its tip says `PanGesture.Hint`.
- **User-visible punctuation is plain ASCII** (hyphens, straight quotes) unless an app label already
  has a character such as "·".
- **A tab, a menu or a button named in the UI is in quotes** (owner request, 2026-11-01: "when there is a reference to a
  tab, a menu or a button then it must be encapsulated with """): `the "Schematics" tab`, `under "Schematic images"`, `press
  "Align KiCad data"`, `"Undo" takes it back` - sentences, tooltips, a label's own words (`Enable "Workbooks" tab`) and the
  server's. A name that only opens a heading or a title ("Align KiCad data on [Top]") is the label itself. **Machine-checked:**
  `UiReferenceQuotingTests` (every tab, and a list of names that cannot be anything else, a sentence carried over to the next
  line read as one); a word that is also plain English
  ("the board's KiCad data", "Try again in a moment") is left to whoever writes it. **The introduction tours are the
  exception: they name a tab, a button, a panel or a key in BOLD, never in quotes** (owner request, 2026-11-11: "in all
  introductions, then try and use bold instead of encapsulating in """; `<b>Contribute</b> tab` - the test takes a bold
  name as marked, and `TourCatalogueTests` refuses a quote in a tour).
- **User-visible text says "color", never "colour"** (owner request, repeated 2026-10-30: "always use the
  American 'color' spelling") - in the app, the undo steps' names and the Wiki alike. Code comments and
  stored names (`PictureColours`, `traceHighlightColor`) are not text anybody reads. **Machine-checked:**
  `AmericanSpellingTests` (string literals in `src/`, and the markup's text attributes); grep the Wiki
  by hand.
- **A component, a trace or a net is HIGHLIGHTED in every word a person reads, never "lit up" or "lights up"** (owner
  request, 2026-11-11: "Please always use the term "highlight" for it") - the app, the tours and the Wiki; "can NOT be
  highlighted" where a heading said "will NOT light up". Code comments and names (`LitColor`, "Lit" classes) are not
  read by anybody. `TourCatalogueTests` holds the tours to it; grep the rest by hand.
- **The board's pictures are "schematic images" in every word a person reads** (owner request, 2026-11-11: "Always
  refer to the 'picture' in the 'Schematics' tab as a 'schematic image'"): the Schematics tab and its tips, the picture
  tools' bar, the Contribute tab's cards of "Board: Schematic images" and the board status about them, the Workbooks
  tab's board pane, warnings the "Errors and warnings" window lists, the tours and the Wiki. **A component's pictures
  stay pictures** (its "Pictures" page, the popup's), as do a worklog's or a device's photos, the "Feedback" tab's
  pictures and the scope's screen picture - and the photo viewer's pan tip (`PanGesture.Hint`; the Schematics tab's is
  `SchematicImageHint`). Code names (`PictureToolWording`, `PicturePreview`, `PaneBoardPart.Pictures`) are read by
  nobody and stay. `TourCatalogueTests` and `PanGestureTests` hold the tour and the tip to it; grep the rest by hand.
- Biggest files, for context budgeting: `Tabs/Schematics/ComponentInfoWindow.axaml.cs`, `Handlers/Data/UserSettings.cs`,
  `Handlers/Data/WorklogManager.cs`, `Tabs/Schematics/TabSchematics.Worklog.cs`,
  `Handlers/Data/DataManager.cs` (~1,700-2,200 lines each). Read the part you need.

## Architecture

### Bootstrap (`Main/`)

- `Main/Program.cs` - entry point; initialises Velopack, then Avalonia.
- `Main/App.axaml.cs` - holds `AppConfig`, nearly every tunable value (file/folder names, URLs,
  timeouts, zoom limits, version helpers, the Wiki page names). **Check here before hardcoding a
  constant.** `App` wires theme application, global exception logging and the startup sequence:
  `Splash` -> `DataManager.InitializeAsync` -> `Main` -> fire-and-forget check-in. **The splash stays until
  the first board is drawn** (owner report, 2026-11-09: "a couple of seconds ... I see an empty UI";
  `Main.FirstBoard.cs`, `Main.WaitForFirstBoardAsync`, at most `AppConfig.SplashWaitsForFirstBoardAtMost`):
  the window is shown only once its lists are filled, its thumbnails placed and its picture read - it
  loads unshown safely, the picture fitted by layout. **Every way a board load ends says so**
  (`FirstBoardDone`) - a new early return in `ReadSelectedBoardAsync` must too, or the splash waits out
  its limit. **`Main.StartAsync` awaits nothing before `PopulateHardwareDropDown`** (2026-11-09: an awaited
  crash-file read there handed the window back to App with empty drop-downs and no board load begun) -
  start such work, and await it after. Thumbnails decode several at once
  (`SchematicThumbnailSet.DecodedAtOnce`; 635 ms one at a time, 276 ms four at once on the C64 250407).
- `Main/Main.axaml.cs` - the main window, the central controller for board selection, schematics,
  and cross-tab state. **Split into partials by area** (`Main.BoardSelection.cs`, `Main.Worklog.cs`,
  `Main.Maintainer.cs`, `Main.SubmissionChecks.cs`, `Main.ModeHint.cs`, ...); its `.axaml.cs` header is
  the file map. It reaches into `TabSchematicsControl` members, so changes there ripple here.

### Tabs (`Tabs/`)

One folder per tab - `About`, `Configuration`, `Drafts` (the Contribute tab), `Feedback`,
`Maintainer`, `Oscilloscope`, `Overview`, `Resources`, `Schematics`, `Workbooks` - each a
`UserControl` with its logic in code-behind. `Worklog/` holds the worklog's dialog windows (the
component editor's window, `Contribute/`, went in CRT 4.0's step 2 - the component pane replaced it).
**Conditional tabs:**
Oscilloscope (`ApplyOscilloscopeTabVisibility`), "Workbooks" (below) with the worklog bar
(`ApplyWorklogBarVisibility`), Maintainer (`EnableMaintainerTab`, off by default). Hiding a selected
tab moves the selection to the first visible one. **Code asks for a tab by its TabItem, never its header**
(`ReferenceEquals(..., this.WorkbooksTabItem)`): "Workbooks" was compared as text in two places, and the tab was
renamed (2026-11-06).

**The "About" tab shows the whole CHANGELOG.md below the credits, drawn as GitHub draws it** (owner request, 2026-11-06;
`TabAbout.Changelog.cs`, the parse `GitHubMarkdown`, the drawing `MarkdownRenderer`, colors `Changelog_*` in both themes):
the file is EMBEDDED by CRT.App.csproj (`CRT.CHANGELOG.md`) - read, never written, so "Hands off CHANGELOG.md" holds - and
so a CRT shows the changelog of its own version. **The tab is one virtualized list** - the page first, then one item per
block, built as it scrolls into view and only once the tab is first shown: drawn at once it was ~1,100 controls and up
to 450 ms of layout at every width change. Links open through `ExternalTargetLauncher`. **A version's heading reads
"Version **3.0.0**, 2026-October-11", only the number bold** (owner request, 2026-11-06; `ChangelogVersionHeadings` rewords
it as it is shown - `MarkdownHeading.BoldAsMarked` - CHANGELOG.md's "# 3.0.0, 2026-October-11" is never touched; any other
heading stays bold as a whole). **Home, End, Page Up and Page Down scroll the tab from wherever the keyboard is**
(owner request, 2026-11-06; `Main.KeyIsForAboutPage` -> `TabAbout.TakeScrollKey`): every click on the tab sends the focus
back to the component filter, so the keys moved its caret; they are the page's unless the focus is in another text box,
an open drop-down or a list on the left. End brings the last block into view first - the list guesses the heights of the
blocks it has not built, and the end worked out from the guess fell short.

#### The Contribute tab (`Tabs/Drafts/`, class `TabDrafts`) - local-first contributing

**One tab replaced the Contribute and Drafts tabs** (CRT 4.0 step 2, 2026-10-25 - the plan: "One tab
instead of Contribute and Drafts"); the class and folder keep the Drafts name. It is ALWAYS on screen,
its title IndianRed while a draft exists (`ContributeTabTitle`/`hasDrafts` in Main.axaml, set by
`Main.ApplyContributeTabState`), and it is about the board shown in CRT: Main hands it every board
it shows (`Main.ShowBoardInContributeTab` -> `TabDrafts.ShowBoard`, `TabDrafts.Workspace.cs`). From
top to bottom: **the tab's top line naming the board shown, "Improve data for the [Commodore 64] hardware and
[250425 (long board)] board.", its two names bold, then what a change does** (owner request, 2026-11-05;
`ContributeBoardParts.TopLine` - so nothing else says "your first change makes your own draft"), beside it "My
drafts" (a drop-down list that SELECTS that board in CRT and goes back to its placeholder), "My submissions",
"Add a new board"; the shown board's draft header, if it has a draft - the row's first line is its
summary, "18 rows changed" and the badges - and **under it ONE row of buttons, the header's toolbar**
(owner request, 2026-10-30: "Saved", Undo and Redo on one line with four buttons under them, "not
aligned or anything"): **"View" and the joined switch "UI" | "Table" on the left** (owner request,
2026-11-04: "as those circles looks just wrong ... as these UI buttons are more important"; `Button.Segment`,
App.axaml's since then, as tall as the buttons beside it), then on the right Undo and Redo, set apart, BEFORE
"Review changes", Submit and "Discard draft" (`HeaderToolbar` - NOT in the
row's template: its DataContext is the draft's row, set by `ShowHeader`, so its buttons are never made
again under a press; the two views show with no draft too, the rest only with one). **A draft saves itself, so "Saved" and "Saving..." are said nowhere** (owner request,
2026-11-04: "this is implied"): only "Could not save - ..." with Try again, at the row's top right, and with nothing to say the row is gone and the buttons move up. **"Review changes" is greyed out with
nothing to review** (`DraftListItem.CanReviewChanges` - the same "has something" Submit asks; owner
request, 2026-10-30). **Over a board with a draft, App.axaml's notice panel says the draft is local**,
seen by nobody until submitted and reviewed (`LocalDraftNotice`, words
`ContributeBoardParts.LocalDraftNotice`; owner request, 2026-10-30). Then a left column, **a tree** (owner decisions,
2026-11-01 and 2026-11-03; "Board contents" until then; every entry `ContributeEntry`, `TabDrafts.Entries.cs`): **"Overview"
first, then "Board" over "Name", "Regions", "KiCad" over "Files" and "Important signals", then "Schematic images",
"Categories", "Files", "Links" and "Credits", then "Components" over "Basics", "Replacement parts", "Pinout", "Pictures",
"Oscilloscope baseline", "Files" and "Links"** - "Board", "KiCad" and "Components" are headings, no entries. **A page's
headline and a sentence naming an entry say its group too** ("Board: Name", "KiCad: Files", "KiCad: Important signals",
"Board: Schematic images", "Board: Credits", "Components: Pinout" with none picked - the code and the notes below still say
"Board status", "Board name" and "KiCad data" for them; the page "Overview" is just that, owner request, 2026-11-03: it also
names a tab, so never a bare "Overview" in a sentence about the page). Counts beside them - the KiCad files (its FILES counted, read off the UI thread -
`TabDrafts.ShowKiCadCount`, `ContributeBoardParts.KiCadFileCount`; for a draft held in memory counted
again only when its file list changes or "KiCad data" catches up - `thisKiCadCountedFrom`, code review
2026-10-30: every edit in the pane walked the KiCad folder; the KiCad panel's Summary says where that
number comes from, `KiCadImportSummary.Files` by the same names), and the rest (`ContributeBoardParts`; KiCad data opens its panel, the
rest open in the pane). **The important signals are named `ComponentPaneWording.ImportantSignals`** where the entry is
meant, and "Important KiCad signals" in the plain words of the undo steps and "Review changes"
(`DraftStepWording.ImportantSignals`); the sheet keeps its stored name, "Important signals". **Its entries carry no
tooltip** (owner request, 2026-10-30; a component's page in the red that stops Submit says why). **Another board shown selects
"Overview"** (owner report, 2026-11-01). On the
right ONE of the table, the component pane, the KiCad data panel, Board name, a component's page saying where to pick one, or Board status. **No page there has a Close button** (owner request, 2026-10-30: "remove
those 'Close' buttons on all those pages, as they are not needed anymore"): "Board status" is the way
back (`TabDrafts.ShowBoardStatusAsync` - closes the pane, the KiCad panel and the table, each as its
own close does, and the components' list), and it is SELECTED whenever nothing else is open. Only the
new component's form keeps its Cancel (`ComponentPane.CloseRequested`) - it adds nothing. **The splitter beside the left column is a 4px column in the global GridSplitter style**, as
the Schematics tab's (owner request, 2026-10-30: "bulky and visible" - `TabDrafts.SplitterWidth`; the
previews' handle on a picture card is drawn the same, a 4px bar in the `Splitter` brush). **THE COMPONENTS ARE PICKED ON THE FAR LEFT,
WHERE CRT'S LISTS STAND** (owner request, 2026-11-03: "have those lists work as selector for the component. Only one component should
be possible to select at a time"; a search and a list beside the tab's column until then; `Main.ContributeLists.cs`), so the column has only
**"Add component", just below the "Board" panel** (on a low window the tree scrolls instead, `TabDrafts.AddComponentKeepsAtLeast`).
**THE TAB HAS LISTS AND A FILTER BOX OF ITS OWN IN THAT PLACE, NEVER THE OTHER TABS' LENT TO IT** (code review, 2026-11-04;
`ContributeCategoryListBox`, `ContributeComponentListBox`, `ContributeSearchTextBox`, looking the same by `ComponentRowTemplate` and the
`SideList` styles - lent until then, every switch swapped rows, selection mode and words, which took a guard wherever a list was touched,
and every review found another way to lose a selection in it): **showing the tab changes only which are visible**
(`FollowContributeTabInLists`), the keyboard in "the filter box" staying in it. The tab's show its `ComponentItems`/`CategoryItems`, which
follow the draft as it is edited (`BoardForList`) - the board on screen is read again only once the pane closes - and **the Schematics tab, the
Overview and the rest read only `ComponentFilterListBox`/`CategoryFilterListBox`, whichever tab is shown**. **ONE SELECTION FOR BOTH
TABS' LISTS** (owner request, 2026-11-08: "always keep the same selection for the left-side panel "Category" and "Component" lists,
meaning it should not change selections when shifting between e.g. "Schematics" tab and "Contribute" tab" - every category ticked the
first time and the tab's own pick until then). **AND THE SAME LISTS** (owner request, 2026-11-09: "Toggling between tabs (e.g.
"Contribute" and "Schematics") should keep the exact same components and categories selected, and it should not change position in the
lists ... it should be the exact same lists" - only what was chosen on the tab went back until then, each filter box its own and each list
at its own place): entering the tab takes EVERYTHING the other tabs' lists show - the categories ticked, the word in the filter box (it
narrows here), the components selected (`TakeListsIntoContribute`, `TabDrafts.TakePickFromOtherTabs` - opened where a component's page or its
table shows) - and leaving it gives everything back (`GiveListsBackFromContribute`): the ticks as the board's saved filter, as a tick there
is, the word narrowing their list WITHOUT selecting all it finds (typed there it does), the components picked selected and lit, and held
for the load that follows (`_selectionToKeep`). **Each list opens where the one shown before it stood** - the same row at its top
(`ListPosition`, `StandOf`/`PutAt`; the tab's component list has `AutoScrollToSelectedItem` off too, or a pick scrolled it after). **A side
nothing was done on hands nothing over** (`_contributeSaidAtEntry`, `_othersSaidAtExit` - `ListsSaid`): lists that are the same stay
exactly as they were, a pick a word hides on the tab included, which the other list cannot hold. Nothing is handed across while either set of
lists is of another board, or still being read (`ListsAreOfOneBoard`) - entered then, the tab takes them as its lists come
(`OnContributeListsShown`). What is ticked and typed here is kept per board for the run (`ContributeListState`): the categories UNTICKED and
the words typed, so a category the draft makes is ticked by itself, and a board shown here before keeps what its lists said here
(`ContributeListState.Kept`) through another board shown and back (`ShowContributeTicks`; a category going as the lists are read again is nobody's choice -
`IsReadingComponentLists`). A component with no category has no tick to hide it by, and stays listed. **CRT sets the filter boxes only through
`ShowComponentSearchText` and `ShowContributeSearchText`** (code review, 2026-11-04; measured): Avalonia POSTS `TextChanged`, so a flag
around `Text =` has been lifted when it comes, and 200 ms later the filter replaced the lists and let go of the selection. A `TextChanged` saying what the lists show (`_componentSearchInEffect`, `_contributeSearchInEffect`) is nothing to do. **A word in the tab's box is filed under the board it was typed for** (`_contributeListsBoard`, moved as soon as another board is chosen - `FollowBoardInContributeLists`), **and a load ending leaves a word not taken yet and takes it** (`UseContributeListsOfBoardShown`). **SEVERAL COMPONENTS ARE PICKED WITH Ctrl OR Shift AND EDITED AS ONE** (owner request, 2026-11-08: "I want to be able to mark multiple
components and when doing that I am bulk editing whatever data I changes"; the proposal agreed the same day - one at a time until then;
`SelectionMode="Multiple"`, `TabDrafts.thisPickedSeveral`/`PickedLabels`, `ComponentPane.Several.cs`, every rule CRT.Data's `ComponentBulk`):
a khaki bar names them - **the number and each name in bold, nothing between them** ("Editing **2** components: **RP7**, **RP8**";
owner requests, 2026-11-09) - and says every change reaches all of them, its second line what only some have on a list page or why a page is
held back, nothing on "Basics" ("1 field here differs ..." went the same day: "no need for that"; `SeveralComponentsWording`); the label is
off; a field they differ in is EMPTY, "Differs - N values" in it, so nothing is set until something is typed or picked, then on all - **and
WHILE IT HAS THE FOCUS, never under the pointer, their values are listed under it, each once with whose it is, narrowed to what is typed**
(owner requests, 2026-11-09: "show the values in a drop-down when I hover that field - just in the same way as in "Board: Links" for the
"Category" field", then "I do not want that mouse hover will trigger that list ... Only when cursor is active in a field it should show the
list ... where I can start typing also and then it will filter the list"; `DiffersPopup`, `ComponentPaneDifferingValue.Matching` - a note
with formatting listed and found by its words; the empty value too, which nothing typed could say): one clicked, or lit by Down/Up and
picked by Enter, is set on all of them - **the lines take no focus**, so the box keeps it and a click is no box left; **what is typed there
waits for Enter, a pick or the box being left, no typing pause** (half a word would go to all of them while the list is read), and the
list's keys are taken by the pane's own tunnelling key handler first (`TakeDifferingListKey` - its Enter would otherwise commit the typed
filter); a drop-down list's values come by the keyboard only, its own list opening on a click; a list page lists the rows ALL of them have (compared without
the label), a row changed, removed or added there is changed, removed or added on every one, and what only some have stays - no row of
several moves; **"Make all have the same data as [one] *1"** ("Make all like" until 2026-11-09 - owner request: "make it clear ... that it
then will affect only this data on this page"; the "*1" small, bold and raised, `FootnoteMark`) makes the page shown the same on all, ONE
step, the note under it saying the button changes this page alone (`MakeAllLikeNote`); **"This data is different per region" stays with
several picked** (owner request, 2026-11-09: "Can't it still have this, even if selecting multiple components?"): ticked for all, unticked
for none, half-ticked for some, a click making all of them so as ONE step (`ComponentPane.AllOrNone`, `SetRegional`); Pinout only while they share a package, Oscilloscope
baseline never (measured on one component) - neither offers "Make all like" then; nothing acting on one component as a whole is offered.
Each change is one step named after them all. **A click with Ctrl, Shift or Cmd is no click on the one picked** (owner report, 2026-11-08: "I cannot mark multiple components when I am in "Contribute" tab" - the list's Tapped, which opens again the one clicked, put the selection back to it after every Ctrl+click; `OnComponentsListTapped`), **and a pick overtaken by a newer one opens nothing** (`thisPickGeneration`, `thisPicksOpening` - opening is under the BusyOverlay and reads off the UI thread, so a click and a Shift+click right after it each opened, and the first, ending last, took the pick back to one; `HoldOpeningForTests`). **The other tabs' component list picks the same way** (owner request, 2026-11-08: "in "Schematics" then clicking one component only will TOGGLE that component on/off? I want it to work like "Contribute""): `ComponentFilterListBox` is `Multiple` - a click selects the one clicked alone, Ctrl+click adds or takes away, Shift+click a range. It was `Multiple,Toggle` from the first commit (2026-03-07), every click toggling. The category lists stay `Multiple,Toggle`: their rows are ticks. Tests of picking use real presses (`MainContributeListsTests.Ctrl_click_and_Shift_click_pick_several_on_the_tab`): `PickComponentsForTests` goes round the click. A page's problems are every picked one's (`ContributeProblems.OnEntry(..., selected)`).
**In "Table" the component sheets show the picked components' rows alone** (`BoardTableEditor.Components.cs`, `ShowOnlyComponents` - a
line over the table says whose; a sheet of the board as a whole is not narrowed), which narrows as a pill does - so "Replace" replaces in
them alone, a paste adds no rows, nothing moves. "Mark all" stays off on the tab. "Clear" goes back to how the tab starts - every category, nothing typed, no component picked, its page
saying where to pick one; the region buttons read `BoardForList` too, so ticking "Board: Regions" shows them at once. **THE OTHER TABS' LISTS STAY BEHIND THEM AS THEY WERE**, selection and scroll with them - **only a board read again reaches them, which drops a selection as any load does, so what was selected as the tab was entered is held** (`_selectionToKeep`, `HoldSelectionOfTheOtherTabs`; through every load while the tab is shown and the first after it is left - `LetGoOfSelectionShown` - until somebody chooses or types), **before all that a word in the box finds, nothing selected too** (`ShowListsOfBoard`, `ApplyRegionFilterAsync`; code review, 2026-11-04). **The board on screen read again keeps its rows until the new ones come, and those come selected in ONE go, the list where it stood** (owner report, 2026-11-05: "it slowly populated the list of components ... and then marked all components"; `LetGoOfListsForLoad`, `SelectInComponentList`; measured, 276 scrolls for 300 rows): row by row, each asked the list to scroll to it, a layout pass apiece. The list's `AutoScrollToSelectedItem` is off in the markup, as Avalonia also scrolls to a row not reached yet as the list comes back into view, so the one selection that must be seen asks for it (`SelectComponentByBoardLabel`). Rows kept like that have no board behind them for the moment `_currentBoardData` is null, so **whatever acts on the lists asks for it first** ("Clear" ended CRT; code review, 2026-11-05), and a read that finds no board lets them go (`LetGoOfOtherTabsLists`). **The rows kept take no click until the new ones come** (`OtherTabsListsTakeClicks`, looking the same - greyed out, the flicker would be back; second code review, 2026-11-05: a tick or a pick then was dropped without a word), **and take clicks again however the load ends** (`LoadSelectedBoardAsync` around `ReadSelectedBoardAsync`), or the lists would never take one again. **A load's last step lights up what the list has selected by then, whichever tab is shown** (`ShowHighlightsOfSelectionAfterLoad`), **and the end of the KiCad load lights its traces** (`LoadKiCadProjectForCurrentBoardAsync`; code review, 2026-11-05: that load lets go of them, and a selection kept came back with rectangles alone). **BACK ON THE TAB EVERYTHING STANDS WHERE IT STOOD** (owner request, 2026-11-04: "navigate to the
exact same position ... if toggling between the main tabs"): a tab not shown is only detached, so the pane, the left column and the table keep their
page, entry and scroll by themselves - measured, and held by `TabDraftsPositionTests`, as anything reading them again on the way back would lose it - and the lists are only hidden.
**The list stands outside the tab, so "CRT has to be updated" does not cover it**: a pick there opens nothing. **The entry open on
the right is shown selected, worked out from what is open every time** (`TabDrafts.EntryOpen`, `ShowNavSelection` - never a flag beside
the pane). **THE COMPONENT PICKED STAYS SELECTED IN THE LIST WHATEVER ELSE IS OPEN** (`thisPickedComponent`; owner request, 2026-11-03),
following a rename in the pane, until another board is shown or neither the board on screen nor the draft held in memory has such a component (`FollowPickedComponent`; the board on screen is behind for a moment once the pane or table closes); a click on it opens it again though no
selection changes (`OnComponentsListTapped`). **"Components"'s entries are ITS pages** (`ComponentPane.Page`, what each holds beside it -
`ComponentPane.PagesShown` from the rows on screen, `ContributeEntries.FactsOf` for one not open - and red while something on it stops
Submit); **with none picked the entry says where to pick one** (`NoComponentPanel`, `thisNoComponentPage`), and a component picked then opens on
that page. **"UI" OR "TABLE"** (owner request, 2026-11-03: "when I select e.g. 'Board: Schematic images' then I will see the table for
'Board schematics' in the content area, instead of the normal UI elements"; `ShowViewAsync`,
`thisTableMode`, for the run; "Table view" and "Close table view" until then, a try-out): **an entry with a sheet shows it, the sheet tabs
gone** (`ContributeEntries.TableOf`; `ContributeEntriesTests` fails on a sheet with no entry), **the left column STAYING** (2026-10-29's
"the table takes the whole workspace" is over), the same entry kept as the view is switched. **Five entries are no sheet** ("Overview",
"Name", "Regions", "KiCad: Files", "Categories"): they show their page under `NoTableNotice`. **The new component's form is no entry, so the switch is greyed out over it** (`EntryOpen` null; code review, 2026-11-04). **"Component images" is two entries**, one
sheet (CRT.Data's `BoardTableSheetParts`, below). A component picked with "Table" on goes to its first row of the sheet of the page it stood
on (`GoToRowsOfComponent`). **A pill or a search moving the table moves the entry** (`SheetShown`), and turns off the entries it shows nothing of,
as such a sheet's tab went (`ShowTableEntriesEnabled`) - **the five that are no sheet too, which neither looks in** (owner request, 2026-11-07;
`BoardTableEditor.IsFilteredByPillsOrSearch`); turned off, an entry is faded with no grey behind it. **What opens a page with a table of its own is the "UI" view, however it was asked for**
("Edit component", `OpenPaneOnDraftAsync`), **and a table opened from elsewhere the "Table" view** ("Show in the table", Submit on a draft
with errors, `OpenTableOnDraftAsync`). The tab's lists
are ONE collection each, changed in place (`InPlaceList.Sync`, never inside itself: a category that goes is a selection changed while it
is changing, which nobody chose - `IsReadingComponentLists`) - replaced, the list scrolled to its top and let go
of its selection after every save, **and with the list's own scrolling off while they change** (`ReadComponentLists`, owner report, 2026-11-04: every row put in above the selection moved its index and asked to be scrolled to it, 392 scrolls for 399 rows, seen as the list scrolling down slowly; it comes back to the selection once, only when that moved); a selection set from code to follow the pane opens nothing
(`thisFollowingComponentSelection`). **A selection the arrow keys move opens its component only once the
keys pause** (`TabDrafts.ComponentKeyPause`; Enter at once, a click at once - code review, 2026-10-30:
each component passed was opened and saved under the BusyOverlay, which took the next key);
`ShowNavSelection` leaves the list where the keys put it while that waits. **"Pictures" is "Schematic
images" wherever the part is named**
(owner request, 2026-10-29 - `ComponentPaneWording.SchematicImages`; the undo steps, "Review changes"
and the readiness step too; a COMPONENT's pictures stay pictures), and **Files and Links are two
parts** (`PaneBoardPart.Files`/`Links`, "Files and links" until then). **The left column's width is
dragged at its splitter and kept PER BOARD** - a board never dragged on takes the width dragged last
(`ContributeWidths`, `UserSettings.ContributeNavWidthFor`, `TabDrafts.Splitters.cs`; seams handed in by
`Initialize`, so a test's tab never touches the settings).
**The table and the pane are views of the board SHOWN** - showing another board closes them
(`CloseTableForAnotherBoard`, `CloseComponent`), and the left column reads the draft as they
left it (`BoardForList`, the pane's "Draft" marks in `thisChangedHere` until Main shows the board
again). **That close cannot be called off and asks nothing** (code review, 2026-10-26: Cancel left
the table on the previous board's draft under the new board's header): what the table holds goes
into its draft and a save starts; one that fails waits in memory for its board or quitting.
**Closing the table settles ITS board's draft only** (`CloseTableAsync` -> `SettleTableEditsAsync`);
`ConfirmLeavingTableAsync` - every draft - is quitting's. **One timer saves every draft no editor
holds** (`SaveDueDraftsNoEditorHoldsAsync` over `DraftSessions.Due`, started by `Initialize`); the
table and the pane save their own, whose Saved events the row and the board on screen follow. Every
pool save is `DraftSession.SaveAsync` - never another `Task.Run(session.Save)`. **A save waiting for the one in flight waits
only while that one is UNFINISHED** (`thisSaveInFlight is { IsCompleted: false }`, the pane's and the table's; owner report,
2026-11-06: CRT froze for good, a note typed and the "Resources" tab clicked): the finished save's own continuation clears the
field at the priority it began under - the save timer's, low - so a save asked for from a click came first, found it still
there, and `await` on a finished task looped for ever on the UI thread. A tool for the next hang: `dotnet-stack report -p <pid>`
(installed with `--tool-path`) names the method the UI thread spins in.

- **"Draft" IS THE ONE WORD FOR "DIFFERS FROM THE PUBLISHED BOARD"** (owner request, 2026-11-02: "Changed" until then in the
  Contribute tab's list and the pane's header; the component list on the left of CRT always said "Draft"), **and A NEW BOARD HAS
  NONE OF IT** - nothing is published to differ from: no mark in either list, on the pane's header or a board part, no tint on the
  Schematics tab (`DataManager`'s summary is Empty with no published board) and no "Undo changes to", in the pane
  (`ComponentPane.IsComponentChanged`) or "Review changes" (`ReviewChangesAsync`).
- **A NEW BOARD'S CREDITS START WITH THEIR ROWS** (owner request, 2026-11-02; CRT.Data's `NewBoardCredits`,
  `TabDrafts.NewBoardCredits.cs`): one per schematic image, "Board data", "Oscilloscope baseline" per region - to be filled in
  or deleted. Nobody named is an error on purpose (`BoardRowRequired`), so Submit waits. Offered as "Board: Credits" opens, as
  the table opens and before Submit - only while the board's entry AND the marker say it is new (`entry.IsDraftOnly`:
  the marker alone still says so once the board is published here), ONE step when anything changes. **Reading "Board status" changes
  nothing in the draft** (code review, 2026-11-02: an Undo there read the board again, which offered the very step back with the Redo
  gone, and its marker write ran into "Board: Name" moving the folder): it says how many rows are to come
  (`NewBoardCredits.RowsToAdd`, `BoardReadiness.CreditsToFillIn`). **The left column goes by what is held in memory on a new board**
  (`BoardForList`), so it counts them as they come. **Each row ONCE - what was offered
  is in the draft's marker (`DraftMarker.CreditsOffered`), so a row deleted or undone stays gone**; a row nothing was typed in follows
  the board (an image renamed or deleted, regions ticked), one with a name stays. **A step that took a row away, undone, puts the
  marker back** (`CreditsFollowedBoard`; code review, 2026-11-02 - the rows are the step's and what was offered the marker's), so the
  next offer follows again: the old name's row was left for good and the new name's never added.
- **Each replacement part, picture and oscilloscope baseline stands in a dark gray frame, "Applies to" its last field** (owner
  request, 2026-11-01; App's picture cards' look, `Border.PaneEntry`), and a picture's note is a box of several lines while empty too.
- **A COMPONENT'S PAGES ARE ENTRIES OF THE LEFT COLUMN, THE PANE HAS NO TABS** (owner decisions, 2026-11-01: "too massive, if all data is
  filled"; tabs over the pane until 2026-11-03: "no tabs in the 'component view'"; `ComponentPane.Pages.cs`): Basics (its rows: "1" when the
  region is not different, "2" if it is), Replacement parts, Pinout, Pictures, Oscilloscope baseline, Files,
  Links, each saying how many it holds and red while something on it stops Submit (no category, a pin past the case, a baseline
  lacking its signal name or image - said under it as it is shown - and a link's address `IsRefusedLink` refuses). One shows; **every section stays
  filled behind the one shown**, so what the pane commits is what it always was (a test reaching a hidden section's row
  CONTROLS calls `ShowPage` first: no container is made in a section not shown). **The page stays as another component
  is picked.** Its **label is a box under Basics** (owner report, 2026-11-01: "unable to rename"; there was none):
  renamed on Enter or when left, never on a typing pause, every row moving with it (`ComponentPaneEdits.WithLabel`,
  rectangles too - `DraftEditRules` follows only a row whose other cells stayed); a label another component has stays
  in its box with the reason; an undo of it finds the component again (`FollowRelabelInDraft`, below); the rows are not read again
  (`ShowRelabelled` - a pin just added and a baseline held back went with them, code review 2026-11-01). Worklogs on this
  machine keep the old label, as for a rename in the table.
- **A BOARD MAY COME IN REGIONS OF ITS OWN** (owner request, 2026-11-09: "do the region steps 1-9, so we are prepared for other
  regions" - for hardware no computer of the time: a stereo, a console sold as "EU" and "US"; `BoardRegions.Of`, THE ONE RULE): the board
  file's `regions` lists them in the board's order, PAL and NTSC or names of its own; a board that lists none comes in PAL and NTSC when a row
  says either, as before - a row's other word never makes a region. **Every published board looks exactly as it did.** A name is 1 to 20
  letters, digits, spaces and "- + / ( ) .", never "All regions", each once, at most 6 (`BoardRegions.NameProblem` - the pane, the table and
  the server's `regions.unknown` alike, server 10.2.0). **"Board: Regions" is ONE tick, "This board comes with multiple regions", and every
  region listed in its button's colors** (owner request, 2026-11-09: "have the checkbox state "This board comes with multiple regions" and
  then list all regions and its colors? It should be possible to reorder the list of regions up/down" - a PAL and NTSC tick with the board's
  own regions under it until then; `ComponentPaneRegionRow`, its pill `Border.RegionPill` styled by a class per set so the theme and a
  color set for PAL and NTSC follow): ticked on a board with none it comes in PAL and NTSC, unticked in none - not while a row says a region
  (`BoardRegions.WithRegions`); "Delete region" only while no row says it and more than `FewestRegions` (2) are left
  (`BoardRegions.CanDelete` - the last two go by unticking), why not said ONCE under it, its count bold ("**2** rows of this board have it
  applied, so it stays ..." - owner request, 2026-11-09: a count beside the pill too was "double-wording"; `theme:BoldRuns.MarkedText`, the
  attached property a template's block shows a marked sentence by), a drag handle or Alt+Up/Down on its row for the buttons' order, ONE step
  (`BoardRegions.InOrder`), "Add region" by name under the list. **The buttons under the component list are one per region, in the
  board's order** (`Main.LayOutRegionButtons`; PAL's and NTSC's are the named buttons they always were, a board's own made for it, three to a
  row - `RegionButtonLayout`, `Controls/RegionButtonGrid`). **Their grids' rows and columns are FIXED in the markup, never changed in code**
  (owner report, 2026-11-09: CRT threw at launch, `ArgumentOutOfRangeException` in `Grid.GetMeasureSizeForRange`): **Avalonia 12.1.4's
  Grid keeps each child's cell as last laid out and does not look again when its definitions change** - only a child added or removed,
  or its Grid.Row/Column, makes it - so a grid that lost a column measured a hidden button against it. The list's is six equal columns
  6px apart (`ColumnSpacing`), a button spanning three (the halves "*,6,*" gave) or two, over two rows; the popup's two Auto columns.
  A new grid whose definitions would change at run time is built the same way (`MainWindowTests`, `ComponentInfoWindowTests` lay them out
  shown, board after board). **And so on the popup** (stacked from its footer's bottom, PAL and NTSC where they
  always stood - `PopupCellOf`), each own region in one of four colors (`RegionChoice.ColorsOf`, `Schematics_Region_Own1..4`, the label over
  the picture too). "Applies to" lists the board's regions (`ComponentPaneRegionalRow.UseRegions`, `AppliesToChoices(regions, ...)`), a
  component told apart per region gets a row in each, and the device form lists them (`OwnedDeviceWindow.BoardRegionsAsViewedAsync`). **The
  region shown** is the one picked on that board when it is one of its own (`UserSettings.RegionOfBoard`, `regionByBoard`), else the one
  picked anywhere - `region`, which CRT 3.x reads, so it only ever holds PAL or NTSC - else the board's first (`RegionChoice.ToShow`, as a
  board loads). A change of the list is said by name ("This board now comes in PAL, NTSC and Japan versions" - `BoardRegions.ChangedLine`;
  the review's summary and the history carry `regionNamesBefore`/`regionNamesNow`, null from before meaning PAL and NTSC).
  **A REGION MAY HAVE COLORS OF THE BOARD'S OWN, IN THE BOARD DATA** (owner request, 2026-11-11: "set the background color, and then you
  should calculate the complementary text color ... I would also like to be able to set the text color"; asked, "In the board data"; every
  rule CRT.Data's `RegionColors`): the board file's `regionColors` (`region`, `background`, `text`, written only when any), "#RRGGBB" ONLY -
  the error `regions.color` on both sides, server 10.4.0 - an empty text being black or white, whichever has the higher WCAG contrast ratio
  (`ContrastText`; SeaGreen takes black by it). `SubmissionRows.RegionColors` null keeps what is published; `regionColorsChanged` on the
  review's summary and the history; a change of them is a change of the draft, said in "Review changes". A region deleted or all taken away
  takes its colors (`RegionColors.Tidied`). With none, a region keeps CRT's colors. Every place a region is drawn asks: the buttons by class
  `BoardColors` and their own resources (`RegionPaint.UseBoardColors`, styles in App.axaml after the per-region ones), the labels over the
  pictures by `RegionPaint.PaintLabel`, the popup by `ComponentInfoWindow.RegionColorRows` (set by Main before `SetComponent`), and "Board:
  Regions" by two pickers on each region's own line (owner request, the same day: "one line instead of 2 lines"), "Automatic text color" and "Reset colors", committed on the typing pause as a picture's colors are. **A
  region's pill is not bold** and is drawn as its lit button (2px edge in its background). **An own region's button shows its name in a
  `RegionButtonGrid.NameBlock(button)` bound to the button's text color** (owner report, 2026-11-11): App.axaml's style for every TextBlock
  painted it in the window's text color, black on the region's green. **Under each region its components are listed** (owner request,
  2026-11-11; `BoardRegions.RegionHolding.Components` - every component with a row saying it, those told apart in every region too: on a PAL and NTSC board that is nearly every one), the first seven as buttons (a hand over each) opening each on "Basics" in
  that region (`TabDrafts.OpenComponentInRegionAsync`), "... and **N** more" and "Show all" for the rest as the last item of the same line, the pill and both pickers
  one height (`ComponentPaneRegionRow.LineHeight`), and the board files and links counted.
  **The "Regions" entry of the left column keeps its word whole**, the names cut short with "..." in the room left (owner report, 2026-11-11:
  six long names took the row).
- **REGIONS** (owner decisions, 2026-11-01: PAL and NTSC or none, stored and published, on a page of their own;
  `BoardRegions` - THE ONE RULE, replacing `HasExplicitRegionComponents`, which asked the component rows alone, so regional
  baselines on a board with no regional component hid the buttons and one region's baselines for good). "Board: Regions" is
  one tick, `regions` in the board file (written only when set - every other board keeps its bytes; no conversion: a board
  without it is worked out from its rows), not removable while a row says a region. It travels as `SubmissionRows.Regions`
  (null keeps what is published, as the pins; a name `BoardRegions.NameProblem` refuses - anything but PAL or NTSC until 2026-11-09 - is the error `regions.unknown` on both sides -
  code review, 2026-11-01, Security: it reached the publish as sent, server 7.0.0), counts as a change of the draft (`DraftOtherChanges`, retirement, "made only
  to look") and is said in "Review changes", the review and the board's history (`RegionsNow`). **THE REGION IS SWITCHED BY THE BUTTONS
  UNDER CRT'S COMPONENT LIST, no others** (owner request, 2026-11-03: "use the regions from the left-side instead of adding new region
  buttons in the component pages"; `Main.PickRegion` -> `TabDrafts.UseRegion`): a line says which region a page shows where it depends on it, **the region in bold**, as a
  regional row's "The **PAL** version" line (owner request, 2026-11-09: "please bold the selected region"; `ComponentPane.ShowRegionShown`) - **Basics and Pinout while their tick is ticked, always on
  Replacement parts, Pictures and Oscilloscope baseline**, each listing what the region shown shows.
  **"Files" and "Links" LIST EVERY ROW of the component** (owner request, 2026-11-03: "the complete list can be viewed and
  sorted ... it should work in same was as 'Board: Files' and 'Board: Links'"; `ComponentPaneEdits.AllFilesOf`/`WithAllFiles`, `AllLinksOf`/`WithAllLinks`):
  a row given another region stays where it is, and moves past any row. What a region shows of them is the popup's
  (`ComponentPopupContent.LinksOf`).
  **Baselines, pictures and replacement parts have "Applies to" per row, as files and links do** (`ComponentPaneRegionalRow`; of the first three, a row moved out of the region shown is let go of here and said -
  left, it was filed again at the next commit; a new part with nothing typed stays until it holds something - `HoldsNothingYet`;
  **a new part given another region waits, out of the draft and on the page, until the focus leaves it**
  (`IsWaitingToBeTyped`; code review, 2026-11-02: the typing pause took it half typed from under the caret, and a Tab to its next box
  would have too) - a commit nothing can wait out, closing or handing over, takes it as it is;
  **the box's list stays while it holds the word**: a list replaced in the middle of a pick undoes the pick, so a test of a pick picks IN THE BOX - setting `row.AppliesTo` saw nothing wrong, code review 2026-11-02); **Basics and Pinout a "This data is different per region" tick**, as those are sets
  (`WithRegionalRows`, `WithRegionalPins` - no step, the reason said, for a component with no pins, or one taken off in a region
  with none, which deleted every pin; code review, 2026-11-01). A new baseline starts in the region shown, as 99% are; a new picture in all,
  as 99% are; a new pin with its set - **a component's row per region says nothing of its pins**, the tick does.
  **A FILE OR A LINK, OF A COMPONENT OR OF THE BOARD, CAN BE OF ONE REGION** (owner request, 2026-11-02: "a new column with the
  region"; `Region` on the four row types, `region` in the board file written only when filled, the last column of their sheets but
  the board's two, part of a row's identity only when filled - `BoardDraftNaturalKeys`). **A BOARD FILE AND A BOARD LINK HAVE A NOTE**
  (owner request, 2026-11-06: "describe what this link or file is about ... For now I do not want this on the component files and
  links"): `Note`, several lines, a box under each row of "Board: Files" and "Board: Links", the last column of their sheets, `note` in
  the board file written only when filled; the "Resources" tab shows it under the name, its web addresses clickable and nothing read as formatting (`TextLinkRenderer.PlainLinkText` - it is typed in a plain box).
  A row with nothing but a note typed is a row (`BoardPartEdits.IsBlank`). The popup and the Overview list what their region shows,
  and the "Resources" tab what the region shown shows, following the switch (`Main.ShowResources`), the board's "Scope measurements"
  notes too (`ScopeMeasurementFiles.Of(board, region)`). "Board: Files" and "Board: Links", and a component's "Files" and "Links",
  list every row whatever region CRT shows - they are the whole lists - with "Applies to" as a column while the board has regions.
  **Any other word in that cell shows the row in neither region: a warning on the cell** (`row.region`, the table only). **A component in
  every region has rows of one region** - the usual shape - **so its rows are read in the region SHOWN, never its row's blank region**
  (code review, 2026-11-02).
- **A COMPONENT'S CATEGORY IS PICKED FROM A DROP-DOWN LIST, NEVER TYPED** (owner request, 2026-11-03: "Only in the 'Board: Category' we
  should allow for an input field, as this is the place we define a new category"): Basics (`CategoryBox`, its row's own spelling listed
  too, or showing it would change it - `CategoryChoicesWith`) and the new component's form (`NewCategoryBox`, saying where one is made when
  the board has none - `NoCategoriesToPick`); the label editor's prompt always was. **A file's, link's or credit's category stays a box to
  type in** - their lists are their own and nothing else defines theirs. The table's cells stay text too.
- **"Package" and a pin's "Direction" offer the usual words AND every other word the board's components use** (owner request, 2026-11-11:
  "Diskette" typed in the table is offered in the pane too; `ComponentVocabulary.PackagesOf`/`PinDirectionsOf` - after the usual ones,
  alphabetically, "Unknown" staying the last direction; `ComponentPane.ShowWordsOfBoard`, never under the package box's caret).
- **"BOARD: CATEGORIES"** (owner request, 2026-11-01; `ComponentPane.RegionsAndCategories.cs`, `BoardCategoryEdits`):
  rename, add, delete. **No list is stored** (owner decision: "kept until first used"): a category is the word its
  components carry, so a rename or delete is one step changing those rows - named as another it joins it, under its
  spelling; deleting asks where its components go. One no component uses is held by the pane per board until something
  does or CRT closes, and offered in the category boxes. Rows stay where they stand, so no order changes. **THE ORDER IS
  DRAGGED or Alt+Up/Down** (owner request, 2026-11-01: "alike Files or Links"; `BoardCategoryEdits.InOrder`): the order IS which
  is seen first in the rows, so it is one step moving those rows, said as the components' order; the components with none and
  one nothing uses have no handle. **Component
  categories only** - a file's, link's or credit's are their own lists'. **The filter beside the schematics follows**
  (code review, 2026-11-01): a category it has never seen is ticked, however it came - renamed, typed, edited in the table, a
  data update (`SavedCategoryFilter`; what the board had when it was chosen is kept, `SeenCategoriesByBoard`) - and one moved here
  keeps its components shown, even under a category the filter had unticked (`UserSettings.FollowCategoryMove`).
- **What Submit refuses, added 2026-11-01 (errors on both sides, server 6.0.0):** `component.no_category` ("No components
  must have an empty category" - the form's "Add component" needs one too; marked on the box, and on "Categories") and
  `pin.outside_package` (a baseline's or a pinout's pin past its package's count - `BoardDataChecks.PackageRefusing`, whose
  region rule is that a row for every region must be off EVERY region's case; nothing said for a case CRT cannot count, no
  package, or a pin named rather than numbered; "2 [64K]" is pin 2, while digits running straight into a letter are a name -
  a regulator's "12V", a 7400's "1A" - `ComponentVocabulary.PinNumberOf`). **Added 2026-11-02 (server 9.0.0):**
  `baseline.required_empty` - an oscilloscope baseline with no signal name or no image.
- **Directions: "Input" or "Output" for a pin carrying a voltage level as its signal too** (owner request, 2026-11-02: "Analog
  input" and "Analog output" went again after a day; a board holding one keeps it), **"Power" for the supply pin, "Unknown" last**
  - `ComponentVocabulary.PinDirections`.
- **A component opens in the COMPONENT PANE** (CRT 4.0 step 2's slice 3a, 2026-10-25;
  `ComponentPane`, `TabDrafts.Component.cs`; the plan: "A component opens in a pane, never in a
  maximized window"). **Its headline is "Component: " and the popup's title, "Component: U8 | CPU | 6510"** (owner requests, 2026-11-04;
  `ComponentPopupContent.TitleOf`, the one rule; the popup itself has no "Component:"). Basics (in the owner's order, 2026-11-04: board
  label, friendly name, technical name and short description down the left, category and part number on the right, every box one column
  wide - **"Purpose or note" under "Part number"** (owner request, 2026-11-09: "a new textfield (with formatting possibility) named "Purpose
  or note" ... why is it there etc." - `ComponentEntry.Purpose`, `[RichNoteMarkup]`, `purpose` in the board file written only when filled,
  a sheet's note column; a `RichNoteTextBox` with no height cap, each column of Basics a stack of its own so it grows down its own; **its bottom
  edge meets "Other technical details"'s in the left column** (owner request, 2026-11-09: "vertically align to the end-border of the "Other
  technical details""; a least height worked out after each layout pass, `ComponentPane.AlignPurposeWithDescription`, `BoxAlignment`); shown
  in the popup, `PurposeSection`) - and **the short description reads "Other technical details" wherever a person reads it** (owner request,
  2026-11-09; `BoardWorkbookSchema.DescriptionHeading`, the table's heading by `HeadingOf` - its stored names, `description` and the Excel
  header, stay); the short description with its counter to `ShortDescriptionLimit` and, past it, a warning in words right under the box, the counter
  beside it (owner reports, 2026-11-03: the counter's color was not seen, and 2026-11-04; `ComponentPaneWording.DescriptionTooLong`) - repair notes went on 2026-11-02, owner request: "no need for
  it", everywhere, as no released CRT ever read them). **"This data is different per region" tells the fields apart, never the board
  label** (owner report, 2026-11-04: typed "C1 PAL", then "C1 NTSC", and both read "C1 NTSC"): a label names ONE component, what every row of it is filed under,
  so renaming goes for every region and `LabelInEveryRegion` says so under the box while ticked - two labels are two components, one per
  region, by "Add component"'s "Applies to". Then Pinout (the pins,
  "Add the N pins of DIP-40" from `ComponentVocabulary.PinCountOf`, the case drawn from above -
  `PinoutDrawing` in `Controls/`, the popup draws it too; every position `PinoutGeometry`), oscilloscope baselines, pictures, replacement parts,
  files and links. **Every change is a
  transform of the draft held in memory** (`ComponentPaneEdits` - the component AS SHOWN, the region
  CRT shows by the component list's rule; the other region's rows and every other component's are
  carried across in place; a change that changes nothing returns the SAME board, so no step): a box
  when it is left, Enter pressed or typing paused a second, a row when its box is left or its list
  picked - each one named step (`DraftStepWording.ComponentField`/`ComponentPins`/...). **The pane
  saves the draft itself** a second later, off the UI thread, as the table does; the header says
  only "Could not save". **Another change lands on top of what the pane holds** (ChangeComing
  commits it first) and the pane reads the component again after any change but its OWN EDIT - an
  undo or redo carries the mark of the step it moves, the pane's own included (`IsOwnEdit`). A drop-down
  list's empty write is never a clear (`ComponentPanePinRow.Direction`). **A row's label in other
  capitals is no change** (code review, 2026-10-26: compared exactly, every commit rewrote a component's
  "u8" links, pins and replacements under "U8"); a real change keeps the spelling they use. Closing, another board,
  opening the table and leaving the tab hand over and save; submitting, managing files, quitting and
  installing an update save first and ask only when that fails (`SaveDraftEditorsNowAsync`,
  `SettleTableEditsAsync` - the prompt is said of the draft, not the table); a discard or a retirement
  lets go WITHOUT saving (`ComponentPane.Forget`). **Quitting saves EVERY draft held in memory with a
  change not saved** (`TabDrafts.OtherUnsavedDrafts`, the current drafts folder's) - a save the pane
  started as it closed can fail, and only an open pane or table was asked about. **Nothing reads the
  draft's file while the pane's close-save replaces it**: closing refreshes from memory and the row
  follows the save (`OnComponentPaneSaved`); opening the table awaits `ComponentPane.LastSave`, under
  the BusyOverlay. **Submit and "Manage files" save the draft held in memory itself too** - a change
  neither the table nor the pane shows (a close-save refused, the header's Undo with nothing open) was
  left out of the submission (code review, 2026-10-26; `SettleTableEditsAsync` - only that board's
  draft is saved, asked about or discarded). **The pane takes the session's `Changed` on the UI thread**
  (`OnSessionChanged` posts itself there, as the table does): a direct write (`WriteFileDirectly` - the
  KiCad data window's Remove) raises it on the pool, and the pane threw.
  **Oscilloscope baselines ("Scope baselines" on screen until 2026-11-02, owner request - the stored names stay: the "Scope
  baseline/" folder, the board file's `scopeBaselines`, the "Scope measurements" category), pictures and files are the pane's
  too** (slice 3b, 2026-10-26; `ComponentPane.Files.cs`, the rows `ComponentPaneEdits.BaselinesOf`/`PicturesOf`/`AllFilesOf` - one
  "Component images" list told apart by `ComponentImageKinds`): **a file added is copied in under a
  name of the board's own and the draft SAVED at once** (`ComponentPaneFileNames` - its kind's usual
  folder, never a name the board has, which it would replace for every row; a baseline named
  `<label>_<pin>_<signal>` as the published ones), so no undo leaves a file lying uncited. **A row is a
  baseline by its PIN alone** (`ComponentImageKinds`; owner request, 2026-11-02: "there is no requirement that one of T/DIV or
  V/DIV or Trigger level must be set" - measured that day, all 33,390 shipped rows with a pin have one too, so no published
  row changed kind) **and needs its pin, its signal name and an image** (`ComponentPaneWording.BaselineProblem`, and the error
  `baseline.required_empty` on the table and the server): a new one waits in its panel until it has them, and an edit that would
  take one away is held back - the draft keeps it as it was, the reason under it. **A new one's signal name comes with its pin**
  (owner request, 2026-11-02; `KiCadComponentPins.NameOfPin`, the pinout's, else the board's KiCad data's; `FollowNewBaselinePin`
  never over a name typed). T/DIV and V/DIV are `ComponentVocabulary`'s
  1-2-5 steps spelt as the data spells them ("50us", "1V"). **"Capture from my oscilloscope"** shows
  while the Oscilloscope tab has a session (`TabDrafts.UseScopeCapture`/`ShowScopeConnected`, from
  Main with the popups') and takes `TabOscilloscope.CaptureOscilloscopePngAsync` - no image folder
  needed; the capture waits for "Add oscilloscope image" in the user's OWN app data folder
  (`AppConfig.ScopeCaptureWaitFolder` - never the shared temporary one, where another account could
  swap it). Thumbnails decode small on the pool and are never disposed by hand (one still drawn takes the
  render thread down). **How the board's baselines were measured is its board files of the category
  "Scope measurements"** (owner decision, 2026-10-26: "Do what you recommend" - the live trees already
  keep it so, five boards with "Introduction to ... scope measurements" texts; NO board field, no wire
  or server change; `ScopeMeasurementFiles`): linked above a component's baselines in the pane and
  under a baseline in the popup, edited under Files like any board file.
  **The component as a whole is the pane's too; the component editor's window is GONE** (slice 3c,
  2026-10-26; `ComponentPane.Whole.cs` - the plan: "The 2,200-line component editor is replaced, not
  adapted"). **"Add component" opens the pane's form** (`TabDrafts.AddComponentAsync` ->
  `ComponentPane.ShowNew`): the board label, "Applies to" only on a board with regions - "All regions", "PAL", "NTSC", the list every
  row has (owner request, 2026-11-04; `AppliesToChoices`) - and the
  category; nothing goes into the draft until "Add component" there, then ONE step
  (`ComponentPaneEdits.WithNewComponent` - placed by `ComponentPlacement`; a label the board has in ANY
  region refused with the reason, `ContributionPackaging.ValidateNewComponent`; a category needed since 2026-11-01). **A component of one region opened while CRT shows the
  other is worked on in its own region** (`ShowComponent` - else every section read back empty and each
  commit filed its rows again). **"Delete component"** is `ComponentBoardWriter.ApplyComponentDelete` -
  every row, every region - as ONE step, and "Bring [U8] back" undoes it while it is the latest step.
  **"Show on board"** (enabled only for a component with a highlight, `IsMarkedOnAPicture`) saves the
  pane, then `Main.ShowComponentOnBoard`: the Schematics tab with ONLY that component selected
  (`TabSchematics.ShowOnlyComponent`) - when the pane's saves left the board on screen behind, it is
  read again first and selected once that load has its highlights (`ShowComponentAskedForAfterLoad`),
  the catch-up handed over so leaving the tab reads nothing again. **There a bar offers "Back to
  Contribute"** (owner request, 2026-10-29; a net lit up offers it too - `Main.OfferBackToContribute`,
  `TabSchematics.ShowBackToContribute`), **A `ReturnBar` - white, a 2px black dashed frame - AS THE PICTURE TOOL'S BAR IS:
  every bar taking a person back to the Contribute tab looks the same** (owner request, 2026-11-11: "This panel should be identical
  everywhere ... lightgrey background with 2px black dashed border", then white in place of the lightgray; `Controls/ReturnBar`, colors `ReturnBar_*` in neither theme -
  Khaki until then; a new such bar is a ReturnBar too): the pane is still open where it was. It goes once the Contribute
  tab is shown again, another board is shown (the same board read again keeps it), or a picture tool
  starts (whose own end goes back); a tool's bar takes its place while up. **"Preview component info popup"** opens the popup
  on the draft AS HELD IN MEMORY with the draft's folder whatever "view boards as officially published"
  says (`Main.OpenComponentPreview` -> `ShowComponentInfoPopup`, the popups' one path). **The popup and
  the Overview show the new data**: the package on the category line,
  the pinout (drawn, and every pin with name and direction BESIDE it, under it only when the popup is too narrow - of the
  region shown; owner request, 2026-11-01: **a pin of either lights up under the pointer and a click goes to its first image**,
  `PinoutDrawing.PinsCanBePicked`, `ComponentInfoWindow.GoToPin`, `ComponentPopupContent.FirstImageOfPin`; a drawing takes the
  pointer only where it draws, so it fills itself with Transparent; owner request, 2026-11-02: **the pin of the image shown stays
  lit on both**, `ShowViewedPin`, `PinoutDrawing.MarkPin` - `HoverPin` for the pointer in the list; owner request, 2026-11-09:
  **the pin under the pointer is lit on BOTH, whichever it is on** - the chip's case and a part drawn as it looks raise
  `PinHovered` for the POINTER alone, never for `HoverPin`, or the two would answer each other; the row's `IsPointedAt` - and
  **last, each pin's "Expected oscilloscope reading"**, its heading on three lines, from its baselines in the region shown,
  readings that differ a line each with what the pin says beside its number, "Pulsing [64K]" (`ComponentPopupContent.ReadingOf`);
  **every column as wide as its longest text, the same 20 after each** (owner request, 2026-11-09 - the space "Pin" had in its
  36 wide column; `PinTableGeometry`, `ComponentInfoWindow.PinList.cs`): Auto columns sharing their widths, a `PinCell` style's
  margin, and only when the room is short do Signal, Net name and the reading wrap - widest first, none below its widest word
  unless not even the words fit, when they break down to `PinTableGeometry.Narrowest` - to widths put in resources their MaxWidth reads (rows made again as a pin is pointed at take them too); the list stands beside
  the drawing only when it fits there whole), replacement parts,
  a baseline's machine state (`ComponentPopupContent`); the Overview has a Package column and
  its Notes are the first image's note (`OverviewHtmlBuilder.BuildRows`). **"Show pinout in its own window when the component info window would scroll"** (owner's choice of words, 2026-11-09 - "Detach pinout image in component info popup", then "... window", until then; right under "Open multiple component info windows" since)
  (owner request, 2026-11-07; `ComponentInfoWindow.Pinout.cs`) moves the popup's pinout section into a window of its own,
  870x960 the first time, its size and place one for all - **only when the popup has no room for it** (owner request,
  2026-11-09: "if no vertical scrolling will be implied with this info inside of component info popup, then you will not open
  a new window"; `PinoutFitsInPopup`): measured before a window opens, the popup's other sections, the pinout at the popup's
  width and the TALLEST image note and baseline section among the component's images (so the pins stay in sight on every
  image; measured apart, and of each only the `MeasuredAtMost` (3) wordings with the most lines and characters - forty
  baselines with notes of their own cost 135 ms a click, measured), against the info column's room; decided when a component
  or region is shown, the popup opens or the setting is ticked - never while images are gone through or the popup is
  resized. **A component shown is decided once Main is done with the popup** (`FollowPinoutDetachmentSoon`, posted at Normal,
  above Render; code review, 2026-11-09): Main calls `OfferEdit` and `UpdateOscilloscopeSessionTitleState` after
  `SetComponent`, and "Edit component" takes room from the title row - a test asserting on the window right after
  `SetComponent` runs `Dispatcher.UIThread.RunJobs(DispatcherPriority.Normal)` first. **A section is measured AFRESH**
  (`MeasuredAfresh`, every part invalidated): Avalonia keeps a control's measure while only something inside it changed, and
  a note section with a new note measured at the same width kept its old height. **The section is taken out at once and put in only after the window
  it left has had its layout** (`PlaceSectionAfterLayoutOf`; owner report, 2026-11-07: CRT ended on "Attempt to call
  InvalidateArrange on wrong LayoutManager") - Avalonia keeps a control queued for the old window's layout, and arranging it
  there once it stands in another throws. **A resistor's color bands are drawn
  IN THE PINOUT'S PLACE, read off its value** (owner requests, 2026-11-07, after the owner's own hand-made picture - under the pinout
  until 2026-11-09: "should be in the pinout instead (like the other components you are drawing, including the explanation of colors";
  `Controls/ResistorPicture` - the drawing, the value and the colors in words - the popup's `InfoResistorPicture` in the pinout, so it
  detaches with it, and the pane's `PinoutResistorPicture` on "Pinout", never for several; **UNDER the pin list, never beside it**
  (owner request, 2026-11-09: "I want the pinout table to be shown first and then the image and explanation afterwards" - the grids'
  second row, `ComponentInfoWindow.PlacePinList` putting the list across the first);
  the rule `ResistorColorCode`, every position `ResistorBandsGeometry`, the drawing `ResistorBandsDrawing` - a part lying on
  the page since 2026-11-08, owner request "more real": a soft shadow, round metal leads, the body and its bands shaded as a
  cylinder, the narrower middle over its own height):
  the values as the published data writes them ("560Ω", "4K7Ω", "1.5KΩ 1/4W 5%" - 116 spellings measured),
  **and a number alone, "100", "3.3k", "100 K", as KiCad data writes them** (owner request, 2026-11-07 - only a
  "Resistor" gets bands, so its number is ohms; nothing but a rating and tolerance may follow it),
  four bands, five for 2% or closer or a value two digits cannot say, a pill under each band, the value
  ("8.2K Ω 5%", owner request, 2026-11-07), its power rating and each band's color in words ("Colors:", a swatch
  each, the tolerance's last) under it. **A wrong picture is worse than none**: a resistor pack ("RP", "Resistor
  pack" - printed with a number code), a trimmer, an array or network, "wire", "N/A", a lower-case "m" and a
  tolerance no band stands for draw nothing. The colors are IEC 60062's, the same in both themes.
  **An axial ceramic capacitor is drawn the same way, its bands AS THE OWNER READ THEM ON HIS OWN 250407** (owner request,
  2026-11-09: "since that has color rings, would it then be possible to show and translate those also, just alike a
  resistor?"; asked: "Every axial ceramic", a band the data does not say "Show it as not known"; `CapacitorColorCode`, shown
  by `ResistorPicture`, `PartPicture.IsAxialCeramicCapacitor` - whose plain tube it then replaces): C23 "360pF 50V 10%" is
  orange | blue | brown | silver | silver, C36 "20pF 50V" red | black | black | gold | white - FIVE bands: two digits, the
  multiplier in pF and the tolerance in the RESISTOR'S colors (brown 1%, red 2%, gold 5%, silver 10%), then a fifth that
  differs from part to part (both are 50 V, so no voltage - most likely the temperature coefficient). **A band the data does
  not say is drawn as a dashed outline, `Band.Unknown`, never guessed**: the fifth always (`BandKind.Extra`), the tolerance
  when none is typed; a line under the colors says so (`NoteOf`). **Do not go back to the ceramic capacitor code** (a
  temperature band FIRST, green 5%, white 10%): the first version read the board's photo so, and the owner's parts proved
  it wrong - read digits first, the photo is 51 pF, the board's C38. Only what is certain: 10 pF or more, two digits, black
  x1 to yellow x10 nF, a typed tolerance with a color - else the plain tube with its printed code. A light-green body, as
  the photo's; the voltage where a resistor's power rating stands. 26 rows on the published boards (VIC-20 324003's eight
  100 nF "Ceramic Axial" among them).
  **The other leaded parts are drawn as they look, in the same style** (owner request, 2026-11-09: "could you also draw an axial
  and radial electrolytic capacitor in same style... and a disc capacitor ... And how about transistors? And diodes? Actually all
  of the current components for all boards"; the rule `PartPicture`, every position `PartPictureGeometry`, the drawing
  `Controls/PartPictureDrawing` - the popup's `InfoPartDrawing` in the case's place, and the pane's `PinoutPartDrawing`, never for
  several): radial electrolytics with their minus stripe and the PLUS leg the longer, axial ones with NO minus band but the groove pressed
  round the can by the plus end, its outline pinched (owner requests, 2026-11-09: "its "plus" leg is longer"; "I have only seen capacitors
  with an "indention ring" at the plus-side"), a disc FLAT and matte with its edge showing - never a shiny ball (owner report, 2026-11-09:
  "looks more like a round bubble") - dipped and axial capacitors, DO-35/DO-41 diodes with the
  cathode band, **their leads named "Cathode" and "Anode" in full, in English, whatever the data abbreviates them to** ("K", "KATODE";
  owner request, 2026-11-09 - a capacitor's typed names stay as typed), a TO-92 seen from its flat side and a little above, the D of its
  top over its face, its thin legs leaving the case together and bending out (owner request, 2026-11-09: "it does not look like a
  transistor"; `PartPictureGeometry.LegPath`), a TO-220 with each leg named - **a transistor's letters explained under it, "E = Emitter",
  "C = Collector", "B = Base"** (owner request, 2026-11-09; G, D, S for a field-effect one - `PartPicture.LegendOf`, only when every named
  lead is one of a set's letters), **a 78xx regulator by ROLE - INPUT, GROUND, OUTPUT from left to right, as every maker's stands, "IN = Input",
  "GND = Ground", "OUT = Output, +12 V" under it - numbered AS THE BOARD NUMBERS IT** (owner question, 2026-11-09: "is it always implied that
  1=Input, 2=Ground, 3=Output?"; `PartPicture.RegulatorOf`): the C128 310378's U59 is 1, 2, 3, but the C64 250407's and KU-14194HB's KiCad
  footprint numbers VR1 and VR2 1 = input, 2 = output, 3 = ground (the middle leg and the tab) - measured from the pads' places - so the pins'
  names say which is which, else the board's KiCad nets (ground's by name, the output's naming the voltage; `PartPicture.Of(..., nets)` - the
  popup's `_pinNets`, the pane's KiCad pins, drawn again as they arrive), else the datasheet's 1, 2, 3; what does not settle it draws nothing
  - crystals, ferrite beads and coils, EMI filters, resistor packs and pin headers. **A part lying across
  has leads only as long as their labels need** (owner request, 2026-11-09: "there is no reason to have such long legs, as it does take
  space from the right-side table/info" - `PartPictureGeometry.AxialLeadOut`; 250 wide whatever the body until then) -
  each lead with its pin number and name, clickable as the chip drawing's pins are. **The shape comes from the PACKAGE** (a capacitor
  with none only when its friendly name says the shape), **electrolytic from the friendly name and the value, never the description**,
  **polarity from the pins' names, else KiCad's convention** (a polarised capacitor's pin 1 is +, a diode's pin 1 the cathode -
  checked against the boards' KiCad nets, and again 2026-11-09 on the C128 310378: KiCad's DO-35/DO-41 footprints make pad 1 the square
  one, the cathode's, and every diode there to ground has pad 2 on it); a doubt, a contradiction or a pin with no lead draws nothing - a trimmer too (Open128's C20,
  "Trimmer" 4pF-40pF, is no disc). Measured on the converted boards: 750 of 2,906 components drawn; the rest have no package - setting one
  draws them.
  **The board's own lists open in the pane too** (slice 4a, 2026-10-26; `ComponentPane.Board.cs`,
  `TabDrafts.OpenBoardPartAsync`): the left column's "Files", "Links", "Credits" and "Important
  signals" show one of them in place of a component (`PaneBoardPart`) - Schematic images too since
  slice 4b; KiCad data opens a panel of its own there (below). Each list is ONE step when changed
  (`BoardPartEdits`, replaced as a whole); **a list is shown and written in the board's OWN order** -
  grouped by category until 2026-10-30 - so the board's list as the pane shows it is the SAME board,
  and opening and leaving a part changes nothing. **A file, a link, a credit or an important signal
  MOVES up or down past ANY row, whatever its category** (owner requests, 2026-10-29 and 2026-10-30:
  "the entire line moving up/down, and not only within its own category ... I also want to be able to
  rearrange the categories") - **DRAGGED by its handle**, Font Awesome's up-down (U+F338) at the row's
  right edge under the HAND cursor (the north/south one until 2026-11-01 - owner request: make it clear
  it can be grabbed), the row a dashed placeholder while it moves (owner request,
  2026-10-30: "instead of having the up/down arrows, then please have a 'North/South' icon and then drag
  the whole line up/down, just alike in the workbook for photos. It should show a placeholder when
  dragging"; `ListRowDrag`, the Maintainer tab's own; the Schematic images' cards keep Move up/Move down
  - "those panels are much larger"), or **Alt+Up / Alt+Down in its boxes** (`OnRowKeyDown`, the
  focus kept in the same box - never on a drop-down list, where those keys open it; `BoardPartEdits.CanMove`/`Moved`, `ComponentPane.MoveRow`). A drop or a
  move is the list's ONE step; the rows are replaced only after `ResetRowDrags`. **A FIELD A ROW NEEDS
  IS AN ERROR, MARKED ON ITS BOX** (owner requests, 2026-10-30: "'What' and 'Name or handle' must be set
  ... In 'Links' then all 3 fields must be filled ... Same for 'Files' ... the 'Important signal name'
  must not be empty"; and 2026-11-01: "highlight the problematic cell with a helper what is missing ...
  These errors must be flagged in 'Board status' and then it should not be possible to submit the draft
  board"): a file and a link all three, a credit "What" and "Name or handle", a signal its net and its
  name - **CRT.Data's `BoardRowRequired`, the ONE rule**: `BoardDataChecks` makes each empty one the
  error `row.required_empty` on its cell in BOTH scopes (the table, the draft's row, `StopForErrorsAsync`
  - and the server refuses it; a submission's credits and important signals are checked since, server
  5.0.0); Board status lists each list with such rows (`BoardGapKind.Incomplete*`, its button opening the
  list), the errors row counting the others. **The pane puts the row into the draft AS IT IS** and frames
  the empty box red with "Required before submitting" under it (`...Missing` on the row, the `Missing` and
  `Required` styles); a row with nothing typed in it is no row - the category "Add link" offers from the row above is no
  typing (`ComponentPaneBoardLinkRow.OfferedCategory`; code review, 2026-11-01). Until 2026-11-01 the pane held such a row
  back with a line under it - so nothing but the pane knew - and the table DROPPED an important signal
  missing a half on save (a warning, "left out when saving"); **no mapper drops a row now**
  (`BoardWorkbookSchemaTests`). `BoardPartRequired` only names the fields by the pane's headings. "Add
  file" fills every field - the last row's category, else "Resources" (`BoardPartRequired.NewFile`).
  **A LINK'S ADDRESS SUBMIT WOULD REFUSE IS MARKED THE SAME WAY** (code review, 2026-11-01:
  `BoardDataChecks.IsRefusedLink`, the rule of the error `link.not_web`, on a board's links and a
  component's, the helper `ComponentPaneWording.LinkNotWeb`) - "Open link" opens "www.lemon64.com", as
  `LinkTarget` adds the scheme, and nothing on the row said Submit would not take it.
  **ORDER IS A CHANGE** (owner decision,
  2026-10-30: "a change of (any) order should also be submitted as a change" - `BoardRowOrder`, below):
  **a board part's "Draft" mark is the differ's answer for that section OR its order**
  (`ComponentPane.ListDiffers`). **"Open file" on a file and "Open link" on a link** (owner requests,
  2026-10-29, and 2026-10-30 for the labels - "Open" until then): a file found as every file of the draft is (`ComponentPane.OpenFileOfBoard` - the "Scope
  measurements" notes' path too), a link by `LinkTarget`'s rule (the PDF export's - a scheme added when
  typed without one; mailto: as it stands, any other "scheme://" no link), both through `ExternalTargetLauncher`. **The part's ONE headline is the pane's
  title with its count** ("Files (3)", `ComponentPaneWording.PartTitle`; owner request, 2026-10-29: two
  headlines for the same thing) - a row not in the draft yet not counted (its `Committed` null). A board file added goes to "Board local files/" under a free name and
  is saved at once, as a component's. **An important signal's KiCad net is SHOWN AS TEXT, first, and
  never typed** (owner request, 2026-10-29: "they should not be editable ... switch the columns"): a new
  signal PICKS its net under the list from the board's nets no row has
  (`ImportantSignalSuggestions.NetsToAdd`; "Add signal" on only for one of them), and comes named as it
  would be suggested (`StartingNameOf`) - the name is the row's box. The nets are read on the pool
  (`TabDrafts.ReadNetChoicesAsync` - the files the Schematics tab draws, `KiCadProjectLoader`) and
  handed over once read (`UseNetChoices`); without KiCad data there is nothing to pick; a row naming a
  net the panel will not find (from the table, or older data) says so under it
  (`KiCadNetChoices.Finds` - `KiCadNetNames`' one rule). A heading row's button column shares its
  width with the rows' (`SharedSizeGroup`, in a `Grid.IsSharedSizeScope` panel) - an empty Auto column
  in the heading widened its star columns, and the labels stood off their boxes.
  **Pictures and their colours** (slice 4b, 2026-10-26; `ComponentPane.BoardPictures.cs`, the left
  column's "Schematic images"): a card per picture - a bordered panel of its own, its name box under
  "Schematic name" - with a preview of the picture and the components marked on it (`PicturePreview`,
  every position `PicturePreviewGeometry`), drawn in the colours as they are picked - a colour picker
  and opacity slider for "Schematic image" ("Selected components" until 2026-10-30) and "Thumbnail",
  and **a colour picker and opacity slider for "Traces on this side" and for "Traces on the other side"**
  (owner request, 2026-10-30 - before, this side's copper took the highlight colour and both sides its
  opacity + a quarter; that is still what a blank column draws) - and "Use for all schematic images".
  **The traces' colours are ONE rule, `TraceColours.Resolve`** (CRT.App's `Handlers/Geometry`), which the
  Schematics tab's copper (`TabSchematics.KiCad.Render.cs`) and the cards' previews both draw by - a schematic
  sheet with its own blank, `TraceColours.SheetBlankHighlight`, on its card too once the views say it is one; their
  columns are the board file's `traceHighlightColor`, `traceHighlightOpacity`,
  `oppositeTraceHighlightOpacity` (`BoardSchematicEntry`, written only when filled - server 5.0.0). **An opacity given
  holds for every trace, the one under the pointer too, and on a schematic sheet** (owner report, 2026-11-03: "even if I change
  the opacity down to 5% ... it uses a solid/100% opacity" - a hovered net was always 1.0, and a sheet went by the highlight's
  columns alone); it stands out under the pointer only while blank (`TraceColours.ThisSideStroke`/`OtherSideStroke`/`...ZoneFill`),
  as published boards keep. **A new board's image starts with Thumbnail at 80% and its traces in the Schematic image's color**
  (owner request, 2026-11-03; `PictureColours.ForNewBoard`, only with no published board: the highlight's color is written, so a
  blank trace color follows it - blank, it is DeepSkyBlue beside an IndianRed rectangle). **The previews stand on the card's RIGHT**
  (owner request, 2026-11-03; `ContributeWidths.PreviewDragged`: their handle dragged left widens them), **"Use colors for all
  schematic images" stands right under "Mark components manually"** (owner request, 2026-11-03), and **a color picker's list has a dropper** (below). **The previews** (owner's pick, 2026-10-29;
  `ComponentPane.PicturePreviews.cs`): the big one ZOOMS - Ctrl+scroll at the pointer (the wheel alone
  scrolls the cards), drag to move, double-click whole - and, ZOOMED only, asks for the picture decoded
  sharper (`SharperWanted`, `PicturePreviewGeometry.SharpWidth` steps, ONE card's sharp image held at a
  time - asked on every resize, widening the previews decoded every card; one decode per card at a
  time, the widest asked meanwhile next, a decode never shown disposed; asked again at the next zoom, or
  once the pointer RESTS on it (`PicturePreview.PointerRest` - asked as it passed, two zoomed cards took
  the one sharp image from each other), once another card took it). Shown whole, the pictures are read
  for the previews' width in `PicturePreviewGeometry.WholeWidth` steps - again only when a drag of the
  width ENDS past one, only while the pictures part is open, never for a card already being read as
  wide (`WantsPreviewFor`); never wider than the picture itself (`NoWiderThanThePicture`, the sharp
  reads' cap too); one read held per picture (the narrower let go of, `PaneImageCache.Forget`), and **a
  wider one held serves narrower previews with no read** (`TryGetWiderRead`; code review, 2026-11-01:
  every card made again decoded from disk once they were dragged narrow); a wider read that fails
  blanks nothing and may be asked again. Under it a thumbnail-sized
  one in the Thumbnail colours; "Traces preview" ("Traces on the preview" until 2026-10-30) has two
  ticks, "This side" and "Other side" (owner request, 2026-10-29: "so we can have both in preview") -
  this side's copper in its colour (blank: the Schematic image colour, its blank DeepSkyBlue, the tab's
  for copper - not IndianRed, the rectangles'), drawn OVER the other side's, the tab's order, each side
  at its own opacity (`PicturePreview.ThisSideTraceOpacity`/`OtherSideTraceOpacity`); both sides read ONCE
  (`TabDrafts.PreviewTracesAsync` - "Mark all automatically"'s plan, `KiCadPreviewTraces` placing them as
  the tab places copper; the cards asking at once share ONE read of the KiCad project,
  `ReadKiCadForPreviewsAsync`), the ticks kept by name AND file while the part is open - never put back
  on an image that can no longer show traces - a read answers only while it is the latest for its card
  (`TracesAsk`), and another KiCad view picked lets go of what was read, ticked or not (`ForgetTraces`).
  **An image with no KiCad data
  SHOWS NO KiCad controls** (owner requests, 2026-10-29 and 2026-10-30: "less disabled elements to
  relate to"): the traces' colors and opacities, "Traces preview", "Align KiCad" and "Mark all
  automatically from KiCad" are HIDDEN unless the board has KiCad files and the image a view picked
  (`HasKiCadData`); shown, the ticks and Mark all are greyed out until KiCad is aligned
  (`CanShowTraces`, `CanMarkAutomatically`), the reason in their tooltips. **A picture with a KiCad view, a side of the board
  or a schematic sheet, not aligned yet, says what comes next** in the amber
  notice under its KiCad view (owner report, 2026-11-01: "it is not clear that I must first do the
  'Align KiCad', before I can do the 'Mark all automatically from KiCad'"; `AlignsFirst`,
  `PictureToolWording.AlignFirst`). Which views are sides of the board the pane is handed with the views
  (`UseKiCadViews(..., boardSides)`, `BoardKiCadFacts.BoardViewsOf`), for a blank trace color alone (Orange on a sheet); null
  while not known. Highlights are
  drawn with the Schematics tab's edge rule (`SchematicHighlightsOverlay`: the edge 1.4 x the fill).
  The previews' width is dragged at any card's handle, the same on every card, kept per board as the
  left column's is (`ContributeWidths`) - **and never takes the settings beside them below
  `ContributeWidths.SettingsMinWidth`** (owner report, 2026-10-30: dragged to 0px): a drag stops there,
  and narrower cards show the previews narrower, the width chosen kept (`PreviewShown`,
  `ComponentPane.UseCardsWidth` from the cards' SizeChanged). **Under the thumbnail, once a picture shows a KiCad view, how many components that view
  shows and how many are marked on it** (owner request, 2026-11-01; `KiCadComponentTable.CountsByView` - a side by
  marking's own `IsOnSide`, a sheet by its parts - handed over with the views, `UseKiCadViews`; the marks of the
  region shown), in the default text color, not a hint's grey (owner request: "must be more visible"). **Told apart
  once any is by hand** (owner request: "If any components have been manually marked, then I think it should show
  like this"): **a mark is KiCad's while it lies exactly where its footprint or a unit of its symbol puts it under the picture's alignment**
  (`KiCadFootprintMarks.CountAutomatic`, `ViewFootprints`) - told from the rectangle, so NO board file changes; one
  moved by hand, or on a picture not aligned, is manual. **The card's
  buttons** (owner request, 2026-11-01: "the user most likely will NOT manually mark components when KiCad data is
  available"): "Align KiCad data" ("Align KiCad" until then) and "Mark all components automatically from KiCad" beside
  it, then "Mark components manually" ("Mark components") on a row of its own; "Delete schematic image" and "Use colors
  for all schematic images". **The Schematic images part takes the
  pane's whole width**; every other part and the component keep `ComponentPane.ReadableWidth` (980) -
  under that cap a wide preview was what squeezed the settings. ONE step for every picture's
  colours (`BoardPartEdits.WithPictureColours`, by name AND image file); written as the board keeps
  them, a colour by NAME when it has one (`RectGeometry.ColorText`), an opacity as "30%"
  (`PictureColours.OpacityText`). **A picker writing back the colour it shows writes nothing** (the
  row's setters compare with what is shown - a blank colour shows CRT's default, and the picker would
  otherwise write it). Adding, naming, matching, ordering and removing pictures are the cards' too
  since step 3 (below, "The pictures ... ON THEIR CARDS"). `BoardSchematicEntry.ReadColours`
  is a METHOD: the row's properties are its columns, counted by reflection. **The part marks the
  components of the region CRT shows** (`ShowBoardPart(..., region)`). **The pane reads each picture
  ONCE** - its thumbnails and the previews alike (`PaneImageCache`; code review, 2026-10-26): kept by
  WHERE THE BYTES LIE, their length and write time - never by name, which another board's file or bytes
  put back share - and a load from before the cache was let go keeps nothing (`Generation`). **A
  picture's own size comes from its header for every type CRT shows** (`ImagePixelSize`: PNG, JPEG,
  GIF, BMP, WebP - `FromBytes` for bytes in memory, the scope's screen picture; the ONE reader of picture
  headers, never a copy beside it); with it unknown, the preview draws no marks rather than wrong ones.
  **Contributor mode is GONE** (slice 6, 2026-10-26 - the plan: "Mark components and Align KiCad open
  the Schematics tab in that tool, with a bar and a Done button."). **Each picture card has "Mark components" and "Align KiCad"**
  (the latter only on a board whose draft lists "KiCad data/" files, for a picture with a CAD name):
  what the pane holds is saved, then `Main.StartPictureTool` opens the Schematics tab - on the board
  read again first when the pane's saves left it behind (`Main.OnBoardShownUpToDate`, the one
  after-load hook "Show on board" uses too, `DoWhatWasAskedAfterLoad`). **The tool WAITS for what it
  works on** (`TabSchematics.PictureTools.cs`; `PictureToolStart` decides): the picture's image read
  (`ImageSettled`, end of the selection change), and for aligning the board's KiCad data settled
  (`KiCadSettled` - at the end of the KiCad load, and from Main for a board with no KiCad files, which
  asks for no load); what can never come is REFUSED with the reason on the bar, never waited for. The
  bar FOLLOWS the label editor and the calibration (`FollowPictureTool`, called wherever either starts
  or ends): Done is their own save, Cancel their own cancel; a tool someone asked for reports its end
  (`PictureToolEnded`) and Main goes back to the Contribute tab; a board load forgets it silently.
  **The label editor's Done follows the bar BEFORE it reads the board again** (code review,
  2026-10-26: the reload's reset forgot who asked, and Done never went back). **A board read again
  WAITS while a tool is on** (`Main.WaitsForTool`, run when the tool goes off - `ToolOff`): the
  reload reset the tab and threw the label editor's unsaved rectangles away; a tool asked for while
  one is on shows the one that is on. **What waits for a board load belongs to the load that TOOK
  it** (`Main.TakeWhatWaitsForLoad`): one that never reached its end leaves nothing for a later
  reload to run out of the blue.
  **A right-click on a component turns its highlight on or off, and that is all** (owner request, 2026-11-05: "I
  do not want that, and I want its old behaviour" - a menu from CRT 4.0 until then). **"Edit component" is a button
  on its popup, for everyone** (`Main.EditComponentFromPopup` -> `EditComponentAsync` -> the pane, its draft made by
  `OnCurrentDraftAsync`): the popup goes, as it stands over CRT's window, and with a picture tool on that tool is
  shown instead. Not on the Contribute tab's own preview of the popup, nor on a popup still open once another board
  is shown (`ComponentInfoWindow.OfferEdit`, `FollowBoardShown`), **and it opens on "Basics"** (owner request,
  2026-11-07; `TabDrafts.EditComponentAsync`). **The log ALWAYS holds its debug lines** (owner request, 2026-11-07:
  "remove this checkbox and then always write this detailed logfile" - "Write a detailed log", contributor mode's
  heir, went; `Logger` has no switch). **What it also did, every board's data checked at launch into the log, went with it**
  (owner decision, 2026-11-07: its findings "should show just like any other error or warning in the "Contribute" tab"):
  `DataValidator` is gone, and **a published board with no draft shows its OWN errors and warnings** there, badges, list
  rows and "Overview" panels as a draft's (`DraftStatusReader.PublishedProblemsCached` - the very checks, on the published
  file, read off the UI thread for the board SHOWN, one published board's read kept at a time; `TabDrafts.ReadPublishedProblemsAsync`),
  **only once the tab is shown** (code review, 2026-11-08: every board shown on another tab was checked whole, for badges
  nothing else shows), the board status's panels waiting for that very check (`PublishedProblemsOf`).
  The STORED `contributorMode`, `validateDataOnLaunch` and `debugLogging` stay written as found: CRT 3.x reads the same
  settings file and goes by them.
  **"Clone component from another board"** (owner request, 2026-11-04: "select a hardware, a board and then a component ... clone all its data
  to this component. It will not change the component name, but all other data should be identical"; `ComponentPane.Copy.cs`, every rule CRT.Data's
  `ComponentClone`; "Copy from another board" until then, by part number with parts to tick): a panel over the page picks the hardware
  (the board shown's first), the board - any OTHER published board, read when picked (`TabDrafts.ReadCloneSourceAsync`, under the BusyOverlay; never
  a draft of it) - and the component, and says what that one is (`ComponentPaneWording.CloneFacts`). **Once used it opens on that hardware and board, the component
  empty** (owner request, 2026-11-04; `thisCloneRemembered`, for the run: one board's chips go across one after another). "Clone" is ONE step: every field, category and
  package included, pins, replacement parts, pictures, baselines, files and links, what the component held of those gone. **Its label and where it is
  marked stay** - a rectangle is a place on THIS board's pictures. **A file comes with its bytes** under a name no file here has, or an undo can bring back
  (`NamesHere`; the same bytes under its name, or another name in the same folder, are there already, so cloning again is no step), and **a row citing one that did not come is left out and said**.
  **Regions as far as this component has room** (`ReadingOf`): in every region on both, its rows as they are; of one region there, for every
  region here (the component must not vanish from the other); of one region here, what that region shows there, in it; on a board with no regions nothing
  says a region - a row that did would switch its PAL/NTSC buttons on.
  **"Also on other boards" IS GONE** (owner request, 2026-11-11: "for now we should remove the button 'Also on other
  boards' and its functionality, as I am not sure I want it done like this"): the component's button and the board files'
  one, their panels, `ComponentCopy`, `ChangedFiles` and `AlsoOnOtherBoardsOutcome` - all in git history (before
  2026-11-11) should it come back in another form. A change to the same part on another board is made in that board's draft.
  **"Fill pins from KiCad" ("Fill from KiCad" until 2026-11-03) and "net on this board"** (2026-10-26 - the plan: "On a KiCad board, Fill from
  KiCad adds each pin's net and its name from the symbol. The popup shows the signal name and 'net on
  this board' side by side"; `KiCadComponentPins`, read from the loaded data by `KiCadPinTable`): the
  FIRST PCB file's pads give the pin numbers and nets (`KiCadReferenceMatcher`'s rule), the schematic's
  symbol pins - every unit, from the sheet's own `lib_symbols` (`KiCadSchematicSymbol.Pins`) - name a pin
  before the pad's own `pinfunction`/`pintype` (KiCad 6+; all a board without a schematic has). Names
  in CRT's spelling ("~{RESET}" is "~RESET", "~" none; a pad number EVERY pad's name ends in is taken
  off - a real name may hold "_"), directions in `ComponentVocabulary`'s words, a supply pin told by its
  NAME too (the live 250469's library types every pin "input"). **The net is never stored** - it is the
  board's, read when shown: beside each pin in the pane (the host reads the pins whenever a component
  is shown - `ComponentPane.ComponentShown`, `TabDrafts.LoadKiCadPinsAsync`) and in the popup
  (`TabSchematics.KiCadPinNetsOf`, the KiCad data loaded for the board on screen). "Fill from KiCad"
  adds the pins missing and fills in EMPTY names and directions only - never what was typed - as ONE
  step. The pins are asked for again when the draft's KiCad files change while the component is shown
  (`ComponentPane.FollowKiCadFiles`, after every change and save), and a read answers only the ask it
  was started for (`KiCadPinsAsk` - an older read ending last never wins); an older file's pad on
  KiCad's "no net", (net 0 ""), is on no net (the loader marks it `IsNoNet` - its name reads "0").
  KiCad's escapes in a name ("~{RESET}", "{slash}", "~") are read by ONE rule,
  `KiCadNetNames.Readable` - the pins, the nets shown and the signal suggestions alike. The board's
  KiCad project is read for the pane by ONE helper (`TabDrafts.ReadBoardKiCadAsync`), and "Show on
  board", lighting a net and the picture tools hand the board on screen over by one
  (`SaveAndHandOverBoardOnScreenAsync`).
  **Important signals: suggestions and click-to-light-up** (2026-10-26 - the plan: "pick nets from a
  searchable KiCad list, with suggestions and click-to-light-up"; `ImportantSignalSuggestions`,
  `ComponentPane.SignalSuggestions.cs`): once the nets are read, every net whose WHOLE name says what it
  is - a supply, ground, a clock, a control line, a bus line - is suggested under the name the live
  boards use ("+5VDC", "9VAC", "Ground", "Clock; PHI2", "RESET", "A0"); a net a row has is never
  offered (`KiCadNetChoices.NormalizedOf` - the panel's rule). **A suggestion's name is a BOX**, changed
  before it is added (owner request, 2026-10-30: "show the suggested names in an input text field") -
  added under what it holds, the suggested name when it is emptied (`ComponentPaneSignalSuggestion
  .NameToAdd`). "Add" and "Add all suggestions" are one step each, taken after what the boxes hold is in
  (a net just typed is not added twice). **The list is changed in place** (`InPlaceList.Sync`, a
  suggestion kept by its net and suggested name - `IsTheSameSuggestion` - so a name typed survives):
  every box left commits the signals, and a list rebuilt then threw away the "Add" a press had gone down
  on. "Show on board" is off on a board without KiCad
  data - there is nothing to light. "Show on board" on a row or a suggestion is `Main.ShowNetOnBoard` ->
  `TabSchematics.ShowOnlyNet`: only that net lit (locked, as a click on its trace locks it; every
  component selected let go of and the highlights rebuilt - one selected only on the schematic kept
  its nets lit), on the
  board read again first when the pane's saves left it behind; while the board's KiCad data is still
  loading it WAITS for `KiCadSettled` (the load lets go of every locked net), and a board load drops it.
  **Every other net is listed last** (owner request, 2026-11-01: "all other nets not already used ...
  below the 'Suggested' ones, so all nets will be visible in here"; `ImportantSignalSuggestions.OtherNets`,
  `OtherNetsItemsControl`): each with the name it would be added under in a box (owner request, 2026-11-01: "a proposed net
  name in an input field"; `ImportantSignalSuggestions.StartingNameOf`), "Show on board" and "Add".
  A board has hundreds, so the list is a `VirtualizingStackPanel` and changed in place.
  **Credits know who is at the keyboard** (owner decisions, 2026-10-26: "Yes, let us try that", "Do it
  now"): a credit name and contact are REMEMBERED (`UserSettings.CreditName`/`CreditContact`, set
  together by `RememberCreditIdentity` - a blank name changes nothing; NEVER the contact email, since
  credits are shown to everybody). **The Credits part has no "Add me"** (owner decision, 2026-10-29:
  "Remove the button entirely - there is already the 'Add credit' button") - the name is remembered by
  the Submit dialog's credits below, and used there and by "Submit this model". **The Submit dialog offers a credit for the
  work sent, BEFORE sending** (`CreditSuggestions.For` - the draft against the published board, each
  kind of change its category in the board's own spelling; `SubmitDraftWindow.AddCreditsFirstAsync`):
  ticked ones go into the draft as one "Credits" step and are saved, and the board and fingerprint
  sent are read again (`TabDrafts.AddCreditsToDraftAsync`) - a credit added after sending would need
  a submission of its own. A draft that cannot take them sends nothing (`CreditsNotAdded`).

- **THE DRAFT'S ERRORS AND WARNINGS STAND WHERE THEY ARE** (owner request, 2026-11-04: "move the error and warning badges to where the
  problem is. E.g. if there are errors in 'Component: Basics' then there should be a badge there ... and then those errors and warnings should be
  clearly visible in there"; two counts in the header until then; `TabDrafts.Problems.cs`, every rule `ContributeProblems`). **Every problem has ONE
  entry of the left column**, by the sheet and row the checks report (`EntryOf`; "Component images" by what its row is, a highlight on the images, what
  no entry is about on "Overview", which says it in a panel of its own - a new sheet with no place fails `ContributeProblemsTests`). **"Overview" HAS THEM ALL, the entries below it their own**
  (owner request, 2026-11-04: "the 'Overview' entry should show all errors and warnings, and then you can see the details per entry in the list below"; `CountsShown`): its badge the
  draft's total, no panel saying so over its page (owner request, 2026-11-04: "The user can scroll down and see exactly where the issues are"), and **each of its panels has its own in a panel of their own right under its headline - how many, then each, a component's with "Open [U7]", and "Show in the table" - but on "Components", whose lines each open their component** (owner requests, 2026-11-04: "the error and warning panels should be shown specifically inside those panels and then detail the issues in there also"; "encapsulated in its own panel, as now the text is kind of mixed with the normal text"; `ShowStatusProblems`, `Border.StatusProblems`). Nothing is left out for being said another way there; the components marked nowhere are one sentence and a button each in that panel (`UnmarkedGroup`, not bold), not the same sentence eleven times, and **a component with a row per region is ONE warning, on its first row** (owner decision, 2026-11-04: "it is ONE physical component only"; `CheckHighlightCoverage`, a mark belonging to the label - "no category" stays per row, as each row has its own); a list's line in red counting rows left incomplete goes once the errors say them (`WithProblemsOf`). **A problem named by its row** (`ContributeProblems.TextOf`: "This credit has no Name or handle" says nothing in a list of three), **its region said whichever way, and "Open" shows that region** (`RegionToOpen`, all of CRT's, as the buttons under the list pick it; code review, 2026-11-04: a page lists the region shown, so a row of the other was counted on the entry and on no page - "Files" and "Links" list every row, so nothing switches for those). **WITH A COMPONENT SELECTED, A COMPONENT'S ENTRIES AND PAGES HAVE THAT ONE'S ALONE** (owner request, 2026-11-04: "I do not want to see any warnings or
  errors for other components"; `OnEntry`) - nothing wrong, no panel; with none selected they have every component's, each line with "Open [U7]" - that component, on that page. **EACH COMPONENT'S ROW IN THE LIST ON THE
  FAR LEFT SAYS ITS OWN** (owner request, 2026-11-04; `CountsByComponent`, `ComponentListItem.Problems`), changed on the row as it stands - a row made again lets go of the selection and the list scrolls back - in a column at
  the row's end, so a long name leaves them in sight; on no other tab. **"With issues" over that list narrows it to those with any** (owner requests, 2026-11-04: "only show components with either a warning or an error", and "I only want to see this checkbox when it at least has one component with an issue"; `ComponentsWithIssuesCheckBox`, there on this tab alone, the size of the "Labels visible" panel's, after its text, a press anywhere on the row; `thisBoardsOnlyWithIssues` per board for the run, "Clear" lets go of it): **what is being fixed stays listed until another is picked**, as the table keeps the row its cursor is on; one picked with nothing wrong goes as it is ticked, the pane staying on it; **the last issue gone takes the tick with it** - left ticked underneath, the next warning would narrow the list by itself. **Under the page's headline a panel says what they are** (owner request, 2026-11-04: "should come below the headline everywhere"; `EntryProblemsPanel`, moved to `ComponentPane.ProblemsHost` or `NoComponentProblemsHost`, under a component's buttons so those stay put; not over a table, which marks its cells and has its pills). **Read with the drafts' list** (`DraftStatusReader.ProblemsCached`, the very same object while nothing moved), so it follows a save, a second after a
  change; what a page marks at once stays its own. **The header says no count of a new board** (owner request, 2026-11-04: "New board, 3503 rows so far - I do not think we need this text at all";
  `HeaderSummary`, nothing for one); "My drafts" and the Submit dialog say "New board", and a published board's "N rows changed" stays.
- **What belongs to the component shown goes as another is shown** (code review, 2026-11-02;
  `LetGoOfWhatIsOfTheComponentShown`): the new baseline being filled in, a capture with it, the pane's
  message line; the same component read again keeps them. **A pane edit keeps each row in its own place**
  (`ComponentPaneEdits.ReplaceInScope`): kept rows stay, freed places go to new rows, and a real reorder
  takes the places in turn - shipped chips alternate PAL and NTSC rows. **Enter on the new-component form
  adds only from its label and category boxes, never with a drop-down list open** (`AddOnEnter`). **An undo
  or a redo of a rename is followed by what that very step renamed, never by a record of renames**
  (`FollowRelabelInDraft`, `ComponentRelabel.Between`; code review, 2026-11-02): which component a label means
  depends on where in its history the draft stands, and no record could tell a new U1 from the old U1 undone
  back. **A pane save that fails stops "Show on board", lighting a net and the picture tools, and the pane says so**
  (`SaveAndHandOverBoardOnScreenAsync`). **Opening the pane or the table reads the draft off the UI thread,
  under the BusyOverlay** (`ReadDraftToOpenAsync`). **A draft made only to look whose action opened nothing
  goes as that action ends** (`OnCurrentDraftAsync`); an earlier one left behind goes as the next look
  starts, and as CRT closes whichever board is shown. **One holding nothing but a step to redo stays
  while CRT runs, and goes as it closes** (owner request, 2026-11-02; `thisLooksKeptForRedo`): the Redo can
  still bring the change back, and it ends with the run.
- **No action asks whether to make a draft - it is made** (`Main.EnsureDraftAsync`, `Main.FirstDraft.cs`;
  "Edit board as draft" and `DraftExistsWindow` are gone): seeded from the published FILE
  (`DraftSeeder.SeedFromPublishedFile`) under the BusyOverlay, after the API revision is asked
  (`AskApiRevisionBeforeFirstDraftAsync`); a covered tab, or a seed that fails, makes nothing and says
  why on the tab's problem line (`TabDrafts.ShowDraftProblem`). A draft made meanwhile counts as made.
  **A draft made only to OPEN something - a component ("Edit component" is on everybody's
  popup), a board part, the table - goes again when that closes with nothing in it**
  (`TabDrafts.LetGoOfDraftMadeOnlyToLook`; code review, 2026-10-26): nothing waiting to be saved, no
  change against the published board, never a new board's, never one there before - **nor one holding
  its rows in another ORDER** (`BoardRowOrder`; code review 2026-10-28: pictures moved and the draft went
  with the new order). **Until its first change it is not shown as a draft at all** (owner request, 2026-11-01: "It should
  not show it as draft until I have done my first edit, whatever that might be"; `TabDrafts.ShowsAsDraft`, `IsUntouchedLook`):
  no row, no "Discard draft", no red tab title, the no-draft line in the header - though its folder exists and what was
  opened stays open on it (`ApplyTableMode`). Its session's `Changed` shows it at once; a write the session does not see is
  found as the list is read again. Quitting lets it go (`LetGoOfUntouchedDraftOnQuit`), and so does installing an update, once
  the download is in (`Main.BeforeUpdateRestartAsync` - the restart exits without a Closing event; code review, 2026-11-01), closing
  what is open on it, as that restart can still fail (`LetGoOfUntouchedDraftBeforeRestart`).
- **ORDER IS A CHANGE** (owner decision, 2026-10-30: "a change of (any) order should also be submitted as
  a change"; CRT.Data's `BoardRowOrder` - the lists of `ReviewSummary.AllSections` but the component
  highlights, whose order nobody sees; only rows BOTH boards have are compared). It is counted BESIDE
  `BoardDataDiffer`, which still counts rows (a moved row is none added, changed or removed - the table
  colours none). **Component images are compared as a board file holds them, the pictures first and the
  scope baselines after them, each kind among its own** (`AsAFileHolds`; code review, 2026-11-01: a
  baseline dragged among the pictures was said to be an order the save then lost); and **"Undo changes
  to [U8]" puts each row back where it was** (`ComponentRevert.Restore`/`RestoreImages`: a component's
  rows that do not stand together went back together, an order change nothing could take back).
  **A component's OWN rows in another order are a change of it** (`BoardRowOrder.ComponentsReordered`;
  code review, 2026-11-02: its links dragged read as unchanged): "Draft" in both lists (`ComponentRevert
  .HasChanges`, `BoardDraftSummary.FromChanges`), a line of its own in "Review changes" and an "Undo changes
  to" that takes it back. **One only moved past another is none** - a category dragged would mark every chip
  of it - and is the list's line under "The board". It is
  asked by: the draft's row and Submit (`DraftOtherChanges.Reordered`, against the published board
  as rows are counted), "Review changes" and the Submit dialog ("Order of the files changed",
  `DraftChangeSummary`), the board parts' "Draft" mark, retirement (a draft in another order is NOT
  retired), and the maintainer's side - the server's summary (`ReviewChangeSummary.ReorderedSections`,
  the wire's `reorderedSections`), a line above the submission's table (`ReviewNotInTable`) and the
  board's History (`SubmissionChanges.Reordered`). The server publishes a submission's rows as sent, so
  the order reaches the published board with it. **"Bring my draft up to date" keeps a section the
  contributor reordered in their order** (`DraftRebase`, code review 2026-10-30 - it took the update's
  and the reorder was lost); a section left in the base's order follows the update's.
- **"Board name" is the left column's first entry and a page of its own** (owner request, 2026-11-01: "how
  to edit those details again afterwards - I think this should be the very first item";
  `TabDrafts.BoardName.cs`, `BoardNamePanel`). **The fields and every rule about them are ONE control,
  `BoardNameFields`, shared with "Add a new board"**. **A new board's names and notes change there until
  it is sent for review.** Its names ARE its identity and its folder ("<Manufacturer>/<Hardware>/<Board>/
  board.json"), so a change moves the draft (`DraftBoardRename`: the folder whole, then the marker and the
  board file; never over anything, all or nothing, capitals alone through a name in between, emptied parents
  removed; a file held for a moment is waited for, as `FileReplacer` waits - Windows will not move a folder with
  a file open in it, nor replace a file something has open, and CRT's own reads or a virus scanner hold one so;
  the move and `DraftMarkerStore.Save`, 2026-11-02). What the draft holds is settled first; its session closes only once the folder has moved, or its
  names changed - its Undo steps go then, which the page says, and never for notes alone or a rename refused (code review, 2026-11-01). **`Main.AfterBoardRenamed` carries over what this machine files under "{hardware}|{board}"**
  (`BoardKeyMove`: settings - `UserSettings.MoveBoard` finds every "...ByBoard" dictionary BY THAT NAME, so a
  new one follows by itself and `BoardKeyMoveTests` proves it - traces, workbooks, devices) and shows the board
  by the steps a board just made takes. **A published board's name is only shown**, why said, and no
  draft is made (owner decision, 2026-11-01: "a published board should not (for now) be renamable" - the
  server and every CRT find it by the name, so renaming one would need the server). **Sent for review, a new board keeps its name while a maintainer has it**
  (under a new name a submission would be a second board). **Only a decision against it frees the name**
  (`SubmissionReceiptPresenter.IsDecidedAgainst`; code review, 2026-11-01): accepted and published are final too, and the
  board then reads as published (`IsInPublishedData`) on a machine that has not synced it yet. Never while a picture tool is on.
- **"Board status" is a panel per part of the board, new or published** (owner request, 2026-11-01: "each
  point in its own panel with a headline, then an explanation what this is about, a button for it and
  lastly also a status where it will count the data for it, if any ... some coloring and the checkmark icon
  or a 'missing to do' icon"; it replaced the new board's checklist and "What's missing";
  `TabDrafts.Readiness.cs`, every rule and word `BoardStatusPage`). Board name, Regions, KiCad data, Important KiCad signals,
  Schematic images, Categories, Files, Links, Credits - and Components last (owner request, 2026-11-03; the enum's order is the page's). **Every error and warning is said on the panel it is
  about** (owner request, 2026-11-01: "it should have the errors and warnings directly per section"; `BoardStatusPage.PartOf`
  by its sheet, a highlight's the images'; the checks `DraftStatusReader.ProblemsOf`, on the draft held in memory, off the
  UI thread): the first `ListedAtMost` in full, the rest counted, and "Show in the table" on its sheet (none on "Components"). A warning leaves no
  panel done or optional, an error makes it one. They stand in a panel of their own under its headline, counted (`ShowStatusProblems`). Only what is about no part keeps "Errors and
  warnings", first, only while there is any. **A panel is Done, ToDo, Error or Optional** (`BoardStatusState`
  - by what the board holds, never a box ticked), its frame and icon saying so (`Border.StatusPanel`; the
  icons differ in shape too), a line off in its tone. **The KiCad project is optional**: without it its panel
  and the signals' are Optional, nothing held against the board. **The Schematic images panel counts in the
  owner's words** (owner decision, 2026-11-01: "25 images known by KiCad data", "4 images added", "2 KiCad files
  mapped", "1 KiCad file aligned" - a picture card says "Associate KiCad file"): what is counted is the views, so a
  PCB file - ALWAYS read as two, its top and its bottom, `KiCadRawProjectLoader.BuildProjectViews` - is 2 images
  and, both sides mapped, "2 KiCad files mapped". **"Net" is one connection**, a netlist the list of them (the
  Schematics tab's "Netlist names" panel became "KiCad nets" that day). **"Can be marked automatically" is a footprint or a
  schematic symbol to mark from** (`BoardKiCadFacts.Markable`, each component once). **"Not on the board yet" is the KiCad data's components the board has none
  of** (owner request, 2026-11-01 - on the KiCad panel too, the very list "KiCad data" shows, `KiCadReferenceMatcher.Build`'s
  `componentReferences`; "not labelled yet" until then, which is NOT "not marked": a component can be on the board and marked
  nowhere). **Under Schematic images the suggested matches, under Components what
  each lacks, a button per component** (`BoardReadiness.Missing`'s groups, the first `ListedAtMost` named).
  A row with a field it needs empty (`BoardRowRequired`) is an Error on its list's panel. **Worked out ONCE per board shown and when the page comes back** (`ShowStartPanel`;
  ShowBoard skips its own when that did it), never on every `RefreshDrafts`, and **not while the tab is off
  screen** - then when it is shown (`IsOnScreen`, `ShowReadinessIfStale` from `OnAttachedToVisualTree`; code
  review, 2026-10-26); from the SESSION's content when the board has one (the pane's saves leave the board
  shown behind). **Working it out changes nothing in the draft** (`ShowStatusOfAsync`; code review, 2026-11-02 - see "A NEW BOARD'S
  CREDITS"). The KiCad facts arrive a moment later through `ReadBoardKiCadAsync` **into the same panels**:
  every panel and button is made once and changed in place (`CreateStatusRows`, `InPlaceList.Sync`) - a control
  made again under a press threw the click away; a newer ask wins (`thisReadinessAsk`). A board's last facts
  and KiCad files are KEPT and shown while they are read again; "not known yet" is said, no count guessed.
  **KiCad files that could not be read are NOT empty KiCad data** (`BoardKiCadFacts.Read`): an Error on the KiCad
  panel. A new board whose draft cannot be opened shows no panels.
- **The KiCad components added to the board** (step 3's slice 3b, 2026-10-26 - the plan: "one press
  adds every KiCad component that isn't on the board yet"; the rule CRT.Data's `KiCadBoardComponents`,
  the KiCad types handed over by `KiCadComponentTable`, the one implementation
  `TabDrafts.AddFromKiCadToAsync`). **Since 2026-11-01 THE KiCad IMPORT ADDS THEM** (owner request: "I
  instead do think it should add all components when importing the KiCad data ... Have checkbox if it
  should add all data automatically also?"): while "KiCad data"'s **"Add the KiCad components to the
  board"** is ticked - ticked until unticked, remembered (`UserSettings.KiCadImportAddsComponents`,
  handed in as `TabDrafts.KiCadImportAddsComponents`) - an import runs it as part of its queued change
  (`KiCadDataPanel.AddKiCadComponents`, the panel catching up as the change ends, the report's labels
  following); the summary says what came of it and offers the rest, **"Add N components from KiCad"**
  (`KiCadImportSummary.AddComponentsButton`). Board status's "Add N from KiCad" (the components step's
  button, and its gap) went. **The components are the FIRST PCB file's footprints, once per reference** (a project
  with no PCB, its symbols) - never a placeholder ("REF**", "#PWR", "U?"), a footprint with no numbered
  pad (a logo), a mounting hole (what a footprint IS, by its library or own name - owner report, 2026-11-01: a D-sub's
  "..._MountingHolesOffset8.20mm" is a connector, and the Open128 lost four), a fiducial or a net tie; Board status counts the
  very same ones (`BoardKiCadFacts`). **The category is what the live boards use for the reference
  prefix** (measured on the stable tree: C Capacitor, R/RP Resistor, U*/IC IC, CR/D/DP Diode, CN/P
  Connector, Q/TR Transistor; the rest Misc). **J is a jumper, Misc** (owner decision, 2026-10-31: "the
  'J' component should be a 'Misc' category, as it is a jumper" - 30 of the live boards' 32 J parts, and
  every J of the shipped KiCad projects; KiCad's own libraries' J connector is changed by hand). **The package only when the footprint's NAME says it and its pads agree** (a name STARTING
  "DIP16", "PDIP-40", "DIP-14_W7.62mm" with that many pads - "SOIC-DIP-8" and "Oscillator_DIP-14" are
  none; a 24/28-pin DIP "wide" by W15.24 in its name or pin 1 and the last pin 600 mil apart; SIP and a resistor array; TO-92
  and TO-220 with three pads; **since 2026-11-03** (owner request: "what package is a resistor? Or a ferrite bead or a capacitor?" -
  read off the six published boards' footprints, `ComponentVocabulary.Packages` growing with them; 1,206 of their 1,506 footprints get one):
  DO-35, DO-41, HC-49, HC-18, a pin header, a three-legged "EMI filter", and on two pads Axial, Radial or Disc as the name says - a resistor,
  ferrite bead, inductor or diode by its name, and a resistor whatever its footprint is called; "SMD" or "0805" in a name, or "SMD" in its library, says otherwise, a resistor array too) - else
  empty: **"CAP_P5.08mm" says only its pitch** (KU-14194HB's 125 capacitors). The value is KiCad's ("~" none). A
  component with a package gets its PINS (`KiCadPinTable` - "Fill pins from KiCad"'s; the net never stored),
  one with a datasheet that is a web address (the symbol's, else the footprint's; `TextLinkFinder`'s
  shapes) a "Datasheet for <value>" link. **ONE step** (`DraftStepWording.ComponentsFromKiCad`), placed
  by `ComponentPlacement` - **a category the board does not have goes where the alphabet puts it** (owner request,
  2026-11-03: chips first until then; `KiCadBoardComponents.IndexFor`, the board's own rows staying where they are) - saved at once under the BusyOverlay - **counted
  and named on the draft as the change finds it** (`DraftEdits.OpenForChangeAsync` first: a file
  written outside read, what editors hold handed over); only a change reads the board on screen again.
  What came of it is said over Board status's panels (`BoardReadiness.AddedFromKiCad`) **only while the draft
  is the content it was said of** (`ReadinessNote` - an Undo takes it); no draft to open is said on the
  tab's problem line. The
  loader reads each footprint's name, value and datasheet (`KiCadPcbFootprint.FootprintName`/`Value`/
  `Datasheet`). On the shipped KiCad data every chip gets its package but 250469's PLA (its own footprint).
- **"Mark all automatically from KiCad"** ("from KiCad" on the button since 2026-10-30 - owner request:
  what it marks from was not clear; `PictureToolWording.MarkAutomaticallyButton`; **a picture with marks already asks first**,
  saying they stay and how many get one - owner request, 2026-11-01, `MarkOverMarksWindow`, `TabDrafts.AskBeforeMarkingAsync`;
  nothing marked or nothing left to mark asks nothing, nor does the marking after aligning) (step 3's slice 3c,
  2026-10-26 - the plan: "on aligned PCB pictures,
  Mark all automatically draws each component's rectangle from its footprint"; `Handlers/Geometry/
  KiCadFootprintMarks`, `TabDrafts.MarkAutomatically.cs`): a green button on each picture card in the
  Pictures part, on for a picture with a CAD name AND a calibration in the draft (`KiCadFootprintMarks.Plan` -
  the first thing missing is said under the part's heading). **A SCHEMATIC SHEET IS MARKED TOO, FROM ITS SYMBOLS** (owner
  decision, 2026-11-04: "Yes, go ahead with all 3" - passed over until then, as having nothing to mark from; `KiCadSchematicSymbol.Extent`,
  `KiCadFootprintMarks.ShapesOf`): a rectangle per UNIT drawn on it - what the unit draws with its pin stubs, not its reference or value text,
  hidden pins left out - so "components" in what is said counts each once (`ComponentsIn`), and a chip marked in one unit counts as marked. A
  point of a library symbol lands by `KiCadSymbolPlacement`: **turned first, flipped on the sheet after, an order MEASURED** (pin tips on wires on
  every shipped project, 100% and 94% against 29% and 53%). **Never a footprint's or a symbol's rule apart: both go through
  `KiCadViewFootprints`** (`ByLabel`, a list per label). The card is
  its name AND its image file (`KiCadFootprintMarks.PictureRow` - of two pictures of one name, the one
  pressed; code review, 2026-10-30), as for its traces; the Schematics tab's picture tools still go by
  name, as its highlights and alignments are filed. **The board's
  world bounds and the mapping onto the picture are the Schematics tab's own** - its
  `GetKiCadPcbWorldBounds`, `MapKiCadWorldToImagePixel` and `MapKiCadWorldLengthToLocal` call
  `KiCadFootprintMarks.PcbWorldBounds`/`WorldToImagePixel`/`WorldLengthOnPicture` (the last the card
  previews' trace width too) - so a rectangle lands where the copper is drawn; never a second copy of any.
  A rectangle is the footprint's pads (each turned by its own angle) with 0.5 mm around and at least 2 mm
  across, in whole pixels cut to the picture (a corner a hair off a pixel IS that pixel - float noise
  added one). Only components the board HAS, seen from the picture's side (a through-hole pad from both,
  a surface part from its own), with NO mark on the picture yet - a mark drawn by hand is never moved;
  ONE step (`DraftStepWording.ComponentsMarkedFromKiCad`), worked out on the draft as the change finds
  it, saved at once. **Never while a picture tool is on** (`Main.IsPictureToolOn`; code review,
  2026-10-26: the label editor's Done writes its picture's rectangles as it holds them, and the marks
  went without a word). Nothing marked is said for what it is (`KiCadMarkPlan.WhyNothing`: no components
  yet, no footprint on the picture, all marked already), and the note under the part goes once the draft
  moves on (`DropPicturesNoteIfMovedOn` - an Undo takes it).
  **A KEY OR A DRAG WHILE ALIGNING REBUILDS NOTHING** (owner report, 2026-11-06: "incredible slow ... not always react when I use the
  keys"; measured 75-265 ms per arrow key, 37-60 ms of raster per drag frame): the box moves by `OverlayTransform` alone, a board of more
  than 500 primitives drawn from one picture of the overlay while dragged or a key is held (`KiCadOverlayRenderControl.DrawFromSnapshot`,
  every position `OverlaySnapshotGeometry`) - **taken again once the view needs more than it holds** (`StillCovers`; code review,
  2026-11-07: kept across moves, a band of traces went missing past its margin) - rebuilt once 400 ms after a resize or flip settles; the box itself is drawn untransformed. The
  runtime cache's scope key is the one computed at load - it stat-ed every KiCad file once per net. **The arrows (and Escape) reach the box
  from anywhere in CRT's window** while the tool is on and the tab shown (`Main.KeyIsForKiCadAlignment`, `TabSchematics.TakeKiCadCalibrationKey`)
  - never from a text box, nor an open drop-down (code review, 2026-11-08: Escape there ended the alignment, the list still
  open). A held key's repeats were lost behind the rebuilds: Avalonia 12.1.4 ignores Windows' repeat count.
  **DrawImage's SOURCE rectangle is in the image's PIXELS, never its DIPs** (owner report, 2026-11-06, a display at 125%: the traces
  "zooming slightly" as a marker was grabbed; `OverlaySnapshotGeometry.SourceRectOf`): Skia takes it as it is
  (`WriteableBitmapImpl.Draw` -> `Canvas.DrawImage`), and the picture, taken at the screen's scaling, is 80% as many DIPs as pixels
  there - so `new Rect(bitmap.Size)` drew 80% of it over the whole area. Headless tests run at 100%, where the two are one number.
  **WHILE ALIGNING ONLY THE SIDE BEING ALIGNED IS DRAWN**, whatever "Show opposite side traces" says (owner request, 2026-11-03: "if I am
  aligning the top PCB, then I only want to see traces from the top"; `RenderKiCadPcbGeometry`) - a through-hole pad is on both sides, so it stays.
  **AND NO COPPER POUR, whatever "Show zones" says, which is not offered then** (owner request, 2026-11-05: "NOT include any ground plane, as it is hard to
  align those nets"): every zone, by what is drawn and never a net's name - measured, the pours are on "GND" or "Earth" and on "+5V" too. Each net's tracks and pads stay.
  **NOR ANY WORKLOG, whatever "Show worklogs" says** (owner request, 2026-11-07: "When I am aligning KiCad data, then I do not want to see any
  worklogs on top"; `TabSchematics.FollowAligningInWorklogOverlay`, from `FollowPictureTool` as the calibration starts and ends): the overlay
  is hidden, never cleared, so it comes back as it was.
  **A SHEET BEING ALIGNED SHOWS EVERY WIRE, NAMED OR NOT** (owner report, 2026-11-04: "it has no traces in the alignment that I can align
  to" - the Open128's "Power", 221 wires and none drawn; `everyWire`). Half-drawn sheets were aligned off by up to 11% and said nothing.
  **A POWER SYMBOL IS A NET LABEL** (`KiCadSchematicLabels.Power`, `ExtractSchematicPowerLabels`: flagged "(power)" with a power_in pin, so never PWR_FLAG,
  whose pin is power_out; its Value, and its pin's name where that differs - 250469's "9V_UNREG" has a pin "+9V_UNREG", what its PCB calls the net): the
  Open128's wires on a net went from 1,556 to 2,447 of 3,083. **`SheetWorldBounds` MUST LEAVE THE POWER LABELS AND THE SYMBOLS OUT**: every alignment
  ever saved is relative to it, and growing it would move them all. `KiCadSheetNetAliases` does not take them either.
  **"Align KiCad"'s Done MARKS THE PICTURE BY ITSELF, without asking** (owner request, 2026-11-01 -
  "OK, make that"; `TabDrafts.MarkAfterAligningAsync`): every picture aligned with Done (or "Done - next
  picture") on a tool the Contribute tab asked for is told (`TabSchematics.AlignedFromWorkspace`, only
  after the alignment SAVED) and kept by Main (`thisAlignedToMark`, cleared as a tool starts); once the
  tool hands back to the Contribute tab, each of the board still shown is marked as "Mark all" marks
  (`Main.MarkPicturesAlignedAsync`) - but on its first alignment a picture with ANY mark already is left alone
  (its marks were drawn; another round is asked for with the button - `KiCadFootprintMarks.HasMarks`), and what came of it says the alignment
  is in too (`PictureToolWording.MarkedAfterAligning`). The board shown is asked before EACH picture
  (code review, 2026-11-01: each is marked on whatever board is shown then). ONE step per picture, which
  Undo takes back. "Mark all" stays for later rounds. **ALIGNED AGAIN, THE MARKS KiCad DREW MOVE WITH IT** (owner request,
  2026-11-01: "if I then again do the Align KiCad then it seems it will not readjust the components, which I think it
  should"; `KiCadFootprintMarks.Realigned`, the alignment before handed on by `AlignedFromWorkspace`): a mark exactly where
  its footprint or unit lay under the alignment before goes where it lies now (each mark answers for ONE place: two units cut to the same pixels
  part), one now off the picture goes, a component newly on it
  gets one (only newly: one that lay on it all along with no mark had it deleted on purpose, and it stays deleted); a mark
  anywhere else was moved or drawn by hand, and stays. No mark of KiCad's on the picture: left alone. ONE step.
- **The Delete key deletes every rectangle selected in the label editor**, as ONE step of its undo (owner request, 2026-11-03;
  `TryDeleteSelectedLabelEditorHighlights`) - not under a modifier, a drag or the new-label prompt, and not with the
  focus in the component filter, outside the tab, where the key is the box's (`MainPictureToolsTests`).
- **The label editor lists the components not marked on its picture** (slice 3c - the plan: "Photos and
  scanned schematics use the label editor as today, with the unmarked components listed beside it";
  `TabSchematics.LabelEditor.Unmarked.cs`, `UnmarkedComponents`): IN the picture tool's bar (the editor
  only starts from "Mark components"; over the picture's corners it covered Global settings), one button
  per label, changed in place. **One picked names the next rectangle drawn** - label and category, no
  prompt (`AddDrawnLabelEditorRectangle`); picked again it is let go; an undo puts a component back.
  **Done asks a category only of a label that makes a NEW component** (`LabelEditorBoardWriter
  .IsNewComponent` - the save's own rule, in the region it saves in; code review, 2026-10-26): it asked
  every rectangle, and one of a component with no category - picked, or marked automatically - could not
  be saved at all.
- **Suggested matches** (step 3's slice 3d, 2026-10-28 - the plan: "Match: suggested pairs ("...Top.png"
  with "PCB Top") to confirm"; the rule CRT.Data's `SchematicCadNameSuggestions`, the views as
  `BoardKiCadFacts.ToMatch`): under Board status's Schematic images, each picture with no CAD name and the
  view no picture has that its NAME or its FILE's name says it shows - the same words, a side said in
  words (top/front, bottom/back/solder) or KiCad's plot spellings (a name ENDING ".F"/".B", the layer
  "F.Cu"/"B_Cu" - **never a lone letter**: "C64 Rev B" read as the bottom, code review 2026-10-28 - nor a
  revision or model letter after a dot, "Rev.B") for a picture that is no
  schematic, a sheet's name ENDING the words (KiCad plots "<project>-<parent>-<sheet>.png"). The surest
  first; **two pairs equally sure of one picture or one view suggest neither** - no suggestion beats a
  wrong one confirmed without looking. Measured on the five live KiCad boards, CAD names taken away: 34
  suggestions, each the board's own (`SchematicCadNameSuggestionsTests` holds them). "Use" takes one,
  "Use all N" every one - ONE step either way, only what the draft can still take
  (`SchematicCadNameSuggestions.StillOpen`), saved at once; the rows changed in place (`InPlaceList`).
- **"Done - next picture"** (slice 3d - the plan: "Align: opens calibration on the Schematics tab with a
  Done - next picture button"): on the bar while aligning a picture the Contribute tab asked for, when
  another picture has a KiCad view and no alignment in the draft (`PictureToolStart.NextToAlign` - the
  next after it, round to the start). The alignment saved, the next picture is asked for BEFORE the bar
  follows, so no end is told and CRT stays on the Schematics tab; a failed save stays on this picture.
  **Main starts the next one** (`AlignNextRequested` -> `Main.AlignNext`): a board read again that waited
  for the tool is read THEN, with the next picture to come back to, and the tool started by that load -
  run as the tool went off, its reset forgot the next picture halfway.
- **"Did you mean Commodore?"** (step 3's slice 3e, 2026-10-28; CRT.Data's `SimilarNames`, `NewBoardWindow`):
  under the new-board form's Manufacturer and Hardware boxes, a name close to one CRT has - the same
  letters, digits and "+" without regard to capitals or punctuation, or a slip away (1 letter for 4-7, 2
  for longer; never in 3 letters or fewer, **never one that changes the model** - its numbers, "+" signs
  and the letters after the last digit: "Commodore 65", "C128D", "Amiga 500+" are other machines) - with
  "Use <name>". A machine is known by its listed name AND its folder ("C64" offers
  "Commodore 64": typed as the folder, a board came under a second hardware in the drop-down lists - the
  form's example was "C64", now "Commodore 64"). Two names equally close offer neither. Never blocks Create.
- **Files dropped on "Board: Schematic images"** (slice 3e; `TabDrafts.Drop.cs`, the rule `BoardDrop`). **ONLY THAT
  PAGE TAKES THEM** (owner decision, 2026-11-03: "it should not allow to drag any schematics into" the Overview page -
  "it is not the logical place to do this"; anywhere on the tab until then): every other page refuses the drag, so the
  pointer says it. They go in as "Add schematic image(s)" adds them (`ComponentPane.AddPicturesAsync`); a folder with no
  KiCad file but pictures at its top is those pictures. A KiCad project has its own drop target on "KiCad: Files"
  (`KiCadDataPanel`) - dropped here, FOLDER or file, it is told so and no image in it is taken
  (`BoardDrop.KiCadGoesElsewhere`). **Everything dropped counts, and what is left out is said** (`BoardDrop.Note`, on the
  tab's line; code review 2026-10-28 - the first folder decided alone): every picture - loose, or at any folder's top -
  goes in together. **A hidden file is
  never a picture nor a KiCad file** (`ContributionPackaging.IsAddablePictureFile`, the scanner's walk -
  second code review, 2026-10-28: macOS's "._IMG_0001.JPG" beside every picture on a USB stick became rows
  the server refused at Submit). The dropped folder is walked ONCE (`KiCadRawFileScanner.Probe`). A folder with neither
  is said on the tab's problem line (an older KiCad's or EAGLE's project told what to do). Read under the BusyOverlay,
  **taken only by the board it was dropped on and only while the page is still open**, and **never under "CRT has to be
  updated"** (the cover lies inside the tab, so a drop reached it).
- **The KiCad import's summary** (slice 3e - the plan: "drop the project folder and get a summary of what
  was found, what was ignored, and what is left to do"; `KiCadImportSummary`, shown by the "KiCad data"
  panel). **What is left to do went on 2026-11-03** (owner request: "it should show elsewhere" - Board status says it).
  FILES - the board's KiCad files, **a line per kind, the number in bold** (owner request, 2026-11-03; so too what adding the
  components did, `BoardReadiness.AddedFromKiCad` - both carry `<b>` for `BoldRuns.ShowMarked`); FOUND - the first PCB's
  components (`KiCadComponentTable`'s, as "Add N from KiCad" counts them),
  its nets, the sheets; LEFT OUT - what the folder held that the import does not take, counted by kind,
  said only right after an import in that window that COPIED something (the folder is a path on the
  contributor's machine, stored nowhere); a schematic (.sch) or board (.brd) CRT cannot read is SAID with
  what to do - an older KiCad's (told by its FIRST LINE, `KiCadRawFileScanner.IsOlderKiCadFile`) to save it
  in a newer KiCad, another program's (EAGLE uses the same extensions) to import it into KiCad - such a
  project used to get only "No KiCad files in that folder". **The folder is walked ONCE** (`KiCadRawFileScanner.ScanSplit`: taken, left out, older KiCad), and
  **the walk, the copy and the report are ONE of the window's queued changes** under its wait (code review,
  2026-10-28: walked outside the queue, the window could close meanwhile and the import landed unseen).
  It is "KiCad data" alone since step 3; its KiCad part is `KiCadDataPanel.KiCad.cs`.
- **"Contribution in use"** (`Main.thisContributionInUse` - a draft, or unread news on something sent)
  is what the Drafts tab's visibility used to be: it decides whether CRT asks the API revision.
The folder also holds that flow's dialogs (`NewBoardWindow`, `NewBoardMaintainerWindow`,
`SubmitDraftWindow`, `DraftDriftWindow`, `MySubmissionsWindow`, `KiCadDataPanel` (a panel since
2026-10-30, not a window),
`DiscardDraftWindow`, `ForgetSubmissionWindow`, `RemoveSchematicWindow`).

- **The pictures are added, named, matched, ordered and removed ON THEIR CARDS** in the Pictures part
  (CRT 4.0 step 3, 2026-10-28 - the plan: "Pictures: today's Schematic images content, moved inline";
  owner: "do what is best for having the best UX for the user. It should be as easy as possible";
  `ComponentPane.BoardPictureList.cs`). **The "Schematic images" window is GONE**; `KiCadDataPanel` is
  "KiCad data" alone. "Add pictures" (or
  pictures dropped on the tab) copies each in at the board's top under a name of the board's own
  (`ComponentPaneFileNames.ForBoardPicture`), named after its file and never a name a picture has
  (`BoardPictureEdits`) - ONE step, saved at once; never a hidden file; **only the pictures actually
  copied get a card** (`DraftPickedFiles.Copied` - one locked or gone is said, code review 2026-10-28).
  **No card or thumbnail is read until that save is done** (`ComponentPane.thisFilesArriving`, from
  `AddToDraftAsync`; owner report 2026-11-01: a new board's pictures were grey boxes) - a copy lies in
  the draft folder until the save lists it and moves it into the store, and a read in between lost it.
  The name box renames on Enter or when left (Escape puts the name back) through `SchematicRename.Plan`
  and `ApplyWithFollow` - **never on a pause in typing** (`CommitHeld(fromTyping)`: a pause after "Top"
  saved or refused half a name). **A name refused stays in its box** with the reason, and the part
  stays open on it (`KeepsRefusedPictureName`) - "Board status", another part, a component or the
  table asked for (`ShowBoardStatusAsync`, `OpenPaneAsync`, `OpenTableAsync`); only another board shown
  closes it. The pictures'
  undo steps and "Review changes" say "Schematic image [Top] ..." and "KiCad view of [Top]"
  (`DraftStepWording`, `DraftChangeSummary`), and **"Delete"** ("Remove" until 2026-10-29 - the image
  leaves the board and its file the draft; the contributor's own copy is never touched) asks "Delete
  schematic image" with Undo to bring back what goes with it (`SchematicRemoval.WhatGoes`). **"Associate KiCad file"** ("Shows KiCad view" until 2026-11-01) picks the CAD name
  from the board's views - "Not associated" first, the rest alphabetically, numbers as numbers (owner request, 2026-11-03; `SchematicCadNameChoices`) -
  read by the host off the UI thread from the project the Schematics tab draws,
  ASKED by the pane when the part opens and again whenever the draft's KiCad files change while it is
  open (`KiCadViewsWanted`, numbered like the pins' ask - an older read is never taken;
  `TabDrafts.LoadKiCadViewsAsync` -> `UseKiCadViews`; the pane's board is set BEFORE it shows a part, or a
  read done at once was dropped); a view one picture has is LEFT OUT of the other pictures' lists, not
  greyed out (owner request, 2026-10-10; `SchematicCadNameChoices`). **Each card's drop-down list is
  changed IN PLACE** (`ApplyCadNameChoices`, `InPlaceList.Sync`) - a replaced ItemsSource lets go of
  the selection and the binding writes "nothing" back over the CAD name; a null pick is never a pick.
  Move up/Move down move one place (`BoardPictureEdits.Moved` - dragging went with the window). **A card
  is found by its name AND its image file** (`SchematicRows.RefOf`) - two rows can share a name, and by
  name alone a change went to the first and Remove took both. **Remove asks first when the picture has
  highlights or a KiCad calibration** (`SchematicRemoval`, `RemoveSchematicWindow`) - they go with it,
  and Undo brings them back. **A rename or a view picked changes its card IN PLACE**: a name commits as
  its box loses focus - on the PRESS of another card's button - and cards made again would throw that
  click away; only adding, moving and removing show the cards again.
- **"KiCad data" is a PANEL on the right of the Contribute tab, not a window** (owner request,
  2026-10-30: "I do not understand why the 'KiCad data' opens in a new window? ... it should just show
  up in the 'normal' content area"; `KiCadDataPanel`, the BoardFilesWindow until then, hosted by
  `TabDrafts.KiCad.cs` in `KiCadPanelHost`). ONE thing on the right at a time: opening it closes the
  table and the pane (what they hold saved), and opening either, or "Components", closes it; a NEW
  panel is made for each opening, as the window was. Its waits run under Main's BusyOverlay
  (`BusyOverlay.For` finds the window it is in) - it has none of its own; a test hosts it in a window
  with one (`KiCadDataPanelTests.Hosted`). **Each change it makes is caught up with AS IT ENDS**
  (`KiCadDataPanel.DraftChanged` -> `CatchUpAfterBoardFiles`: the row, the left column's count, the
  board on screen read again) - a window did that as it closed, and a panel can stay open while the
  contributor goes to the Schematics tab. **And it follows what moves the draft under it**
  (`KiCadDataPanel.FollowDraft`, from `TabDrafts.AfterDraftChanged`; code review, 2026-11-01): the
  header's Undo and Redo and "Undo changes to" stand beside it, so its labels, signals, summary and
  report are read again, queued with its own changes, and what the last add said goes. It has no Close
  of its own - another entry of the left column closes it. **Closing waits for its queue** (`WaitForChangesAsync`) -
  except another board shown or its draft discarded, which cannot be called off: it goes at once
  (`CloseKiCadNow`), its changes end by themselves and are still caught up with, and a draft made only
  to open it is let go of only when nothing is still being written into it. Quitting waits for it too
  (`SaveDraftEditorsNowAsync`).
- **Every change "KiCad data" makes to the draft is ONE queued change** (code review, 2026-10-10;
  `KiCadDataPanel.DraftChanges.cs`): applied to the draft held in memory and saved off the UI thread
  under CRT's BusyOverlay (`ChangeDraftAsync`), each after the one before it (a click whose press came
  before the overlay still lands). A test waits for `DraftChanging`.
- **What hangs on a schematic's NAME follows it in every draft write** (owner request, 2026-10-10;
  `DraftWorkbookStore.EditCore`): the sidecar files highlights and the KiCad calibration by schematic
  name, so a rename moves them and a deletion drops them - the server's own rule for a table edit
  (`SubmissionRowsBoard.FollowSchematics`: same image file, new name = a rename). A picture card's
  name box names its rename outright (`SchematicRename.Plan`, decided on the draft held in memory and handed in
  by `DraftSession.ApplyWithFollow`). **On a card the credit for the image follows too** (owner request, 2026-11-04; `SchematicRename.Apply`): a
  credit with "Schematic image" under "What" and the image's name under "Sub-category" - 59 live rows, none of another kind naming an image, so
  "Oscilloscope baseline" with "PAL" stays whatever an image is called. A name changed in a "Board schematics" CELL leaves "Credits" alone: the
  table is not built again for its own change, and the Credits sheet is a click away. A KiCad calibration is looked up without regard to
  capitals (`BoardComponentHighlightStorage`), like every other schematic name, and a write reads the
  sidecar's calibrations ONCE (`LoadKiCadCalibrations` - it parsed the file once per schematic, twice).
- **A renamed schematic image is followed by every user's own data** (owner decisions, 2026-10-10:
  "it must also change for everyone updating to this data ... in the board JSON file, as I do not
  want this data in Excel", with the image file "to be unambiguous"). **CRT follows the IMAGE FILE,
  not the name**: it remembers each board's images as last shown (`UserSettings.SchematicImagesByBoard`,
  name + file), and whenever a board is SHOWN `LocalSchematicRename.CatchUp` moves worklog areas, image
  orders, hidden images and the last image to wherever their file is now (`SchematicRenameLog.Resolver`)
  - a published rename, the user's draft rename, a draft discarded, a BETA push-back, a swap, a name
  given to another image. The board's `.json` carries a third root, **"Schematic renames"** (oldest
  first, `{From, To, File, Date}`), which the SERVER adds to at each publish
  (`ApprovePublishFlow.RenameLogFor`: same image file, new name) - nothing is sent with a submission.
  It is used only when CRT has nothing remembered for the board; then a name the board HAS stays.
  **Every sidecar writer passes the list on** (`BoardSidecarWriter` builds the whole file; a list it
  is not handed is gone).
- **The KiCad data panel lists every KiCad net** (nothing to click since 2026-11-01 - owner request: "drop
  that 'copy' option ... as this will not be useful", every net has its "Add" under Important KiCad signals;
  one an important signal uses is a FILLED green badge with white text - owner report, 2026-10-29: green text on grey was
  "almost impossible to see"; "The 45 nets in green ...", `KiCadImportedFiles.UsedNetNamesSentence`) and the Important
  signals rows naming no net. **With no import of the draft's own, its file list says whether the board
  HAS KiCad data** (`KiCadImportedFiles.StateLine` - owner report, 2026-10-29: "No KiCad data imported
  yet" on a board with published KiCad files). `KiCadNetNames.Resolver` IS the Schematics tab's Important signals panel's rule
  - change it there, never beside it. **Its whole report reads the files the Schematics tab draws
  from - the published "KiCad data" AND the draft's** (`KiCadRawFileScanner.ScanBoard`); its file
  list (what Remove deletes) stays the draft's. KiCad's "no net" is the LOADER's mark
  (`KiCadNetRef.IsNoNet`: a declared `(net 0 "")`) - a KiCad 9 net NAMED "0" reads the same by id
  and name.

- **The badge** (`Main.SubmissionChecks.cs`): `SubmissionReceiptStore.UnreadCommentCount`, the same
  number as "My submissions". Sent submissions are asked about at launch and every minute while the
  window is not minimised and something is open; each check is followed by the launch path (retire
  drafts, then `ApplyContributeTabState`) on `onFinished`, never `onChanged`.
- **A send that never reached a maintainer frees a new board's name** (code review, 2026-11-02;
  `SubmissionReceiptPresenter.NeverReachedReview`: no state yet or "uploading", unless the finalise request
  went and nothing has been heard since - `SubmissionReceipt.FinaliseSentUtc`), and answers no change
  request (`NewerForBoard`). **A listed board's FOLDER is taken whatever the board is called**
  (`BoardNameFields`). **A refusal's reasons are kept on the receipt** (`SubmissionReceipt.Findings`, state
  "rejected" only, its ERRORS only - `ReasonsRefused`: "rejected" is a maintainer's Reject too, and the server
  answers the warnings a queued submission held), which "My submissions" shows. **A receipts file that could
  not be read is never written over**: read again at the first write, and while only held at each minute's check and as
  "My submissions" opens (`SubmissionReceiptStore.ReadAgainIfHeld`, its card going once it is back - never from `All`, which
  every refresh reads); set aside as `<file>.broken` only when its contents are broken; while it is held or access to it is
  refused, what is recorded waits beside it in `<file>.waiting` (a token's only copy, code review 2026-11-03), which the next
  read that works takes in. **The fingerprint holds a new board's notes as sent** (`TabDrafts.FingerprintOf`).
- **What was just sent cannot be sent again**: the receipt keeps CRT.Data's `DraftFingerprint` (the
  workbook's VALUES, the sidecar's bytes, every draft file by path and content). While the draft
  matches its LATEST submission, Submit is greyed out with the reason - unless that submission never
  finished sending. An empty fingerprint never blocks. `TabDrafts.SubmitAsync` checks again after
  settling the table's edits.
- **The Submit dialog is no taller than its screen, and centred over CRT whenever its size changes**
  (owner report, 2026-11-01: "I cannot see the submit button", and its progress panel stood at the top):
  its panels scroll, the buttons below them (`SubmitDraftWindow.OnOpened`, `DialogPlacement`) - a window
  sized to its content keeps its top edge as it shrinks.
- **What a submit sends is the draft WORKBOOK read from disk** (`TabDrafts.BoardToSendAsync`) - the
  file its fingerprint is taken from. Never `DataManager.LoadBoardDataAsync` there: it answered from
  its cache (an image renamed for a board not selected in the drop-down lists went to the queue
  under its OLD name, 2026-10-11) and answers the PUBLISHED board under "view boards as officially
  published". `TabDraftsSubmitBoardTests`.
- **Several drafts in one submit** (slice 5, 2026-10-26; `SubmitDraftWindow.OtherDrafts.cs`,
  `TabDrafts.SubmitOthers.cs`, words `SubmitDraftsWording`): the Submit dialog lists the OTHER drafts
  "My drafts" holds, UNTICKED, each greyed out with its row's reason (nothing in it, already sent,
  errors). Each one ticked is sent AFTER the board's own, as a submission of its OWN (the server and a
  maintainer take one board at a time - no server change), with the same address and message. It is
  made ready only once Submit is pressed (`PrepareOtherDraftAsync`: its session saved, the same checks
  as its own Submit, a draft with errors NOT sent - no table can open from inside the dialog - and
  credited under the name typed, for its own work). One that cannot go is said by name and never
  stops the rest; several boards' outcome lists each. Nothing ticked: the dialog is as it always was.
  **The other drafts are credited only while "Credit me on the other drafts I tick too" is ticked**
  (`CreditOthersCheckBox`; unticking every credit of this board unticks it - code review, 2026-10-26),
  and **a credit put into a draft before a send that then failed is SAID** - the draft is not
  "unchanged" (`SubmitDraftWindow.NoConnectionMessage(credited)`, `SubmitDraftsWording.Intro`).
- **A submission the server no longer knows reads "No longer on the server"** (404 after a board
  deletion or the data reset; `SubmissionReceipt.NotFoundUtc`): neutral colour, never blocks Submit,
  no BETA try, not reported on a discard.
- **The draft's header** (CRT 4.0 step 1's last slice, 2026-10-25, at the top of the Contribute tab
  since step 2; `TabDrafts.Header.cs`,
  `DraftHeaderState`): "Could not save - <reason>" with Try again when a save failed,
  Undo and Redo naming their step in the tooltip, "Review changes", and "Discard draft". **One
  `DraftHeaderState` per draft for the run**, compared by reference (`DraftListItem.ShowsTheSameAs`), so
  a save never rebuilds the row; every refresh reads it again (`Follow`). **The row's Undo and Redo are
  the SESSION's**: table closed, the step is SAVED at once and the board on screen reads again
  (`AfterDraftChanged`); table open, they are the table's. **Try again on a file changed somewhere else
  opens the table** (its bar offers both versions) - tried again it is refused for ever. The header's
  save failure hides while the table is open; the table's own line says it.
- **"Review changes" on a new board says where it stands and lists nothing** (owner request, 2026-11-03: "the only thing it needs to state is
  'New board - not submitted yet' ... viewing 'a million' changes on a new board is not helpful"; `ReviewChangesWindow.InitializeNewBoard`):
  "New board - not submitted yet", or once `SentReceiptsOf` has a send that still stands (`SubmissionReceiptPresenter.StillStands`, the rule "Board: Name" keeps its name by - it reached a maintainer and nobody decided against it), that it is submitted and not published,
  or once one is in the stable data, that it is published (`NewBoardStanding`, asked first: a board reads as new here until this machine has downloaded it, code review, 2026-11-04).
  Nothing here holds a copy of what was sent to compare with, so changes are listed once it is published. The Submit dialog's "Your
  changes" is as it was.
- **"Review changes" and "Undo changes to [U8]"** (`ReviewChangesWindow`; CRT.Data's `DraftChangeSummary`,
  read by `DraftChangeSummaryReader` - the session's content, else the board file, against the published
  board file with both file lists): one group per component, plus "The board". Each COMPONENT group
  has a red "Undo changes to [U8]" (`ComponentRevert`: every row of the label, all regions, back as
  published - a relabel back to its old label, an added component removed) as ONE step
  (`DraftEdits.UndoComponentChangesAsync`), which the row's Undo takes back; the board's group has none.
  **A file of the component replaced under its own name is a change, and goes back too** (code review,
  2026-10-26; `ComponentRevert.ReplacedFiles` against both lists - the pane reads the draft's from
  `DraftSession.ListedFiles`, never the file): put back by the draft's file list after the rows
  (`DraftEdits.PutBackPublishedFilesAsync` -> `RemoveFiles`), a write the row's Undo does not take back.
  Rows already as published are no step (`ComponentRevert.RowsToPublished`). **The highlights go back
  under the draft's names for their schematics** - an image the draft renamed keeps its new name
  (`SubmissionRowsBoard.FollowSchematics`; code review, 2026-10-26: under the published name they were
  refused at Submit, fixable nowhere in the table), and the rename alone is no change of the component.
  **The Submit dialog shows the same summary** ("Your changes") and **Discard draft lists what is lost**
  (`AsLines(max: 12)`) - the words are `DraftChangeSummary`'s, nowhere else.
- **Unfinished work is saved** (the plan: "Only Submit asks for problems to be fixed"): a category
  emptied, an image row with neither file nor note, a picture CRT cannot show are left to the checks. One exception holds a row back in the pane instead: a scope baseline without its pin, a
  setting and a screenshot (else it is a picture). A row of the board's files, links, credits or
  important signals with a field it needs empty goes in, an error (`BoardRowRequired`, above - held back
  too until 2026-11-01). A file goes in its kind's usual folder (`ContributionFileLocations.DefaultFolderFor`, through
  `ComponentPaneFileNames`). **A save not made takes its copied files back** (`DraftPickedFiles` - a
  file a copy replaced is set aside in `.crt-set-aside` and put back; code review, 2026-10-26): left
  lying in the draft folder, the next write by ANY editor took them in after "could not be saved". A
  file replaced under the name its row cites changes no row, so the draft is written anyway
  (`DraftSession.SaveAsync(evenUnchanged)`).
- **A NEW DRAFT IS SEEDED FROM THE PUBLISHED FILE AS IT IS ON DISK** (`DraftSeeder.SeedFromPublishedFile`),
  never from `BoardDataReader`'s cache - an Excel edit made while CRT ran would be lost.
- **A draft keeps its BASE** (code review, 2026-10-13): `.crt-base.json` (`DraftBase`) - the published
  board as copied, its calibrations, the copied files' hashes - written at seed and by "Bring my draft
  up to date" (`DraftRebaser` around `DraftRebase`'s three-way merge: rows the contributor left alone
  follow the update, theirs win where both changed, conflicts listed). Without it a stale draft sent
  the update's rows back as removals. **"I have looked at this" (`MarkDriftSeen`) never moves
  `BaseRevision`** - the server is told it. Drift is judged by CONTENT (`DraftDriftCheck`,
  `DraftMarker.BaseContentHash`), revisions being dates - and "I have looked at this" records the
  published content it was said to (`DriftSeenContentHash`). **The marker is noted as up to date only
  once the new base is kept** (`DraftRebaseOutcome.NotNoted` otherwise): a marker ahead of its base
  made every later run "already up to date". Bringing up to date also removes untouched copies of files
  the merged board no longer cites, and follows the update's own renames (`FollowUpdateRenames`). Files named `.crt-*` are bookkeeping
  (`DraftFolderLayout.IsDraftOnlyFile`). Files and calibrations count for Submit (`DraftOtherChanges`,
  measured against the base) - and row ORDER since 2026-10-30 (`BoardRowOrder`, against the published
  board, as rows are counted).
- **Every editor changes the ONE draft held in memory** (CRT 4.0 step 1, 2026-10-25; `DraftEdits`):
  the component pane, the label editor, the KiCad calibration and "KiCad data" hand their
  change to the draft's `DraftSession` and save it at once. **Only handing over what the table holds
  and applying the result run on the UI thread** (code review, 2026-10-25): the session is opened,
  the file asked about (`FileChangedOutside`) and the edit worked out with what follows it
  (`DraftSession.Prepare`) on the pool, then applied (`ApplyPrepared` - worked out again when the
  draft moved meanwhile). **A save that cannot be made is taken back out** (`DraftSession.TakeBack`,
  by the rebase's three-way merge when the table pushed a change after it - the later change stays,
  and no undo brings the taken-back one back) and the editor says why
  (`DraftSaveWording.EditNotSaved`) - an explicit save either happens or changes nothing. A transform
  returning the board it was handed is no step at all. Before any
  change the table hands over what it holds (`ChangeComing`), so the change lands on top and one save
  writes both; **nothing holds another editor's save back** - the `SavingElsewhere` notices and
  `TabDrafts.SavingBlockedFor` are gone. Each change is ONE named step of the draft's history
  (`DraftStepWording`), so the table's Ctrl+Z takes it back. A write that is MORE than the content -
  "Bring my draft up to date", Remove of a KiCad file - goes through `DraftEdits.WriteFileDirectlyAsync`
  (held changes collected on the UI thread, saved first, the file read back; bringing up to date
  clears the history, measured against the old base); a KiCad import takes its files in with
  `TakeInNewFilesAsync`. "KiCad data" settles what the table, the pane and the draft hold before it
  opens (`SettleTableEditsAsync`, while they are all still open, so Cancel leaves them as they were -
  the panel's move from a window dropped it, code review 2026-11-01), so a change that cannot be saved
  is asked about once, not behind every click.
- **Each draft's row counts problems too** (`DraftStatusReader.CountProblemsCached`), and the list is
  read again when the tab is shown or CRT's window is activated - when an Excel edit arrives.
  **That read keeps every unchanged row** (`TabDrafts.KeepUnchangedRows` - `InPlaceList.KeepUnchanged` -
  and `DraftListItem.ShowsTheSameAs`)
  - never Clear and rebuild: a click INTO CRT activates it between the press and the release, and a
  rebuilt row threw the pressed button away ("I need to click it twice", 2026-10-10).
- **Retiring a published draft** (`PublishedDraftRetirer`, CRT.Data's `DraftRetirement`): a draft whose
  work reached the STABLE source is deleted (BETA alone is not enough - a push-back can undo it), after
  every launch status check and when "My submissions" closes; never one holding a change not saved
  yet (`TabDrafts.HasUnsavedTableEditsFor`, `DraftManager.HasUnsavedChanges`). **Nor one made after the
  board's deciding submission was sent** - it holds none of that work (code review, 2026-11-02;
  `DraftRetirement.WasMadeAfter`; a marker or receipt with no time is judged by what the draft holds).
  **What holds a draft is one question, `TabDrafts.IsDraftInUse`** - anything open or opening on it, a
  change not saved, a step to redo - and a later check retires it; `MainRetirePublishedDraftsTests` drives
  the whole step with the test-server gate passed in. **Every file a draft lists must be the published
  board's, name and checksum**: bytes in the tree's store say nothing of this board. A keep caused by
  bytes on their way to the tree's store is not remembered (the store is in neither stamp).
  **A receipt names its board by ID; the drafts flow is keyed by WORKBOOK path** -
  `DraftStatusReader.ResolveForBoard` maps one to the other. A new board's draft is retired once its
  published board is on the machine. A user on the BETA source is told once, when their work reaches
  stable, to switch back (`Main.SourceSwitchNotice.cs`). **Nothing is retired while CRT talks to the
  test instance** (`ContributionServer.RetiresPublishedDrafts`; code review, 2026-10-25): its stable
  data never reaches the live data, so the draft is the work's only lasting copy.

#### The board table editor (`Controls/BoardTable/`, shown by the Contribute and Maintainer tabs)

"Table" shows the shown board's draft as editable sheets on the right of the Contribute tab, the left column picking the sheet
(`TabDrafts.Table.cs` owns table mode, `TabDrafts.Entries.cs` what shows; `BoardTableEditor` the control; `UnsavedTableEditsWindow` the prompt; colours in
`BoardTableColors.axaml`). Colours are the owner's: green added, orange the changed CELL (published
value in its tooltip), red + strike-through deleted, shown where the row used to be.

- **The Contribute tab's left column picks what it shows** (`SheetsPickedElsewhere`, `ShowSheetPart`, `BoardTableEditor.SheetParts.cs`;
  owner request, 2026-11-03): no sheet tabs there. **The Maintainer tab's tabs are the same entries, under the same names and headings** (owner request,
  2026-11-05: "group and order into what is logical for the new setup"; CRT.Data's `BoardTableTabs`, `ContributeEntriesTests` holding the two together):
  Board - "Schematic images", "Files", "Links", "Credits"; KiCad - "Important signals"; Components - "Basics", "Replacement parts", "Pinout", "Pictures",
  "Oscilloscope baseline", "Files", "Links". A heading is a `TabItem` that cannot be picked; KiCad stands after Board there, as a table opens on its first tab and most
  boards have no important signals. **The strip is a LINE PER GROUP, never a paragraph** (owner request, 2026-11-06: "it wraps into next line, which
  makes it hard to overview"; `SheetTabsPanel`, every position `SheetTabsLayout`): Board and KiCad on the first line, Components on its own
  (`BoardTableTabs.StartsLine`), the two line headings in one column so the tabs line up, a long line wrapped under its own tabs. A tab counts its own changes (`BoardTableSheetParts.ChangeCountOf`), a host remembers the part with the sheet, an undo goes to the
  tab and column holding its change, and `BoardTableDocument.Sheets` and the list of problems over the table go in this order, the list naming each by its tab ("Components: Basics", `BoardTableTabs.NameOf`); a workbook keeps `AllSheets`'. With `SheetsPickedElsewhere` no strip is built. **"Component images" is shown in
  two, "Pictures" and "Oscilloscope baseline"** (CRT.Data's `BoardTableSheetParts`): ONE sheet, as the board file and the server keep it; a row is
  a baseline by its pin alone, so a pin typed moves it across once the cursor leaves it; a row with nothing typed shows in either. **A row that can
  still be a baseline also shows among the baselines, where its pin can be typed** (`MayStillBeABaseline`; code review, 2026-11-04 - neither "Pictures" nor the
  pane's page has a pin field, so a pin cleared, or "Insert row" and a label typed, left the row out of reach): worked on there while the table stands
  (`BoardTableRow.StaysWithBaselines`), holding what only a baseline has (15 published rows), or with nothing yet that says picture. A
  picture's part leaves out the columns only a baseline has, so **whatever walks along the columns goes by the columns shown** (`ColumnsShownNow`:
  Tab, `PasteBlock(columnsShown)` - walking the sheet's, a pasted value landed in the hidden Pin and made the picture a baseline), **and its pills, search and
  first problem cell judge only those** (`CellsShown`). Such a part counts as narrowed (`IsNarrowed`): no row moves by its grip, and a paste adds no rows.
- **The control only paints.** Every rule is in CRT.Data (`BoardTableDocument`/`BoardTableSheet`,
  `DraftSession`, `BoardTableClipboard`, `BoardTableHistory`, `BoardTableRowDrag`,
  `BoardTableSearch`, `BoardTableRowFilter`), because the Maintainer tab shows the same table.
- **Pairing is `BoardDataDiffer`'s rule** (`BoardDraftNaturalKeys`; keys case-insensitive, values
  trimmed + ordinal, first row per key wins - **except a component's rectangles on one schematic, which
  all count**: a highlight's key names a component on a schematic, not a rectangle, so every comparison
  lines them up by all their values, each row's place held as `ComparedRow`, a natural key and an
  occurrence number, never one string; their order alone is no change. A key can stand several times in
  a section's lists: `FieldChanges[key]` holds the fields together, and the optional `ChangedRows` says
  what moved on each, only when a key repeats. A rectangle that only followed its schematic, renamed or
  removed, is the board's change: no group, no undo and no "Draft" mark. A row whose key changed carries
  `BoardRowChange.PreviousNaturalKey`, and both labels get the mark (code review, 2026-11-02;
  `BoardDraftNaturalKeys.MayRepeat`, `BoardDataDiffer.LineUpRepeatedRows`, `ReviewSummary` too)), so the tabs' counts equal the draft row's "N rows
  changed". **A key edit with nothing else changed is ONE row modified** (`BoardDataDiffer.PairRenamedRows`,
  owner decision 2026-10-04 - used by the table, the differ and the server's `ReviewSummary`); a key
  AND another cell changed is an add plus a deleted ghost. **A component's identity is its label
  PLUS its region**; **an important signal's is its display name PLUS its KiCad net**.
- **In the Contribute tab the table is a view of the draft held in memory** (CRT 4.0 step 1, 2026-10-25;
  `BoardTableEditor.DraftSession.cs`, the plan: "What goes away: every Save button in drafts, the
  'unsaved table edits' prompts ... and the table's file watch with its Excel bar"). It works on the
  draft's `DraftSession` from `DraftManager.Sessions` (kept for the run - a draft discarded or retired
  takes its own, `DraftManager.DiscardDraft`). **Every settled change is pushed into the session as
  ONE step** (`BoardTableDocument.TakeSessionEdit` - null, so no step, when no row changed) carrying
  where it was (`BoardTableHistory.Latest`, `BoardTableSessionEdit`); never while a cell is typed in or
  a row dragged. **Only the sheets changed since the last push are built, put in and compared** (code
  review, 2026-10-26 - all nine, three times over, on the UI thread at every edit): every change a sheet
  makes calls `BoardTableDocument.MarkDirty(sheet)`, and a sheet not named there is left as the session
  has it, its very rows - a new kind of change MUST call it too. **The session saves itself** a second after a change, off the UI thread (the save
  timer, `DraftSession.IsSaveDueNow`); where "Save changes" was the table says nothing while all is well and
  "Could not save - <reason>" with Try again (`DraftSaveWording`). **Undo and redo are the session's**,
  and the table is BUILT AGAIN from it after each (`ShowSessionContent`) - so row and cell objects do
  not survive an undo; look them up again. **A file changed outside the table** is read again when CRT
  wrote it (`ReadDraftFileAgain`, an undoable step) and caught by any save after it: `ChangedOnDiskBar`
  offers "Use the saved version" (undoable) or "Keep the table's version" (`Save(overwrite)`) - nothing
  thrown away unasked. **New components stay where they were typed** while the table is open and are
  saved into their category (`BuildRows`), said once. **What goes with a change is said once made**
  (`BoardTableSessionWording.WentWith` - the save-time "Remove schematic image?" question is gone).
  The Maintainer tab's tables are documents (`Open`) and keep their explicit Save.
- **The unsaved-edits prompt has wordings** (`UnsavedTableEditsPrompt`): LeavingSubmission and
  LeavingBoard (the Maintainer tab - Save/Discard/Cancel), LeavingUpdateRequired, DraftChangedOnDisk and
  DraftNotWritten (a draft's table whose save just failed - Discard/Cancel, no Save). A draft's table
  asks only after its save failed: closing it, submitting, managing files and quitting save first
  (`TabDrafts.SaveTableNowAsync`).
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
  - The table must never sit inside a ScrollViewer (it would realise every row) - the Contribute
    tab's `WorkspaceGrid` holds the table and the start panel as siblings in a `*` cell, and
    `ApplyTableMode` only switches which is visible. A Grid's `RowDefinitions` are never replaced
    (assigning new ones drew the open row at zero height).
  - Every key the table's markup names must resolve in CRT's App.axaml (`SharedTableColourKeysTests`);
    a missing `DynamicResource` draws nothing, silently.
- **The table owns the keyboard while open**: the component filter backs off on the whole Contribute
  tab (`Main.ShouldReturnFocusToComponentSearch`), and `TabDrafts.FocusTableIfOpen` puts focus in the grid.
  Tab/Shift+Tab move along and wrap (tunnel route); Enter moves down. Buttons hand focus back to the
  grid (`ThenFocusGrid`) - a disabled focused button drops focus to nothing.
- **The grid's view has no filter** (code review, 2026-11-02): a filtered `DataGridCollectionView` removes
  and re-adds the row at every row commit, leaving no current cell and no selection, so what was typed
  next landed on another row. The grid shows `thisShownRows`, which is `sheet.Rows` less what the pills
  and the search hide. A pick or a search replaces it with one Reset; anything else changes it row by
  row, never under a row being edited. `FilteringModel.OwnsViewFilter = false` stays, or the grid's model
  would put a filter back. **The row the cursor is on stays shown until the cursor leaves it**
  (`BoardTableRowFilter.RowsShown`), so a pill can read 0 with one row showing; a row with no cursor still
  drops out, and a pick or a new search keeps nothing. The cursor counts only on a data cell. A click raises
  no `CurrentCellChanged`, and `SelectedCellsChanged` comes before the cursor has moved, so the update is
  posted and reads the cursor afterwards. **Under a pill, many rows changing as one are gone over once, as they end**
  (`BoardTableSheet.RowsTogether`, `RowsChangedTogether`; code review, 2026-11-03) - a new change of many rows says it is one.
  **Neither undo nor redo runs with a row held by its grip**
  (`BoardTableEditor.Undo`; `BoardTableHistory.Apply` returns null while a group is open). **The checks see
  the highlights as the board will hold them** (`HighlightsAsSaved`, against the rows they were last
  followed to: a session follows at each hand-over, a document at Save).
- **Undo/redo (Ctrl+Z / Ctrl+Y)**: a draft's are its session's (above); a document's live in the model
  (`BoardTableHistory`): each step is a snapshot of one sheet's live rows taken just before the change,
  back to the last save. **A new kind of change records `History.Record` before it mutates** - a
  document's undo needs it, and a draft's step takes where the change was from it. Inside a cell being
  edited, Ctrl+Z is the text box's. Undoing back to the saved state clears "unsaved".
- **Rows move by the editor's own drag** (`BoardTableEditor.RowDrag.cs`; ProDataGrid's row drag is
  off): the row moves LIVE with a dashed placeholder, the target read off a FROZEN layout
  (`CaptureRowSlots`) - hit-testing the live grid made the row run away near the bottom edge. A red
  ghost is never a target; a whole drag is one undo step. Alt+Up/Down remain; no move buttons.
- **The colour key IS the filter**: five pills - Added, Modified, Deleted, Errors, Warnings - each a
  toggle; picked ones show rows of any picked kind, and **count the WHOLE DRAFT** (owner decision).
  A pick also hides the tabs of sheets it shows nothing of (`SheetsShown`/`SheetToShow`), and on the Contribute tab
  turns those entries of the left column off. Moving rows
  is off while anything is picked. Three looks: counting nothing (outlined, 0.4 opacity, not
  pickable), counting rows (filled), picked (2px outline). A sheet tab says its name and change count
  only.
- **The checks are in the table** (`BoardDataChecks`, one rule set with the server): **an error in the
  table is a refusal at submit** (`Every_error_the_table_finds_is_one_the_server_refuses`); a new local
  rule is a warning. A problem cell gets a corner triangle drawn as ONE background brush; a problem in
  no row is a line above the table. Duplicates are WARNINGS (owner decision 2026-10-03), marked on every
  row of the set; a save drops no row since 2026-11-01 (an important signal missing a half was the one,
  a warning; `BoardTableRowState.Incomplete` went with it, `BoardWorkbookSchemaTests` holding that no
  mapper drops a row with any text in it) - a field a row needs is an error (`BoardRowRequired`). Submit on a draft with errors sends nothing and opens
  its table on them. One warning is the table's alone (the Everything scope, never the server's;
  2026-10-25, since unfinished work is saved): `image.not_displayable`.
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
- **Blocks are copied and pasted, with find & replace** (CRT 4.0, 2026-10-25 - step 0
  of the plan, so the table does what Excel was used for; `BoardTableEditor.BlockEdits.cs`, the rules
  `BoardTableSheet.PasteBlock`/`PasteIntoCells` and `BoardTableDocument.ReplaceAll`, the
  words `BoardTableEditWording`). **No fill down** (owner request, 2026-11-07: "I do not understand what
  "Fill down" does ... please remove this functionality entirely" - button, Ctrl+D and its rule gone). Ctrl+C copies the selected cells as a spreadsheet block
  (`BoardTableClipboard.ForCells`). Ctrl+V lands a block from the selection's top-left cell down the
  rows SHOWN - red rows passed over, rows added at the end unless a pill or search narrows the table,
  columns past the last and a pasted heading row left out, each said on the status line; ONE value
  onto several selected cells fills them all. Ctrl+H / "Replace": capitals ignored, the
  rows shown on this sheet - of a sheet in two the part shown, its rows and columns - **and of those columns the ones
  ticked under "Columns"** (owner request, 2026-11-07: "restrict replacement to specific columns"; a drop-down with a tick
  per column shown, "Every column" and "Unselect all" (owner request, 2026-11-07: "so I can choose only one column"),
  kept per sheet and part by column NAME while CRT runs -
  `BoardTableEditor.ColumnsTickedFor`): "Find" shows and marks only what is in them (`BoardTableSearch.Literal`'s
  `columnsOf`, `Searches`; a cell knows its column - `BoardTableCellText.ColumnIndex`), and the count and "Replace all" go by
  them. **A tick changed keeps the table on its sheet and part, whatever they then show** (code review, 2026-11-08:
  "Unselect all" moved it to another sheet holding the text, the list then ticking that sheet's columns), and the list open
  as another is shown is made again for it. **Home and End go to the row's first and last DATA column, Ctrl with them to the first and last row shown**
  (`MoveToRowEnd`; owner report, 2026-11-07): the grid's own Home went to the marker column, no cell of the board, and left
  the table scrolled. **"Find" starts with what the search box
  filtered by, the cursor after it, nothing selected** (owner request, 2026-11-07) - trimmed, as the search read it, **and the
  bar opens ON THE SHEET SHOWN** (owner report, 2026-11-08: "Replace" on "Schematic images" searched for "128" went to "Board:
  Files" - columns unticked there earlier, or a space after the word, left "Find" nothing on it; `OpenReplaceBar`, and the
  posted `TextChanged` of the text it opened with is no new search). **Its "Find" IS the
  table's search while the bar is open, as typed** (owner request, 2026-11-07: "only search for the string
  as-is and do not 'AND' multiple words or exclude with minus"; `WorklogSearchQuery.Literal`,
  `BoardTableSearch.Literal`, the cells split by `BoardTableReplaceText`): it narrows and marks the rows as
  the search does, the words in "Replace with" drawn in place of each match in a color of their own
  (`BoardTable_ReplacePreview_*`), "Replace all" counting the places in a badge
  (`BoardTableDocument.CountReplacements`, the very cells `ReplaceAll` changes), the search box empty and
  greyed out until "Cancel" (no "On every sheet" any more). Ctrl+F: the search box. **Each is ONE undo step**
  (`BoardTableSheet.AsOneChange`, a history group ended whatever happens - left open, every later
  edit joined it); a read-only table offers none of them, copying aside.
- **Row order is what the user sees, and new components are PLACED** (`ComponentPlacement`: a new
  component goes into its category in natural label order, a new category last, an edited one stays
  put; a regional twin straight after its twin). **Order is NOT a change `BoardDataDiffer` counts** -
  it is counted beside it, by `BoardRowOrder` (owner decision, 2026-10-30: order is a change - see
  "ORDER IS A CHANGE" above); the table colours no row that only moved.
- **Hover cards** (`BoardTableEditor.FilePreview.cs`): pointing at a FILE cell opens a card at once -
  a `Popup` in the window's overlay layer, light dismiss OFF, beside the cell. Every text tooltip is
  instant (`CellToolTipDelay`) and NOT hit-testable, so the pointer reaches the cell below. The bytes
  come from the host's `IBoardTableFileSource`. **A changed COLOR cell shows both colors in the same card** (owner request,
  2026-11-01: "a popup showing the colors 'Changed from' and 'Changed to'"; `BoardTableEditor.ColourChange.cs`, the rule
  `BoardTableColourCells`): a swatch and its text per side, "(not set)" for a blank one - no bytes needed, so with or without a source.
- **Every cell wraps; a double-click on a heading's edge fits the column to every row**
  (`BoardTableEditor.TextWrap.cs`, `ColumnAutoFitGeometry`; the grid's own fit measured only built rows). **No column starts narrower
  than its heading, the space before the name repeated after it** (owner report, 2026-11-11: 0px between a short column's heading and its
  edge; `FitColumnsToHeadings`, a least width set once the headings are laid out - and dropped back to `LeastColumnWidth` as the header row
  is pressed, so a drag can still make it narrower).
- **Every row is as tall as its cells ON SCREEN, the first row too** (owner report, 2026-11-11: the first row twice as tall as
  the same rows under it): ProDataGrid hides the cells of columns scrolled out of view, so they count for no row's height -
  except at index 0, whose every cell it shows (`DataGridCellsPresenter.MeasureOverride`, "|| OwningRow.Index == 0": it sizes
  columns by that row). `BoardTableEditor.FirstRow.cs` holds that row's off-screen cells to a plain row's height after each
  layout. Scrolled to a long note, every row grows; scrolled back, they all shrink.
- **Several selected cells: a veil and a 1px blue edge on each** (`BoardTable_Selection_Fill`, `..._Edge`; owner report,
  2026-11-07: the "Board label" cells "miss some dividers" - the veil is as light as the grey grid line, which vanished
  into it), never on the current cell, which keeps its frame.
- **The current cell has no fill and a 2px dashed red frame** (the template's `CurrencyVisual`
  restyled; the grid's focus visual, selection outline and fill handle hidden).
- No "Restore row"/"Revert cell" buttons (owner request, once undo existed); the model keeps them for
  the Maintainer tab.

#### Workbooks (`Tabs/Workbooks/`) and the worklog

**The tab reads "Workbooks"** - "My Devices" from 2026-11-06 (owner request, a try) to 2026-11-08 (owner request: "When we
can have cases where there might never be created any devices, then its most significant meaning of this tab is about
workbooks"; the Configuration tab too) - and only what shows devices says devices: the "My devices" list, the device form,
the bar's "Device" drop-down. The scope's choices name both lists, the workbooks first ("Show all workbooks and devices").

**WITH NO DEVICE MADE, THE WORKBOOK LIST STANDS IN MAIN'S LEFT PANEL IN PLACE OF "My devices" (below)** (owner request, 2026-11-08: "when there are no devices
created, then the list of workbooks should move into the left-side panel, as there are no devices to display there. This
would make the content area larger for the actual worklogs"; `TabWorkbooks.PlaceWorkbookList`, called at every device
refresh): the tab's `WorkbookListArea` - its own card styles included - goes into `DeviceListHost` in place of
`DeviceListArea`, looking as that does there (`WorkbookListHeading`, "Workbooks" with its count badge, over a framed list,
"Clear" under it; the tab's "N workbooks" line hidden), and the tab's list column and splitter shut (width 0, the saved width
untouched and put back - `ApplySplitterWidths` leaves the column alone meanwhile). The first device made puts both back.
`MainDeviceListTests.With_no_device_the_workbooks_stand_in_the_left_panel_and_the_first_device_puts_them_back`.

**WHILE THE TAB IS SHOWN, ITS "My devices" STANDS IN MAIN'S LEFT PANEL, THE WORKBOOKS IN THE TAB** (owner requests,
2026-11-06: "when in the 'My Devices' tab, I want the 'My devices' list and the 'Workbooks' list to show in the left-side
panel, where the component category and component list is", then "I have changed my mind in terms of moving the
'Workbooks' to the left-side panel - please again have them visible where they were previously, but do keep the 'My
devices' in the left-side panel"; `Main.DeviceList.cs`, `TabWorkbooks.ShowDevicesIn`): Main takes the tab's
`DeviceListArea` - its heading over its framed list - once, at construction, into `DeviceListHost` (the line under it,
`DeviceListSeparator`, hidden); the
category and component lists (`ComponentListsGrid` - their container, hidden whole, as the Contribute tab switches the
lists in it) and the buttons under them (`ComponentActionsPanel`) give way while the tab is shown, Hardware and Board stay.
**The list stays the tab's**: its controls in the tab's name scope (a test's `tab.GetControl` finds them, and a tab built on
its own keeps it above its workbooks), filled by the tab's code. **Its card styles go with it** -
`TabWorkbooksCardStyles.axaml`, included by the tab AND by the list's border: a style of the tab no longer reaches a control
standing outside it. **No "Add device" beside it** (owner request: "no need for it") - the worklog bar's. **Its heading
stands over the frame as "Hardware" and "Board" stand over theirs, not bold** (owner request, 2026-11-06: "make it a
headline for the panel below"), **how many it lists in an outlined badge after it** (owner request, 2026-11-07;
`DeviceCountHost`), **and "Clear" stands UNDER the frame, outside it, in line with the workbook list's "Clear"** (owner
requests, 2026-11-06: "should always be visible even if scrolling down in list", and 2026-11-07: "outside of the 'My
devices' panel ... vertically align the two buttons"; `ClearDeviceButton`, greyed out with nothing picked; the two the same
height above the window's foot). **A card is up to three lines** (owner requests, 2026-11-06 and 2026-11-07; `BuildDeviceCard`):
the name and status; the serial number or unique identifier with how many workbooks the device has at that line's end - **a
device with none has no blank line**, its count moving up beside the status; its hardware and board, **its condition in stars
at that line's end** (none for 0). **THE ORDER IS "Sorting of devices"** (owner request, 2026-11-07; the Configuration tab's
radio buttons under the scope, `UserSettings.DeviceSorting` - `deviceSorting`, "Name" by default - every rule
`OwnedDeviceSorting`): by name (`NaturalLabelComparer`), by hardware, board and name, or BY HAND - each card a drag handle
at its right edge (`DeviceListCard`, an `IDraggableRow` that is its own container, `ListRowDrag`), the order stored on the
devices themselves (`OwnedDeviceRecord.DisplayOrder`, `OwnedDeviceStore.SetOrder`; a new device goes last, one never placed
after the placed ones by id). The bar's "Device" drop-down lists them in the same order (`TabWorkbooks.SortDevices`).
**"Find a previous repair" searches the devices too** ("so it also covers devices and not only workbooks/worklogs";
`WorklogSearchIndex.ForDevice` - the card's name, serial and board, and everything typed on the device), the list narrowed to
what it finds and the match marked on the three lines (`OwnedDeviceWording.NoDevicesMatch` with none).
**WHAT BOTH LISTS SHOW IS ONE SCOPE** (owner request, 2026-11-06: "This should now follow devices and not workbooks" -
"Show all devices", "Show only devices for selected hardware" in the middle, "Show only devices for selected board"; asked,
"Follow the same scope" for the workbooks; `MyDevicesScope`): every device and workbook, those of any board of the
hardware shown, or those of the board shown - the default. Stored as the workbooks' scope always was (`workbooksScope`,
`UserSettings.WorkbooksScope`), "CurrentHardware" its new value, which CRT 3.x - reading the same file - takes as the board
shown. With devices but none in the scope the list says so (`OwnedDeviceWording.NoDevicesInScope`). Every device is still
read for a device's page and the lookups; only the cards follow the scope.

`TabWorkbooks.axaml(.cs)`, `.BoardPreviews.cs` (the board pane), `.Summary.cs` (the summary strip),
`.Export.cs` (PDF/ZIP), `.PreviewReorder.cs`, `.Devices.cs` (the user's devices). A workbook is a
repair on one board; it holds worklogs.

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
- **The user's own devices** (CRT 4.0 step 4, 2026-10-28 - the plan: "Your own machines, each with its
  repairs"; owner decision 2026-10-24: "Device" wherever a person reads it, `OwnedDevice` in code -
  minipro's "device" is the chip type; `OwnedDeviceStore`, `TabWorkbooks.Devices.cs`). **Stored like
  workbooks**: one `device_{id}/` folder each in the Workbooks folder (its index.json and photos), ids
  from the root index.json's `lastDeviceId` - BESIDE `lastWorkbookId`, one file, so each allocation
  keeps the other's (`WorklogManager.AllocateNextDeviceId`) - never handed out again, **nor an id a
  workbook still names** (a lost counter brought a deleted device's id back, its repairs on the new
  device); a field a later CRT writes is kept by the device, the workbook record and the root file
  (`JsonExtensionData`). **A repair names its device** (`WorkbookRecord.OwnedDeviceId`, null for none,
  `WorklogManager.SetWorkbookDevice`), never a list on the device: a deleted workbook leaves nothing
  behind, and an id no device has reads as none (`OwnedDeviceRepairs.DeviceOf`).
  "My devices" lists the devices of the scope (above), each card saying how many workbooks it has (owner request,
  2026-11-06: "Just as in the 'Device' drop-down list"; `BuildCountPill`, on its second line); the card marked is THE DEVICE
  PICKED (below), and a card clicked picks it and opens its PAGE in place of the workbook: what it is (status as words,
  `OwnedDeviceStatus`, its pill `WorklogInfoPillBuilder.BuildDeviceStatusPill` - the tab's one informational look - its
  condition in stars, `ConditionStars`), its
  repairs on ANY board - headed "Workbooks associated to this device" (owner request, 2026-11-06), the history export too - totalled
  by ONE `WorkbookSummary` over every worklog, time and cost by the workbook headline's own `BuildTimeAndCostStats`
  (`OwnedDeviceRepairs.Summarize`/`HeadlineStats`), then a headline per part (owner request, 2026-11-06: '"Photos" should
  sit just above the photos'): its links of interest and comments (a link clicked opened through `ExternalTargetLauncher`),
  and its photos, EACH IN A PANEL OF ITS OWN WITH ITS COMMENT UNDER IT ("I want each photo to have its own panel and with its
  text displayed below"), a photo clicked opened in the worklog's viewer. **THE PAGE ONLY SHOWS** (owner request, 2026-11-07:
  "it should not allow for direct edits or 'Add X', as this should be a view-only place"): everything is changed through
  "Edit device" - no notes box, no "Add photos", no "New repair" - **and a part with nothing in it has no headline**
  ("Only show headlines if it has data"; `DeviceLinksSection`, `DeviceCommentsSection`, `DevicePhotosSection`); with no
  workbook it says so (`OwnedDeviceWording.NoRepairsYet`). **The device's NOTES ARE GONE, the description is the one text**
  ("Remove the 'Notes' as we now have the 'Description'"): notes an earlier pre-release stored are read into the end of the
  description (`OwnedDeviceStore.FoldNotes`, `notes` written no more - `Notes` null, `WhenWritingNull`), so nothing typed is
  lost. **No "Changes made to it" any more** (owner request,
  2026-11-06: "this is exactly what the workbooks will cover") - deleted with its recording and its part of the export; a
  `changes` array an earlier build wrote stays in the device's index.json untouched (`JsonExtensionData`). **Its
  description is several lines** (Enter a new line, Ctrl+Enter saves) and keeps them on the page and in the export -
  "Where it came from" until 2026-11-06 (owner request: 'change text "Where did it come from" to "Description"'), STORED
  as `origin` still, its property `Description`. **The form's Region shows only for a board with regions**
  (`BoardRegions.Has`, the board read as viewed off the UI thread; a region stored on a board without them is kept, not
  shown). **"Unknown" is the first status and the first region, "Not mine any more" no status** (owner request, 2026-11-06:
  "this is different data as I see it"; "Not set" the region's none until then): a new device's status is "Unknown", a
  blank one reads so, and "Not mine any more" an earlier pre-release stored stays as any word this CRT does not know.
  **Its condition is 0 to 5 stars** (owner request, 2026-11-06; `condition`, `OwnedDeviceCondition` - a star clicked sets
  that many, the one set clicked again none, the stars lit up to the one under the pointer; the export says it in words, a PDF
  font having no star). **The stars are SHAPES, never a font's glyph** (owner request: 'have this show prettier ... more
  "depth" and coloring'; `ConditionStarGeometry`, `ConditionStars`): a lit one gold, light at its top and deeper at its foot,
  a darker edge and a soft shadow; an unlit one a quiet outline - colors `Workbooks_Star_*` in both themes.
  **Its links of interest are a worklog's** (owner request, 2026-11-06: 'the same "Links of interest" for the device';
  `links`, `OwnedDeviceStore.AddLink`/`UpdateLink`/`RemoveLink`, the worklog's `WorklogAddLinkWindow` and `WorklogLinkRow`),
  saved as the comments are, and in the exported history (`ComposeLinkRows`, the worklog export's own).
  **Its photos are dragged into another order** (owner request: "rearrange the photos, just like it is possible to do in the
  worklogs"; `ListRowDrag` over the form's ONE `ObservableCollection` of `WorklogAttachmentRow`s - an `IDraggableRow` since -
  the order `OwnedDeviceStore.SetPhotoOrder`, at once on a device). **A device with workbooks keeps its hardware and board**
  (owner decision, asked: "Locked while it has workbooks" - "I should only be allowed to associate that workbook to a device of
  the same hardware/board"): "Edit device" greys the two boxes out and says why (`OwnedDeviceWording.BoardLocked`), and
  `OwnedDeviceStore.Update` takes no board for such a device whatever it is handed. "Associate workbook to device" offers the
  devices of the workbook's board alone (`OwnedDeviceRepairs.ChoicesFor`); a workbook already on a device of another board -
  from before - keeps showing it until changed.
  **THE DEVICE FORM IS LAID OUT AS THE WORKLOG'S EDITOR** (owner request, 2026-11-06: "more like the 'Worklog' modal, where
  it has these same possibilities ... Comments - Photos"; `OwnedDeviceWindow`): fields and comments left, photos right,
  the worklog's own small windows (`WorklogAddCommentWindow`, `WorklogAddPhotoWindow`) and rows (`WorklogCommentRow`,
  `WorklogAttachmentRow`). **An existing device's comments and photos are saved as they change** (owner decision, asked:
  "At once, as a worklog does"; `OwnedDeviceStore.AddComment`/`UpdateComment`/`RemoveComment`, `AddPhotos` with a comment
  each, `UpdatePhoto`) - Cancel takes back the fields alone, and `ChangedOnDisk` has the page drawn again; a new device's
  are held by the form and written with it (`Create(details, comments)`, then its photos copied in - one that could not be
  is said on its page). **A photo clicked - the
  worklog editor's, the form's, the page's - opens `WorklogPhotoViewerWindow`, which ZOOMS** (owner request, 2026-11-06:
  "zoom in/out with scrollwheel"; `PhotoZoomView`: the wheel at the pointer, a drag to move, a double-click whole, every
  position `PicturePreviewGeometry`'s, the source rectangle in the photo's pixels). **"Add device"** - the worklog bar's -
  adds one on any board, no workbook needed, and PICKS it; with none, the list's line names that button (2026-11-06: the
  owner had not found it). **The worklog bar has "Add device"
  too, its last button, after "Add worklog" and "Add workbook"** ("Create new workbook" until 2026-11-06; owner requests that
  day: with no device and no workbook the bar offered only that, and "show the 'Add device' as the last button";
  `Main.AddDeviceFromBarAsync` - the tab's own `AddDeviceAsync`, the device picked, then the tab shown on its page). **The device form
  opens over CRT's window whichever tab is shown** (`TabWorkbooks.DeviceFormOwner`; owner report, 2026-11-06: "clicking the
  'Add device' in the top-bar result in nothing"): a tab not shown is out of the window's visual tree, so
  `TopLevel.GetTopLevel` of it is null - a form a tab opens for a caller elsewhere asks for `MainWindow` then.
  **Photos are copied in, and a device deleted, off the UI thread under
  the BusyOverlay**, each under a PLAIN name directly in the device's folder
  (`ComponentPaneFileNames.ForDevicePhoto` and a containment check - Security, code review 2026-10-28: on
  Linux a "\" in a picked name took the copy out of the folder), what could not be added said
  (`OwnedDeviceWording.PhotosAdded`); each photo decoded small ONCE (kept by path, length and write
  time), every workbook read once per refresh - handed over by `Main.RefreshWorklogBar`, which has read
  them - and the board pane not rebuilt behind the page (`ShouldRebuildBoardPreviews`). Enter in "Add
  device" saves from a box, never from a focused button. An action on the page
  reads the devices AGAIN before drawing it (`ShowDeviceAgain` - drawn from the old list, a change
  added did not show). "Add workbook" ALWAYS asks the device when creating (the plan:
  "one of yours on this board, a new one, or none"): "None" and the devices with the board, and an
  "Add device" BUTTON beside them - never an item of the list (code review, 2026-10-29: reaching "A new
  device" with the Down key opened the dialog) - whose form keeps THIS board (`boardFixed`) and picks
  what it adds; one on another board is not offered. Its title follows the device picked ("New repair
  on <name>"). **A device added there is on disk even when the form is then cancelled**: the form is
  opened ONLY through `Main.ShowCreateWorkbookAsync` (the bar's "Add workbook"), which shows it whatever the form closes with (`AfterCreateWorkbookForm`) - no caller to
  forget it. Under a workbook's title "Associate workbook to device" ("On my device" until 2026-11-06, the
  form's too - owner request) puts it on a device with its
  board, or the one it is on. "Delete device" asks, and its repairs stay. Words: `OwnedDeviceWording`.
  **THE WORKLOG BAR IS TWO DROP-DOWNS, "Device" THEN "Workbook", AND ITS BUTTONS** (owner requests, 2026-11-06: "a device
  selector as the first drop-down, and then it will show all workbooks for this device in the other 'Workbook' selector",
  "a badge shown as the last in each line", "remove the highlighted text in the top-bar" - the workbook's counts and dates
  and its Open pill beside the picker went - and "Create new workbook" was cut off at the window's edge; slice 4c's picker
  stood above Hardware until then). The buttons are docked right and always whole. **The drop-downs are as wide as the line
  each shows** (owner request: "have the two drop-downs resize, so it will show as much as the text in the drop-downs as
  possible" - fixed shares until then; `BarPickersPanel`, the rule `BarPickersGeometry`): short of room, shared as water
  fills - a short line stays whole beside a long one cut short - and clipped to their place on a narrow window.
  **Each line ends in its badge** - a device's status (`BuildDeviceStatusPill`), a workbook's Open or Closed
  (`BuildStatePill`) - in the open list and the closed box alike, **and the badge gives way before the words do**
  (`BarPickerLinePanel`, the rule `BarPickerLineGeometry`: the closed device box read "Working" and not "C64" at 1100px).
  **With "Show workbooks for all devices" picked, a workbook's line STARTS with its device's name in a badge** (owner
  request, 2026-11-07: "It should not show this when I have selected a specific device"; `Main.DeviceBadgeOnBar`,
  `WorklogInfoPillBuilder.BuildDeviceNamePill`), none for a workbook on no device; short of room it takes at most
  `BarPickerLineGeometry.LeadingShare` of the line, the words before the badges at the end.
  **THE DEVICE DROP-DOWN IS THE DEVICE PICKED, REMEMBERED** (owner request: "rename 'No device' to 'Show workbooks for all devices'
  and then have that as first selection. When this is selected (and it is per default if never set before) then it will
  show all workbooks"; `UserSettings.WorklogBarDeviceId`, null for all): its lines are `BarDeviceChoice`s, "Show workbooks
  for all devices" first, then every device, **each with how many workbooks it holds, an outlined badge last on the line,
  after the status** ("a counter for number of workbooks", then "I actually did mean an outlined badge, and please show
  that as the last, right of the status"; `WorklogInfoPillBuilder.BuildCountPill`). **A drop-down with nothing ever made for it is not there, its label with it**
  (owner request, 2026-11-08: "if no devices or workbooks exists yet ... then its drop-down boxes should not show either";
  greyed out until then; `FillBarPickers`) - the device one with no device, the workbook one with no workbook on any board -
  and the other has the room. All devices lists EVERY workbook, a device its
  own (`OwnedDeviceRepairs.Of`), on any board. **IT IS ONE CHOICE WITH THE DEVICE MARKED IN "My devices"** (owner request,
  2026-11-06: a device selected there "will only show the workbooks for this device in the 2 workbooks lists", and a
  workbook selected there "I want that device selected in both the drop-down and in the 'My devices' panel"): a card clicked
  picks it (`TabWorkbooks.SelectDevice` -> `Main.PickDevice`; the device picked clicked again only shows its page - a workbook
  of it clicked since is kept, never swapped for the repair a pick works on), a workbook card clicked picks the device it is on - or all
  devices for one on none (`TabWorkbooks.SelectWorkbook`) - and "Clear" picks all (`Main.ClearDevicePick`); the tab's list
  is then the device's workbooks, on any board, a card naming its board only when it is not the one shown. **WHILE A DEVICE
  IS PICKED CRT WORKS ONLY IN ITS WORKBOOKS** (owner decision, asked: "Only the device's" - the drop-down went back to all
  devices until then, never hiding the workbook "Add worklog" wrote into): the active workbook is resolved among them
  (`OwnedDeviceRepairs.WorkedIn`, in `RefreshWorklogBar` and `Main.ResolveActiveWorkbookForBoard` - "Add worklog" and the
  popup's capture alike), so with none of them on the board shown there is none: the workbook drop-down says so
  (`OwnedDeviceWording.NoWorkbookOfDeviceYet`/`NoWorkbookOfDeviceOnBoard`) and "Add worklog" is hidden - never another
  device's workbook written into unseen. **The pick is checked FIRST in `RefreshWorklogBar`** (`OwnedDeviceRepairs.BarFilter`,
  saved back): a device gone is all devices, and so is another board shown that is neither its own nor holds one of its
  workbooks (a device whose board CRT cannot show stays - there is nowhere to go); with the worklog off nothing is read and
  the pick is kept. **With NO board shown the pick is kept too** (owner report, 2026-11-06: another device clicked "does the
  exact same thing as 'Clear'"): switching HARDWARE replaces the boards' list before the board is picked, the bar is drawn in
  that moment, and a device on its way to its own board on another hardware was let go of - `MainDeviceListTests` holds it
  through the real drop-downs. **A workbook activated off the device picked takes the pick along** (`Main.ActivateWorkbook`/
  `ActivateWorkbookAcrossBoards` -> `FollowWorkbookInDevicePick`, the rule `OwnedDeviceRepairs.PickFollowing`): its own
  device, or all for one on none - a workbook "Add workbook" made on another device is the one worked in; **and so does the
  workbook shown, put on another device or none under its title** (`TabWorkbooks.PutShownWorkbookOnDevice`; code review,
  2026-11-08: it dropped out of both lists, another workbook worked in). The workbook
  drop-down says "Select a workbook" in its place with none on the board shown and all devices picked; with no workbook
  at all it is not there (above). **A device picked** works on the repair `OwnedDeviceRepairs.RepairToWorkOn` names (newest OPEN,
  any board; else newest on its board now - only on a board CRT can show, `Main.CanShowBoard`), its board selected first,
  and is remembered; **one with no repair to work on has its own board shown (`SwitchToBoard`, a board CRT lists) and its
  page opened on the "Workbooks" tab** (owner request: "it should not open the 'create new workbook' modal, but instead it
  should go to the 'My Devices' tab and select that device"), and stays picked. All devices picked changes no workbook.
  **"Clear" under the tab's workbook list leaves NO WORKBOOK WORKED IN on the board** (owner request, 2026-11-06: 'a similar
  button named "Clear" ... that will clear the selection of the workbook'; asked, "No workbook worked in"; `ClearWorkbookButton`,
  `Main.ClearWorkbookPick`): `WorklogManager.NoWorkbookChosen` (0 - ids start at 1) saved as the board's active workbook, which
  `ResolveActiveWorkbook` reads as none - no card marked, the bar's drop-down "No workbook picked" (`Main.NoWorkbookPickedText`),
  "Add worklog" hidden - till a workbook is clicked, picked, activated or made. CRT 3.x, reading the same settings, finds no
  workbook 0 and takes its newest. In its place the device picked's page, or "No workbook selected".
  **A board key becomes a drop-down selection only through `Main.FindListedEntryForBoardKey`** - a
  board the drop-downs LIST, not unticked in the Configuration tab's list (code review, 2026-10-29:
  selecting a hidden one blanked them, the bar's own workbook picker too) - and a device's board
  switched to must be the one SELECTED afterwards, or no repair is started (it went on the board still
  shown). **A pick is made from the OPEN list**: closed, the arrow keys and the wheel do nothing (each
  step was a board load, and a dialog could open mid-scroll); the device already shown chosen again
  (an item clicked, Enter or Space) is a pick too - no selection change says so
  (`OnDevicePickerPropertyChanged`). **The click is counted on the TUNNEL route**: the ComboBox closes
  its list in its own bubble handler, so a bubble handler saw it closed and never counted a click
  (second review, 2026-10-29). A selection change is the pick, never also the close after it. A pick
  whose board switch lands nowhere (the board in other capitals) draws the bar again on the workbook it
  works in.
  A pick and the forms it opens use what the last `RefreshWorklogBar` read (`thisWorkbooksLastRead`,
  `thisDevicesLastRead`), never a second read of every folder. The bar's workbook picker reads "#3 | No video", one on
  another board than the one shown "#3 | No video (C64/250407)" (`Main.WorkbookPickerText` - the device is the picker beside
  it), and its "Add workbook" preselects the device picked, else the bar's workbook's, when it has the board shown. `RefreshWorklogBar` reads the devices ONCE (none with
  the worklog off) for the drop-downs and the Workbooks tab.
  **"Export history"** (slice 4d, `TabWorkbooks.DeviceHistory.cs`; content `DeviceHistoryExportModel`,
  drawing `WorkbookPdfExporter.History.cs`): two buttons on the device's page, PDF or ZIP, as a
  workbook's export. The device heads every page; first what it is, the summary of every repair
  (`OwnedDeviceRepairs.HeadlineStats` over the workbook's own panel), the list of repairs, its changes
  and photos; then every repair OLDEST FIRST, from a fresh page, drawn by the workbook export's own
  `ComposeHeader`/`ComposeContent` - never a second drawing of a repair. **Each repair is on ITS
  board's pictures**: the board shown's as CRT holds them, any other's read from its board file as
  viewed (`DraftBoardSource` + `BoardDataReader.LoadAsync`, off the UI thread - never
  `DataManager.LoadBoardDataAsync`, whose statics Main reads after its own load). The ZIP: the PDF,
  "device photos/", and each repair's files under "workbook_<id>/worklog_<id>/".
- **A board's pictures are found by ONE rule, `BoardPicturePaths`** (the Schematics tab's thumbnails,
  the Workbooks tab's board pane, a workbook's export, a device's history): in CRT 4.0 a picture row
  NAMES a file of its board, its bytes stored under their checksum (`DraftFileResolver`), else an
  Excel-era path under the data root. The Workbooks tab joined the name to the data root by itself
  until 2026-10-29 - on 4.0 data its pane showed no board and every export went out without one.
  **The Workbooks tab takes Main's map for the board it LOADED** (`Main.CurrentBoardPicturePaths`,
  `LoadedBoardPictures`): worked out FIRST in the load's pool pass (`BoardPicturePaths.Place` - each
  picture resolved once, the thumbnails taken from the same list; guarded, so a later failure in that
  pass cannot take the tab's pictures with it), handed over BEFORE the thumbnails decode (the pane
  redrawn with it, with the worklog on), cleared as a load begins, and answering only for the board
  it was made for (`PathsFor`). Worked out again from the drop-downs at draw time, a refresh between a
  board change and its load resolved the old board's names against the new board's file. Null while
  not worked out: the pane says "Finding the board's pictures...", and worklogs whose pictures none
  could be drawn are SAID, never a blank pane. **Both exports take a board's pictures one way**
  (`TabWorkbooks.PicturesOfBoardAsync`): Main's map for the board shown once there, else read from the
  board's file as viewed (`DraftBoardSource.ResolveViewed` - board file, draft folder and new-board
  from ONE resolution; `BoardFileListings.ReadDocument` - rows and file list from ONE read). Both open
  one save dialog (`TabWorkbooks.PickExportPathAsync`).

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
| `TabSchematics.Thumbnails.cs` / `.ThumbnailsDetach.cs` | Thumbnail loading, order and selection, and the detached window mode. The panel itself - in the tab, the detached window and beside fullscreen - is ONE control, `SchematicsThumbnailGallery` (auto-fit, never scrolled; owner request 2026-10-15). **The selection and the label editor's lock are the TAB's** (`SelectedThumbnail` - a field, never a list bound to the tiles, which a drag's lift emptied - `SelectThumbnailFromGallery`, `IsThumbnailNavigationLocked`, `CancelThumbnailDrags` before the tiles are replaced): a gallery shows them, listens only while in the visual tree, handles EVERY press and release on a tile (the ListBoxItem selected on a press, on a right press, and on another finger's release), selects on its own pointer's left release, and never passes on what its own ListBox does by itself - each "nothing selected" from a lifted tile or a Ctrl+click once blanked the view. **Emptying the tiles tells the tab first, its list last** (owner report, 2026-11-03): a list set in between read the tile it let go of from the empty collection and threw, and Main's load died there - no pictures, no message, "not a picture of the board shown" - so a gallery shows a selection lost that way once its list has caught up (`IsThumbnailsChangeUnderway`) |
| `TabSchematics.Highlights.cs` | Component highlight overlays, blink, hover UI, labels |
| `TabSchematics.LabelEditor*.cs` | Label editor lifecycle, interaction, snap context, test seams, and the components not marked on the picture (`.Unmarked.cs`) |
| `TabSchematics.Worklog.cs` / `.Worklog.TestSeams.cs` | Worklog overlay, area drawing and moving |
| `TabSchematics.KiCad*.cs` | KiCad load, panels, rendering, render cache, geometry, hit-testing, calibration |
| `TabSchematics.Settings.cs` | Board-level and global setting rows |
| `TabSchematics.PictureTools.cs` | "Mark components"/"Align KiCad" asked for from the Contribute tab, and their bar (Done/Cancel) |
| `TabSchematics.ViewTools.cs` | Panning with the right OR middle button (`TryBeginPan`, every mode) and the one-time tip saying so; "Draw trace" |

**The view's small tools** (CRT 4.0 step 5's slice 5a, 2026-10-29): **traces are drawn and changed
ONLY while "Draw trace" is on** (the plan: "Traces are drawn only with Draw trace") - off, a press on
the picture never starts, bends or deletes a trace (a right-click on one pans, a left one selects),
and `PolylineManagement` is not consulted at all; Escape stops it at WINDOW level
(`Main.OnMainKeyDownCloseSinglePopup` - the component filter usually holds the focus) **but only while
the Schematics tab is shown** - elsewhere Escape is the box's it was pressed in (code review,
2026-10-29) - and so do leaving the tab, another tool, another board and no picture (also one with no
file). **Another board or no picture DROPS a point being dragged, never saves it**
(`SetDrawingTraces(false, navigating: true)` -> `StopEditing(saveDrag: false)`): saved, the old
picture's traces went under the new board's key. **A picture's traces are its own** (third review,
2026-10-29): a picture switch lets the last ones go AT ONCE (`LetGoOfTraces` - before the decode, which
left them editable under the new name), they are read with the picture (`ReadTracesFor`), and they are
changed or SAVED only while they are the chosen picture's (`TracesAreThePictures` - the save handler and
`CanEditTraces` in the input; none chosen or none read is never a match). **While another tool is on** (the label editor, a
worklog area, aligning KiCad) **"Draw trace" is greyed out, its tooltip saying why**
(`FollowToolsInDrawTrace`, from `FollowPictureTool` and the worklog area's begin and cancel) - and a press
on it greyed out is not the picture's (`IsPointerOnDrawTraceButton`: a disabled button takes no pointer,
and the label editor lost its selection); "Done
drawing" is PLAIN - it turns a mode off and saves nothing (second review, 2026-10-29). **A "Has KiCad data" badge stands beside
"Draw trace" while the schematic image shown has the board's KiCad traces on it** (owner request, 2026-11-12: "a discrete
badge ... This info should show per schematic image"; `KiCadDataBadge`, set with the KiCad settings' rows in
`UpdateInteractiveCadTraceHoverModeUi`) - the "Schematics" tour's trace step frames it, so that step shows only where it does.
**"Loading KiCad data..." takes its place while the board's KiCad project is read** (owner request, the same day; never both -
`ShowKiCadBadges`), **and the row - the badge and "Draw trace" - stands LEFT of the "Traces visible" panel, never over it**
(`SchematicsViewToolsRow`). **"Draw trace" stays solid under the pointer, with a hand** (owner request: the theme's hover fill let
the schematic image show through - its own style in the markup). **A middle drag pans as the right one does**, and only a
right click that did not move is a click (a component's highlight). **The one-time tip** sits under the
picture's name until "Got it" or the first pan (`UserSettings.SchematicsPanHintSeen`). **"Global
settings" is a panel as the two beside it** (owner request, 2026-11-05: "use that name again instead of the gear icon" - a gear
button's popover from step 5 until then), its "Collapse" remembered (`schematicsGlobalSettingsPanelExpanded`), shown only
while a row is - every row is a KiCad one. **The three panels - "Global settings", "Labels visible", "Important KiCad signals" -
collapse and expand by a left press on their headline and the room around it too** (owner request, 2026-11-06; `OnPanelHeaderPressed`,
the header's border taking the panel's padding back); a press on the button is the button's, never a second toggle.

#### Maintainer (`Tabs/Maintainer/`, logic in `Handlers/Maintainer/`)

**"Maintainer" is ONLY the role** - a person in a board's pool; the person who owns this project is
"the project owner" in every document and comment. The ACTIVITY is "review" (`ReviewEndpoints`,
`/api/review/...`, `ReviewSummary` keep their names). The tab was a separate application until
2026-09-29.

- **Shown only when "Enable Maintainer tab" is ticked** (`UserSettings.EnableMaintainerTab`, default
  off), between Contribute and Configuration. `Main.Maintainer.cs` owns visibility and layout: while
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
- **The tab's badge** (`MaintainerModes.TabAttention`): the two queues' attention counts ADDED UP - and,
  for the administrator, the oscilloscope models waiting (the queue answer's `OscilloscopeModelsWaiting`,
  badged on the Management tab and on its "Oscilloscope models" entry too - owner request, 2026-11-07: "indicate
  which menu item to choose in the 'Account'"; 2026-10-29), and the models screen's own read sets that count at
  once (`OscilloscopeModelsView.ModelsWaitingRead`). One rule: `UpdateModeBadges` calls
  `MaintainerModes.TabAttention`. Which screen the tab OPENS on is the two queues' alone
  (`QueueAttention`). No tooltips on it. A 401 clears what it counts.
- **Another submission chosen takes the decision bar and the removals with it, until its own detail
  arrives** (code review, 2026-11-02; `ShowSubmission`), and `DecideAsync` refuses without a detail of the
  submission selected: the removal list, the version and the approval state all come from it.
- **A 403 on the queue is an answer, not "could not ask"** (code review, 2026-11-02;
  `EmptyListsNoLongerAllowed`): the queue and BETA lists empty and the badges count nothing, unless a
  table holds unsaved changes. **Every read drops an answer for a session no longer in use**
  (`IsSessionInUse`). **The BETA screen's plan is read again when the maintainer asks** - the screen's
  button or coming back to the tab - when none is held, it carries a refusal, or a publish was just
  refused (`ProductionDisplay.PlanIsWorthReadingAgain`); never one that can be published, or its tick would
  go, nor for the administrators-only refusal alone.
- **The minute check** (`QueueRefreshRules.MinuteCheck`): everything while on screen with CRT in
  front; otherwise only the queue and the BETA list while the badge can be seen. Every request slides
  the 30-day session (accepted with the request). Nothing open is reloaded unless its row changed.
- **Five screens as a tab strip** - Boards, "Queue: Contributor submissions", "Queue: Awaiting push
  from BETA to stable", Management (the administrator's alone), Account. **Their labels are CRT.Data's `MaintainerScreenWording`** - never write
  a screen's name as a literal. Switching screen HIDES, never closes. With nothing waiting, the tab
  opens on Boards (`MaintainerModes.ScreenOnOpening`); the queues open on the last-looked-at entry
  (`MaintainerModes.EntryToOpen`).
- **Account**: every maintainer has "My account" (name, address, password, Sign out -
  `MyAccountView`; the session alone is enough, no current password - an accepted risk) and "Server
  version", and nothing else.
- **Management** (owner request, 2026-11-07: 'a new tab named "Management" (shown before "Account") - it should only be
  visible by me ... show a "lock" icon. Then move all the menu items from "Account" into there'; `TabMaintainer.Account.cs`,
  `ShowManagementScreen`): the administrator's entries - Maintainers, Order of boards, Unused files, Rebuild checksum
  manifests, Delete a board, API usage, Oscilloscope models, Reset contribution data - its TAB padlocked and shown only while
  the server says the account is one; the entries carry none. Signed out, or another account, the tab goes and a screen on it
  goes to Account (`ShowModeAsync` sends a non-administrator there). It opens on "Oscilloscope models" while one waits, else
  on "Maintainers" (`EnterManagementAsync`). The server's sentences name it through `MaintainerScreenWording.ManagementQuoted`.
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
  (`MasterListing.NamesTakenBy`). **The way there is one press** (owner report, 2026-11-05: "where can I set this?"):
  an amber panel over the board's views with "Place it in the drop-down lists" (`BoardSections.PlacementNotice`, gone while
  the Maintainer view holding the place is shown), and the same button on the queue's "why Approve is off" panel
  (`TabMaintainer.GoToPlacementAsync`) - neither button for an account that may not place it.
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
  a new "why this cannot be done now" panel uses those classes, never its own colours. **The words are
  NOT bold and the icon stands in the MIDDLE of them** (owner request, 2026-10-30, of the Contribute
  tab's local-draft notice - the same look everywhere; a person named keeps the bold of its own run). The read-only
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
  **A file cell's place is looked up ignoring capitals** - JSON hands `FilePlaces` over without the
  server's comparer, so the source copies it into one that ignores them (code review, 2026-10-26).
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
  Contribute tab's, in one continuation. The tab has no `BusyOverlay` of its own. The component filter
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
`UserSettings`, `Logger` (and `ProblemsLog`, below), `WorklogManager`,
`PublishedDraftRetirer`. The KiCad loaders are CRT.Data's since 2026-11-06 (the conversion makes pinouts with
them - "THE CONVERSION MAKES EVERY BOARD'S PINOUTS" below). Check which side a class belongs on before moving it - a hidden
dependency makes a "move" a rewrite. `HighlightRectBuilder`, `PolygonGeometry` and `RectGeometry` stay
in CRT.App (Avalonia types, 21 call sites).

- **`BoardWorkbookSchema.AllSheets` is the order of every workbook written, ending in "Credits"** (the
  table's tabs have their own: `BoardTableTabs`). Readers find sheets by name.
- **`BoardDataReader`'s cache answers only while the workbook AND its JSON sidecar are as they were
  read** (length + write time, `BoardDataReader.StampOf`; owner report 2026-10-11) - a write by
  anything (CRT, Excel, the sync) is a new read. The `ClearBoardCache` calls stay, but nothing may
  depend on one being made: after a draft write only the SELECTED board's cache was cleared.
- **Every workbook cell is read through `WorkbookCellText`** (code review, 2026-10-13): EPPlus's `.Text`
  formats a NUMBER with the current culture, and the shipped workbooks hold opacities as numbers - a
  Danish CRT read "0,3" and wrote it back. Never `sheet.Cells[..].Text` in a reader.
- **A cache keyed by length + write time trusts the stamp only once SETTLED** (`StampSettling`, 2 s):
  two same-length writes inside one clock tick share a stamp (git's "racily clean").
- **A component relabelled alone (`ComponentRelabel`) takes its highlights along** in every draft write
  and in a maintainer's amendment, by `BoardDataDiffer`'s pairing rule - like a renamed schematic.
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

### The CRT 4.0 data tree (CRT reads it since 2026-10-24; the server publishes it since 2026-10-25)

Owner decisions of 2026-10-24: in CRT 4.0 board data is JSON only and every edit is made in CRT;
the Excel generations are FROZEN at the 4.0 release (2.x/3.x keep their data, and are told to update
to contribute); every board OWNS its files by name, while each file's bytes are stored ONCE under
their checksum; all existing data is migrated. **The app switched on 2026-10-24** ("Yes, do switch
now"): CRT reads and writes only 4.0 boards and drafts. **The server followed on 2026-10-25**
(server 4.0.0, API revision 2 - every 3.x CRT is told to update to contribute): it publishes ONLY
into 4.0 trees, so it is deployed on the release day, not before (INSTALLING.md, "Switching the
service to the 4.0 trees"). CRT reads a main file only at or below its own version, so running it on
4.0 data needs `InformationalVersion` 4.0.0-something.

- **Layout** (`DataTreeLayout`): a tree of its own (`public_html/app-data-v4`, `app-data-v4-BETA` -
  an installed 3.x downloads every file its manifest lists, so 4.0's files never go beside the Excel
  generations). `Data/Classic-Repair-Toolbox.v4.0.0.json` (`MainDocument`), `Data/<board id>/board.json`
  (`BoardDocument`: the nine sheets, the sidecar's three roots, the board's `files` by name + SHA-256),
  `Data/Files/<ab>/<sha256>.<ext>`, `Data/MiniPro/IC tests/`.
- **A board's identity is its board file, "<id>/board.json"** (`DataTreeLayout.BoardFilePath` - the
  main file's entries, a converted draft's marker, `NewBoardIdentity.BuildExcelDataFile`). The name
  `ExcelDataFile` stays in code for now. **Only the folder of an identity matters**: a 3.x draft whose
  identity names a workbook ("…/Data X.xlsx") reads the same `<folder>/board.json`
  (`DataTreeLayout.BoardFileFor`), so every identity of one folder is one board - the draft-only merge
  skips a draft whose FOLDER is listed (`DataManager.MergeDraftOnlyBoards`). The shared readers read
  a `board.json` as a 4.0 board and anything else as Excel (`DataTreeLayout.IsBoardFile`) - the
  converter and the server still read Excel through them.
- **CRT's data folder is "Data v4"** (`AppConfig.DataFolderName`), beside 3.x's "Data", which is left
  as it was (a new folder is seeded from the install package; `DataManager.FindLegacyDataRoot` lets a
  hand-placed Excel-era draft find 3.x's shared files while it is converted). The manifest is
  `app-data-v4`'s. The launch syncs the main file, then the listed board files; everything else (the
  store, the MiniPro tests) syncs in the background (`DataManager.IsListFile`). The orphan cleanup maps
  the main file, its boards and the stored files they name, and **the stored files any draft lists**
  (2026-11-02: a draft keeps a published file by checksum and holds none of its bytes) - and fails closed
  as before, on a draft's board file too.
- **A File cell is joined to a root only when shaped as a relative path** (code review, 2026-11-02:
  `Path.Combine` lets a rooted or UNC cell replace the root, and `File.Exists` on a UNC path can leak the
  NTLM hash). Every fallback is `DraftFileResolver.PathUnderDataRoot`, never `Path.Combine(dataRoot, cell)` -
  so a test's picture row names a file under a data root it set, never an absolute path. **A draft's identity
  is held to the same shape** (`DraftFolderLayout.GetBoardFolder`, empty when `SubmissionPathRules.IsSafelyShaped`
  refuses its folders - the one place every draft path passes): a marker in a folder somebody handed over is
  read at every launch, so it reached `Directory.Exists` with nothing pressed. `DraftManager.GetBoardFolder`
  is that one rule, and a marker's board naming no folder here is not listed. **A name the
  session has seen listed is not free for another file** (`DraftSession.FileNamesInUse`: its history holds
  rows, and a row names its file). **"Bring my draft up to date" keeps a file a kept row cites**, its bytes
  copied into the draft's store (`DraftFiles.KeepInStore`); gone already, the conflict list says so. **A
  rename stopped between its move and its marker is repaired at launch** (`DraftBoardRename.RepairMovedDrafts`),
  unless the folder would collide with another; one stopped between a change of capitals' two moves goes back to where its
  marker says (2026-11-03).
- **A row names a file of its board; `DraftFileResolver` finds it** - a new file lying in the draft
  folder under its name, then the draft's board file (its store, then the tree's), then the published
  board file (pass the board identity), then an Excel-era data-root path. Every caller that opens what
  it finds passes the ROOT it was found under to `ExternalTargetLauncher` (the Resources tab too).
- **KiCad files are copied under their names** into a cache (`BoardKiCadFolder`,
  `AppConfig.KiCadCacheFolder`) - a project's sheets name each other. A board shown with a draft reads
  only the draft's (the draft lists the published ones too). The "KiCad data" window's "imported"
  files are the draft's own (`BoardKiCadFolder.ImportedNames`); Remove takes a name off the draft's
  list (`DraftWorkbookStore.RemoveFiles`) - **or, for a name the published board has, puts the
  published file back** (code review, 2026-10-25: dropped, the board lost it at the next submission).
  **A listed name can say anything** (a hand-placed draft's "KiCad data/../../x"): Remove deletes a
  copy only where `DraftFiles.TryArrivedPath` finds it, and of the store only the removed entries'
  bytes - never the rest, which the session may still cite (code review, 2026-10-26).
  **The copy is NEVER made on the UI thread** (code review,
  2026-10-25: the first showing of a large board froze the window, and every KiCad render and hover
  asked again): `Main.KiCadRawPathsFor` runs on the pool while the thumbnails decode, and the
  Schematics tab keeps the paths its load used (`thisCurrentKiCadRawPaths`) - never ask for them
  per render; the KiCad data window reads its list and its report from one read on the pool
  (`KiCadDataPanel.ReadKiCadAsync`). **A replaced copy gets a LATER write time than the one it
  replaced** (`BoardKiCadFolder.Copy`): `KiCadProjectLoader` keeps its parse by path and write time,
  File.Copy keeps the source's, and two sources written in one clock tick served the old parse.
  Only the types the loaders read are copied, the cache is in the user's OWN app data
  folder (never the shared temp folder - the copy deletes what it does not want), and nothing is
  written or deleted through a folder that links elsewhere (`FolderLinks`). **An import is listed by
  a write at once, and Submit carries KiCad files not listed yet too** (`DraftFiles.ArrivedKiCadNames`,
  `SubmissionKiCadFiles.Collect`) - a project imported last was drawn, counted and not sent.
- **The files are written by `DataTreeJson` only**: the same content is always the same bytes (one row
  per line, fixed field order, "\n", no BOM); a field this CRT does not know is KEPT and written back -
  **on every row too**: every writer that rebuilds a board file from BoardData hands
  `BoardDocumentUnknowns.CarryOver` the file as it was (the draft write, the workbook writer, the
  seeder, the publish - which reads it IN FULL or stops); a new writer does the same. A higher
  `formatVersion` is refused. A row cites a file by the board's NAME (`BoardFileNames`), never a tree
  path. The checksum the store names bytes by is `DataTreeLayout.Sha256Of`, the one helper. **A
  highlight's coordinates are whole pixels, a fraction ROUNDED away from zero** (`BoardDocumentMapping
  .WholeNumber`, as the Excel-era writer rounded) - the label editor hands over real geometry, and
  refused, every rectangle drawn or moved in a 4.0 draft failed to save (code review, 2026-10-26); what
  is no number stops the write. **The main file's list is read ONE way** (`MainFileListing.Rows`: names
  trimmed, a blank hardware name the row above's) - CRT's drop-downs and the server's checks alike.
- **"Component images" is split** into pictures and scope baselines by `ComponentImageKinds` - the one
  rule, which the app's `ComponentImageQueries.IsOscilloscopeImage` asks too.
- **The converter fails closed** (`DataTreeConversion`): whatever CRT's forgiving readers would quietly
  drop (an unknown sheet or column, a value under no heading, a sidecar entry the reader skips, a
  missing cited file, an unlisted newest-generation board) is a PROBLEM, and a plan with one is not
  written. It writes only into an empty folder or a tree it wrote, never into or around the source.
  Run as `CRT.Server convert-data-tree` (INSTALLING.md, "Converting a data tree for CRT 4.0").
- **`DataTreeConversionShippedDataTests` is the proof** the release day's conversion rests on: every
  shipped board's board file holds exactly what CRT reads from its workbook today, file by file as the
  same bytes. A new column on a BoardData row type must be carried by `BoardDocumentMapping`
  (`BoardDocumentMappingTests` counts them). `Assets/Data` is still the Excel tree until the release.
- **A 4.0 draft** is its board's folder laid out like the tree: `board.json` (the complete board, the
  published KiCad files included), `Files/` holding only bytes the tree lacks, `.crt-draft.json` naming
  the board by its board file, `.crt-base.json` as a BoardDocument. Seeding copies the published board
  file and NO bytes (`DraftSeeder`). **A newly picked file lies in the draft folder under its name and
  leaves it only once a write has listed it** (`DraftFiles.ListFor` copies it into the store,
  `RemoveArrived` after the write) - moved first, a failed write would orphan it. **Only a name that
  resolves INSIDE the draft folder, through no linked folder, is such a file** (`DraftFiles.TryArrivedPath`,
  asked by `DraftFileResolver` too; code review, 2026-10-25): the file is DELETED from where it lay,
  and a File cell holding a full path, or "../", had the user's own file deleted at the next save. Folders named
  `.crt-*` are CRT's (`DraftFolderLayout.IsInBookkeepingFolder`) - never the contributor's work, never
  a reason to keep a draft. The fingerprint is the board file's CONTENT plus the folder's other files
  ("v2:"). Retirement compares stored bytes with the tree's store and arrived files with the published
  board's file of that name - **stored bytes the draft's board file does not list are no work**, while
  a draft that lost its board file counts them all. **The draft held in memory keeps the bytes of every
  file it has seen listed** (`DraftFileMemory`, code review 2026-10-26): its history holds rows, not the
  list, and a save dropped a removed row's file and deleted its bytes, so Undo brought the row back
  citing nothing; a cited name the list lost takes the remembered entry (never a KiCad file), and the
  next run's first save tidies. A listed name is looked up as `BoardFileNames.Normalise` spells it -
  kept by its own spelling ("KiCad data/./x"), it was never found and went with its bytes at the next
  save. "Bring my draft up to date" merges the file lists three ways
  (`DraftRebaseFiles`): a file the contributor replaced is theirs AND so are the rows citing it.
- **Excel-era drafts are converted, not read**: `DraftConversion` converts one IN PLACE (names agree
  with the published board's - same bytes, or the base's bytes at that path for a replaced file - the
  old draft kept whole in `.crt-pre-4.0/`, an interrupted run resumes, an odd draft converted with
  notes), and `DraftFolderImport` runs it on an Excel-era folder put into Drafts/ by hand (a 4.0
  folder only gets its marker). The base's checksums are UPPER-case hex, the 4.0 files' lower-case.
  Tests make Excel-era drafts with `LegacyDrafts` (CRT.Data.Tests), 3.1.0's seeder.
- **A user's own scope models** (rows added to their copy of the main Excel file) are found by
  `OwnOscilloscopeModels`; a published model they changed is kept as "<model> (my version)".
- **What 3.x left is taken over at every launch** (`LocalDataMigration`, run by
  `DraftManager.ImportHandPlacedFolders` before its import): every Excel-era draft is converted (one
  that cannot be yet is tried again next start); every board a 3.x "_UserContribution" workbook
  listed is copied into Drafts/ ONCE (`.crt-carried-from-3x.json` in the drafts folder - a discarded
  one never comes back), through a `.crt-` folder the import's walk skips, for the import to take in
  as a new board; a board the main file lists, or one with a draft folder, is left. The user's own
  scope models live in `Classic-Repair-Toolbox.own-oscilloscopes.json` beside the settings
  (`OwnOscilloscopeFile`), made ONCE from 3.x's main workbook (even empty) and listed after the
  published models - **never while that workbook cannot be opened** (held by another program, or
  broken: code review, 2026-10-25 - the empty file stopped every later search), then it is looked
  for again at the next start. 3.x's data folder is only read. Settings are keyed by board NAMES and receipts by
  board ID, so neither needs remapping; a remembered schematic image whose file is not found falls
  back to its name (`SchematicRenameLog.Resolver`). A test's `DataManager.LoadFrom` reads neither the
  3.x folder nor the own-models file.
- **The server publishes by NAME** (2026-10-25). A submitted file's `Path` is the board's name for it
  ("Images/sheet1.png"), never a tree path; the server decides where its bytes lie. `PublishPlan`
  stores each file once (`PlannedFile.StoredPath`, "Files/<ab>/<sha256>.<ext>"; bytes the tree already
  holds are `UnchangedFiles`), and `PublishExecutor` writes the stored files, THEN the board file
  (`BoardDocumentMapping.ToDocument`, keeping the published file's unknown fields), THEN the main
  file's row - a board file never names bytes that are not there. `PublishedBoardFiles` reads a
  published board's names; `PublishedTreeProbe.For(root, board)` is a view of ONE board. A submission
  takes bytes the tree stores by CHECKSUM, whatever name they were published under. **An approval reads
  the published board file ONCE** (`PublishedBoardRead`, handed to the plan, the removals, the summary,
  the rename log and the executor; code review, 2026-10-26 - some five parses under the publish lock).
- **Nothing is shared in the 4.0 tree**: a file is a board's own by name, and bytes another board
  lists too are stored once for both. So `TouchesSharedFiles` is always false (an older row's flag is
  cleared at its approval) and no publish needs two approvals; `ApprovalRules` stays for the record.
- **What a publish, a promotion, a push-back or a deletion removes is a STORED file no board of
  that tree lists afterwards** (`DataTreeUsage`'s JSON branch, `AutomaticRemovalScope.IsStoredFile`) -
  bytes another board lists always stay; a board's folder holds only its board file. The list matched
  is of stored paths (`FileRemovalPreview.Files`); people read each by the board's name for it
  (`FileRemovalPreview.Names`, `NameOf`). A replaced file's earlier bytes are removed under the very
  name it keeps - the Maintainer tab counts and words that as the replacement, not as a file nothing
  uses (`FileRemovalWording.Unused`). Every entry fetched from a tree carries where its bytes lie
  (`BoardFileEntry.TreePath`/`PlaceInTree`, `BoardTableAnswer.FilePlaces`, `ProductionPlanAnswer
  .FileNames`) - the public data is laid out by checksum.
- **A picked file is filed under the board's own names** (2026-10-25, the component editor's "File
  location"; since slice 3c the component pane, `ComponentPaneFileNames`): the folder every converted
  board keeps that kind of file in (`ContributionFileLocations.DefaultFolderFor`) - never the data or
  drafts tree's folders, which in 4.0 are the store's checksum folders and CRT's bookkeeping. A name
  under `Files/` or `.crt-` is never listed by a draft write (`DraftFiles`), so a file filed there was
  copied in and lost. The drafts flow's words name no Excel: the lock-file check stays, worded "open in
  another program".
- **EVERY PRE-RELEASE TALKS TO A TEST INSTANCE OF THE SERVICE, EVERY RELEASE TO THE LIVE ONE, BY ITS
  VERSION** (owner decisions, 2026-10-25: "We need the test server"; 2026-11-09: "if the
  <InformationalVersion> is not a 'real' version, then the application will refer to the
  'crt-server-test' ... If it is a real version, then it should refer to the 'crt-server'";
  `ContributionServer.RootUrlFor` - no "-" before any "+" is a release; no version at all is not):
  its own program folder, settings, database (`crt_review_test`) and port (5198), at
  `https://classic-repair-toolbox.dk/crt-test/api/`, with a shipped unit of its own
  (`crt-server-test.service` - every line the live unit's but the five naming the instance,
  `TestInstanceUnitTests`); set up as INSTALLING.md's "A test instance for pre-releases". **Nothing is
  set back by hand on release day** - `AppConfig.ContributionServerRootUrl` is a PROPERTY over the
  version (a static field declared above `AppDisplayVersionString` would read it as null, and every
  release would talk to the test instance; `ContributionServerTests`). **AND EVERY PRE-RELEASE READS THE
  TEST INSTANCE'S DATA** (owner decision, 2026-11-09: "we need to have a new data source for the test
  server, so it does not mix data with the real server"; asked, the pre-releases read "The test
  server's"): the test instance publishes into `app-data-v4-test-BETA` and `app-data-v4-test` - copies of
  the 4.0 trees, each manifest's addresses rewritten (INSTALLING.md, "Setting it up, once" and "Moving it
  to trees of its own") - and a pre-release's manifest is there (`ContributionServer.ChecksumsUrlFor`,
  `AppConfig.LiveDataRootUrl`/`TestDataRootUrl`, properties for the same reason), its BETA source the
  test BETA tree. Its unit can write those two and nothing of the live service's (`TestInstanceUnitTests`).
  So nothing is moved on release day, and the test instance stays after it. The local "Data v4" folder is
  shared, as the BETA source shares it: a release launched after a pre-release syncs it back by its manifest,
  as switching the BETA source does. **Only the contribution service and the data move**: submissions, the
  Maintainer tab, the health check and the data use the test instance's, while the check-in, feedback and
  board views stay on `CrtServerRootUrl`, the live service - the usage numbers and the feedback mail. What CRT remembers of a service is kept
  per service (`ContributionServer.FileNameFor`: `maintainer-session-test.json`, `submissions-test.json`)
  - shared with the files a release reads, a pre-release would send the live token to the test
  instance and forget it on the 401. The sign-in screen and the Submit dialog say it is the test server,
  in words naming no version of CRT.
- **THE SERVER UPDATES ITSELF FROM GITHUB** (owner request, 2026-11-09: "The goal would be for me to not
  manually update this ... This should of course be as secure as possible"; owner decisions the same
  day: proven by "GitHub attestations", and "the REAL version should work on a release ... But ALPHA/BETA
  works on commit", "can it then update every minute?"; server 10.3.0; `src/CRT.Server/Handlers/Updates/`,
  INSTALLING.md "Automatic updates from GitHub"). **Two channels:** the LIVE service ("releases") installs
  the server build a real CRT release carries - the release workflow's `build-server` job; the TEST
  instance ("main") the build of every push to main whose tests passed - the push workflow's
  `server-build-of-main` job, which puts it on the release of its SERVER VERSION in a public repository of
  its own, **`HovKlan-DH/CRT-Server`** (owner request, 2026-11-11: the standing pre-release `crt-server-main`
  in CRT's repository "will polute the list" - gone since server 10.4.2; `.github/workflows/crt-server-publish-main.sh`:
  a new version's release made as a draft, its files put on it, then published as the newest; an existing one's
  files replaced, the pointer last), with the secret `CRT_SERVER_REPO_TOKEN` (a fine-grained token, Contents
  write on CRT-Server alone). **Its release notes are that version's section of
  [src/CRT.Server/RELEASE-NOTES.md](../src/CRT.Server/RELEASE-NOTES.md)** - one-liners in CHANGELOG.md's shape
  (owner request, 2026-11-11: VERSION.md's row as notes was "very hard to read"), with the bump and API revision
  from VERSION.md under them (`crt-server-release-notes.py`, rewritten at every push of the version). **CRT-Server's README is kept in that repository alone** (no copy
  here): it says what the builds are and that nobody needs them. Both build through ONE composite action,
  `.github/actions/crt-server-build` (publish, `crt-server-release.json` inside - server and CRT version,
  commit, pre-release - pack, `actions/attest-build-provenance`, and the pointer
  `crt-server-linux-x64.json` - tag, server version, commit, SHA-256). **On the box** `CRT.Server
  update-server`, run EVERY MINUTE (a minute after the last look ended - `OnUnitInactiveSec=1min`) by
  `crt-server-update.timer` / `crt-server-test-update.timer` as the user `crt-deploy` - not in `crt-data`,
  so it cannot read the settings - reads its channel's POINTER (`ServerReleaseFeed.PointerUrlFor`:
  `releases/latest/download/...` of CRT's repository - a pre-release is never "latest" - or of CRT-Server's, where
  a pointer's tag must equal its server version; the attestation names CRT's repository whichever one the bytes came from;
  release downloads, which GitHub's API limit of 60 an hour without a login does not count) and downloads
  only a build it has not looked at and that says it is newer - a commit leaving the server's version as
  it was costs one small file. **The pointer is not proven and decides only whether to download**; it must
  name its own channel's release (`ServerBuildPointer.Read`), the build must have its checksum (main's files
  are replaced at every push - one replaced mid-download is retried next minute). **The trust is fixed in
  code, never a setting** (`ServerReleaseFeed.SignerWorkflowFor`, `ServerUpdateHost.VerifyArguments`, held by
  `ServerUpdateCommandTests`): `gh attestation verify` with the bundle, offline from GitHub and with no
  login, pinned to `HovKlan-DH/Classic-Repair-Toolbox`, the CHANNEL'S workflow (`build-and-release.yml` for
  live, `build-and-unittest.yml` for test), `refs/heads/main` and a GitHub-hosted runner - **so both jobs
  must stay in those files, on `ubuntu-latest`, and releases are run from main**; the version and
  pre-release read from INSIDE the attested build, never a release's flag. Never backwards, never a
  pre-release's build on the live service, never a build that failed before (`--retry-failed`); a held
  build is retried every minute from disk; a look saying again what the last said writes nothing to the
  journal. Layout: `<instance>/releases/<version>/`, `releases/current` switched by rename,
  `app -> releases/current` so the units stay, settings in `<instance>/settings/` linked into each build,
  `releases/update-status.json` (`ServerUpdateStatus`) which the service reads beside the build it runs and
  answers in `GET /api/review/server-status`'s `updates` (CRT.Data's `ServerUpdatesAnswer`), shown under
  Account > "Server version" (`ServerVersionDisplay`). After the restart `/api/health` must answer the new
  version within two minutes, else the build before goes back. **A build shipping a changed systemd unit,
  timer or the polkit rule (`50-crt-server-update.rules`, restart of the two services only) is HELD** until
  the installed copy matches or `system-files-installed` is put in its folder - a change to any of those
  files is a step for the project owner, said in VERSION.md. **The updater reads the installed copy as
  `crt-deploy`, so a unit is installed `install -m 0644`, never `cp`** (2026-11-09: cp gave it the build's
  0640, and 10.3.1 stayed held with its unit in place); the polkit folder is root's, so a changed rule
  always needs the marker. The held status says the commands (`ServerUpdater.HeldForSystemFiles`). **A folder the
  updater makes gets its mode IN THE MKDIR, never after** (`ServerArchive`; 2026-11-09): Linux takes the setgid bit
  off a folder re-moded by a user outside its group, which `crt-deploy` always is, so 10.3.1's files came in
  `crt-deploy`'s group and the service died on its runtimeconfig (exit 147) - `ServerArchiveTests` holds it by the
  source, as no test user is outside its own group. The decisions are `ServerUpdater`'s against
  `IServerUpdateHost` (tests: a fake host, virtual time); `ServerUpdateHost` is the thin Linux part.
- **Step 1's foundation, `DraftSession`** (2026-10-25, CRT.Data - used by the Contribute tab's table and,
  through `DraftEdits`, by every other draft editor since the same day): a
  draft held once in memory for every editor. `Apply(description, transform, origin)` edits it in
  memory and raises `Changed`; ONE history for every editor (`Undo`/`Redo`, whole-content snapshots,
  `MaxSteps`); it saves itself - the host's timer asks `IsSaveDue` (a second after the last change)
  and calls `Save` off the UI thread; "saved" is by content, so undoing back to the file's is Saved; a
  file changed outside is never written over unasked (`ChangedOnDisk`, then `ReloadFromFile` - an
  undoable step - or `Save(overwriteOutsideChange)`); a change made during a save stays pending. One
  per draft FOLDER (`DraftSessions`), its board file read OUTSIDE the lock - the UI thread's save timer waits
  on it, code review 2026-11-02 - and read again when a `Close` came meanwhile. What follows an edit is `DraftEditRules`, shared with
  `DraftWorkbookStore.EditCore` - never a second copy; an edit that knows its schematic renames only
  once it sees the draft hands them in with it (`ApplyWithFollow`). **`ChangeComing` is raised before
  every change, on the UI thread** (an edit, undo, redo, take-back, the file read again,
  `CollectHeldChanges`) so whoever holds a change hands it over first; `TakeBack(origin)` removes an
  editor's step when its save failed - the latest, or one with steps after it by a three-way merge
  (`TakeBackEarlier`, the later snapshots rewritten) - not redoable, and also when a change handed
  over as the take-back is said becomes the latest (code review, 2026-10-26); `Save(evenUnchanged)` writes the file
  anyway, to take in new files; `WriteFileDirectly` runs a write that is more than the content under
  the session's lock (saved first - else `DraftNotSavedException` - the file read back after).
  **Asking whether the file changed looks at its stamp first** (`FileMoved`): once seen unchanged with
  a SETTLED stamp, the same stamp is the same file and the board file is not hashed again - it was, on
  the UI thread, at every switch to the Contribute tab and every component opened. **And never while a
  write of its own runs** (another thread holds `thisWriting`: no answer, the next ask judges; code
  review, 2026-10-26 - its own save, not yet taken as its own, was read back as an outside change).
- **CRT 4.0's new component data** (step 2's first slice, 2026-10-25): a component's `Package`, a
  scope baseline's `MachineState`, and two sheets of their own after the
  component sheets - "Component pins" (`ComponentPinEntry`: label, region, pin, name, direction) and
  "Component replacements" (`ComponentReplacementEntry`: label, region, part number, kind, note - one
  `Note` of several lines since 2026-11-01, owner request, when a baseline's "where it comes from" went
  too; server 8.0.0; its part number reads "Replacement part" in the pane since 2026-11-06 - owner request: it may be a technical
  number or a friendly name - the stored column keeping its name). **Stored as WORDS, not codes** (`ComponentVocabulary` - rewording one changes the
  published data). A board file writes each new row field only when filled, so a board without them
  keeps its rows line for line. **A submission that does not carry the two lists keeps the published
  ones** (`SubmissionRows.ComponentPins` null - `PublishMerge`); an empty list is "none". Every rule
  about them is a table WARNING (`BoardDataChecks.CheckCrt4Data`; the short description's limit is
  `ComponentVocabulary.ShortDescriptionLimit`, 50 since 2026-11-03: of the 539 published the longest is 118, and the 14 over 50 are repair notes). A deleted component - pane or table - takes its
  pins and replacements, a regional variant only its region's. **Every hand-made BoardData that
  carries the board's sections names EVERY property** (`BoardDataCaptionTests
  .Every_copy_of_a_board_names_every_section`) - a section a copy leaves out is dropped silently. The
  board-level "how the baselines were measured" is NOT a field: it is the board's "Scope measurements"
  files (owner decision, 2026-10-26; `ScopeMeasurementFiles`). The popup and the Overview tab show the
  new data since slice 3c (`ComponentPopupContent`, `OverviewHtmlBuilder.BuildRows`).
- **THE CONVERSION MAKES EVERY BOARD'S PINOUTS** (owner decisions, 2026-11-06: "populate all boards having KiCad data
  with its pinouts", every board's scope baselines too, and "Conversion adds them"; CRT.Data's `ConversionPinouts`, run by
  `DataTreeConversion.WithPinouts`): the pins a board's scope baselines and its KiCad data say, and the package its footprint or
  baselines make certain, only where the board has none - so release day's conversion makes them again. It was a temporary tool on the
  "Pinout" page until that day ("Create pinout from oscilloscope baselines", "Create pinouts for the whole board",
  `BaselinePinMigration`) - gone with its release guard. **CRT's KiCad reader is CRT.Data's for it** (`KiCadProjectLoader`,
  `KiCadRawProjectLoader`, `KiCadProjectData`, `KiCadComponentTable`, `KiCadPinTable`, `KiCadSheetNetAliases`, `KiCadSymbolPlacement`,
  `StampedLruCache`; their tests stay in CRT.App.Tests) - how many components a VIEW shows needs the app's marking geometry, so that is
  `BoardKiCadFacts.CountsByView`. **The proof allows only that addition** (`DataTreeConversionShippedDataTests`: a package filled where it
  was empty, every pin the workbook had kept). **A user's Excel-era draft takes the pinouts its published board was converted with, in the
  draft and its base** (`ConversionPinouts.FromPublished`, `DraftConversion`): a draft sends its whole pin list, and without them it read as
  every pin removed.
- **Not done yet**: the release day itself (convert once more, work the queue empty, switch the service - a submission CRT 3.x sent names its
  files by their Excel-tree paths; CRT itself needs nothing, its version picks the service), and pasting
  the Wiki's 4.0 pages and deleting the fourteen it no longer has (`Assets/Wiki/README.md`, "Pages
  removed at CRT 4.0").

### External integrations (`Handlers/`)

- `Online/OnlineServices` - the manifest and the data sync; its four manifest-validation predicates
  are a trust boundary and are tested. Every request names its CRT (`User-Agent: CRT <version>`).
- `Online/UpdateService` - Velopack against GitHub Releases; `UpdateChannelFilter` decides which
  release stages the ALPHA/BETA checkboxes admit, applied BEFORE Velopack ranks the feed.
- `Oscilloscope/IScopeClient` - the seam the Oscilloscope tab's sequencing takes; `ScopeScpiClient` is
  the untested TCP boundary. `ScopeCommandResolver`/`ScopeValueMapper` turn the master workbook's
  command strings into scope interactions.
- **"My scope is not listed" is an editor** (CRT 4.0 step 5's slice 5b, 2026-10-29;
  `TabOscilloscope.ModelEditor.cs` - the plan: "start from a similar model and edit each command with
  Try it, then run the full test suite"): a model of the user's OWN made from the one shown
  (`OscilloscopeModelRules.CopyOf` - every value, under a name no other model has, capitals aside:
  the drop-downs pick a model by name), kept in `OwnOscilloscopeFile` beside the settings, never the
  synced data. Every value is `ScopeModelFields`' (the main file's own key for it; a new value of the
  entry needs a field - `ScopeModelFieldsTests` counts them). **A model's values are mapped ONCE**
  (third review, 2026-10-29: four copies - `From`'s list, the resolver's switch, the editor's setter
  switch, the reflection): `OscilloscopeModelValues` (name <-> row property) is the map - the fields
  read and write by key (`Get`/`Set`), `OscilloscopeEntry.From` copies every row value by property
  name and answers `ValueOf(name)`, which `ScopeCommandResolver.GetCommandText` reads for a command's
  field (`OscilloscopeEntryTests` fails on a row value the entry lacks). A box is SAVED AS IT IS LEFT or on Enter
  (`DataManager.SaveOwnScopeModelsAsync` - listed at once, written off the UI thread, the latest list
  last), Escape puts it back, a setting CRT could not use is held back with its reason. **Try it**
  (`ScopeCommandTrial`) sends the command as typed over the tab's session: a Set command sets the value
  the scope HAS (read with the model's Read command first), a command that answers nothing is followed
  by the model's Finished and Last error (a non-zero error number = not taken), a query that does not
  answer in 10 s drops the session. **Its box is framed by what came of it** (owner request, 2026-11-07: "have the
  entire input field border in red (and green if OK)"; `TextBox.TrialPassed`/`TrialFailed`, `ShowTrialMark`): green taken,
  red not - by "Try it", by "Run full test suite" (each command as it runs, `NoteSuiteCommand`) and by a palette sent
  while the suite runs; no frame for "trying" or "not connected", and the frame goes once the box's value changes or
  another model is edited. **The editor's own refresh of the drop-downs keeps the connection**
  (`thisKeepsSessionThroughListRefresh` - connect with the similar model, try your own at once; the
  TCP port typed stays, the model's is not written over it); choosing another model closes the editor.
  **At every launch a model published under an own model's name takes it over** - exactly the same, the
  own copy goes; different, it stays as "<model> (my version)" (`OscilloscopeModelRules.Reconciled`,
  the file written back). **An own-models file that could not be read is NEVER written that run**
  (`DataManager.OwnScopeModelsProblem`, from `OwnOscilloscopeFile.LoadOrFind(..., out notWritable)` - a
  file broken or held, or 3.x's search not made): "Create model" says why and saves nothing (code
  review, 2026-10-29: its first save replaced the unreadable file with the one model). **What is held of
  the file - list, problem, read - is forgotten only where the FILE changes** (`UseOwnOscilloscopesPath`,
  from `LoadFrom`/`InitializeAsync`), never at a reload's start (second review: a failed read looked like
  another file there, the problem was lifted, and a model made during the data check was written over
  it). **ONE RULE carries the rest: the list changes and the file is written only once the file was
  READ WELL** (`_ownScopeModelsReadFrom`; a file not read yet - no main file at launch - is refused too,
  `OwnScopeModelsNotReadYet`). So `LoadOwnScopeModels` is three steps (third review: a reference check
  and a held-list juggle went): the file read only while no read went well (nothing can have changed
  since, so the read simply becomes the list); reconciled under the lock WITH THE LIST HELD NOW (an edit
  made meanwhile is reconciled, never undone); written back by the save's own write
  (`WriteHeldOwnScopeModels`, under the one gate - the latest list, in turn). **The published models stay
  listed until the new main file is read** - only one that cannot be read takes them out
  (`NoPublishedScopeModels`). **As soon as a load lists the models the tab follows**
  (`DataManager.ScopeModelsListed`, raised on the load's thread, subscribed by `Main.StartAsync` only and
  posted to `TabOscilloscope.FollowScopeModelsReloaded` - not after the whole data check; fourth review,
  2026-10-29): drop-downs made again keeping the session - unless the selection lands on another model
  than the one shown, which ends it - an own model shown follows ITS rename
  (`DataManager.OwnScopeModelRenamedTo`, from `Reconciled(..., out renamed)` - never a guess from a
  name's start), an editor whose model was taken over closes and says so. **An action on the model
  edited follows first** when no own model has its name any more (`EditedModelRowFollowingReload` - a
  box left, a name, Try it, a send), so a change typed in between lands in the renamed model. A
  delete the file did not take puts the model back and says so (`PutBackOwnScopeModels`). **The TCP port
  box follows the model shown** - picked, made, or its "Network port" changed - **except while a session
  is open**, which the box's change would end.
- **"Submit this model" sends an own model in for everybody** (slice 5b's second half, 2026-10-29;
  `TabOscilloscope.ModelSubmit.cs`, `SubmissionClient.SubmitOscilloscopeModelAsync`; the plan: "Submit
  this model sends it to the administrator's queue ... a new kind of submission on the server"). One
  request, `POST /api/submissions/oscilloscope-models`: the model's values keyed by the MAIN FILE'S OWN
  NAMES (`OscilloscopeModelValues` - a dictionary key is never put through the naming policy), the
  address as `ContactAddress` chooses it (signed in on the Maintainer tab: the account's, read only,
  sent WITH the token - a session the server ended meanwhile is then taken as that address, as the
  Submit dialog's), a message, and whether "Run full test suite" went through on the model AS IT IS NOW
  (`OscilloscopeModelValues.Fingerprint` - CRT.Data, every value - of the row the run's entry was made
  from, `OscilloscopeEntry.Row`; set at the end of a full run, said at once in an open send panel - a
  change since makes it not so; told, decides nothing). Every check is `OscilloscopeModelRules.SubmissionProblems`, asked by CRT before it
  sends and by the server when it takes it (a command is one line; a value name the server does not
  know is refused; Identify is needed). Sent, the editor says when (`UserSettings.OscilloscopeModelsSent`).
  **Its panel, "Submit this model for everybody", is ALWAYS shown under the editor, "Submit" in it** (owner request,
  2026-11-07: "please always do show this panel and have the 'Submit' button inside of that panel"; no open or
  Cancel button any more - "Run full test suite" and "Delete model" stay where they were).
  **The administrator decides on the Management screen's "Oscilloscope models"** (its tab padlocked;
  `OscilloscopeModelsView`, words `OscilloscopeModelsDisplay`): each model a card - sender in bold, the
  suite line, the message, every value beside the published model of its name, the changes marked -
  with Publish and Reject (a reason needed, mailed). The server's side is "Oscilloscope models sent in"
  under the contribution service below.
- `MiniPro/` - `IMiniproRunner` (`MiniproProcessRunner` spawns the bundled `minipro.exe`, Windows
  only; `MockMiniproRunner` simulates). **The demo mode proves nothing about a test's vectors** -
  `MockMiniproRunner` passes whatever they say. The IC test catalogue (`Assets/Data/.../MiniPro/IC
  tests`, read by released 2.5.0 too) is proven by `IcTestCatalogueShippedDataTests`: every runnable
  logic part against its datasheet model (`LogicIcReferenceModels`, typed in from the datasheet, never
  from the vectors) and through `IcTestService` against a simulated T48 (`SimulatedT48Runner`, the
  bundled minipro's two-pass pull-up/pull-down test) - a good chip passes, one broken pin fails. **A
  new runnable part without a model fails the suite.** Only add a part nobody can try on hardware
  when it uses what the shipped tests already use (driven inputs, outputs read as L/H/Z, fixed roles).
  A pin may swap between driven and read (the 4066's two directions) only across a step
  where nothing on the chip can drive it: the suite checks every mix of each two neighbouring vectors,
  the last-to-first wrap included (the T48 runs the whole list twice). A SEQUENTIAL part's test
  (74LS74, 74LS373, 74LS193, 556 - `ChipInSocket`) must read the same from any power-up state, with
  every plausible clock-pulse shape, and in every order of one vector's pin changes; the T48's "C" is
  played as high-then-low, the only shape the manufacturer's own 74193 test in minipro's DB agrees with.
- `Security/ExternalTargetLauncher` - **the only sanctioned way to open an external link or local
  file**: HTTP/HTTPS/mailto, or a local path inside the data root with an allowlisted extension.
  **A place that says nothing of its own when opening fails calls `OpenOrSay`** (an error on the "Something went wrong"
  card, named by `ShortNameOf` - never the folders of a path); one that says it itself keeps `TryOpen`. **"Reveal" on
  Windows is the shell's `SHOpenFolderAndSelectItems`**, on a thread of its own (owner report, 2026-11-07: explorer.exe's
  "/select" opened the folder without the file selected the first time), "/select" only when the shell refuses. The
  Maintainer tab's saved files open through it too, rooted at `OpenedFiles.Root`
  (`AppConfig.MaintainerOpenedFilesFolder`, the user's own app data folder, never temp; `OpenedFilesTests`
  holds the two lists together, and a `.kicad_pro` opens as JSON). What an earlier CRT left under
  `<temp>/CRT/Maintainer` goes at each open (`OpenedFiles.SweepEarlierRoot`) - through no linked folder and
  only folders named by a GUID, that place being possibly another account's.
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
  **The CRT 4.0 tree has no shared files**, so there it never applies (see "The CRT 4.0 data tree").
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
- **Push back and Reject are held to what the confirmation showed** (code review, 2026-11-02):
  `BetaRollbackRequest.ExpectedBetaContentHash`, from the rollback plan's own `BetaContentHash`. Optional,
  so an older CRT gets the old behaviour; BETA moved answers 409, and the tab reads the lists again.
- **"Push back to queue" rolls BETA back to stable's state** (`BetaRollbackFlow`), per board - it
  returns EVERY submission merged since the last promotion to the queue, so its confirmation names
  them all; "Reject" on that screen is the same rollback. It works because merged submissions keep
  their blobs. "Never promoted" is the board's RECORD, never an empty folder. Recorded as `returned`
  (`submission_beta_returns`). Paths are ORDINAL.
- **A maintainer's edit of a BOARD goes straight to BETA through the ordinary approval**
  (`BoardEditFlow`: a submission from the account, then `ApprovePublishFlow.ApproveAsync`). **Do not
  give it a path that writes BETA by itself** - the approval keeps every rule.
- **Automatic removal stays inside the board's own folder** (`AutomaticRemovalScope`) - and, in the
  CRT 4.0 tree, takes the stored files (`Files/...`) no board lists any more; a removal is always from
  a list the approver was shown (`FileRemovalPreview`, 409 if it differs). `DataTreeUsage` is the one
  rule for what a tree uses - a change may only make it keep MORE; a folder CRT reads by name must be
  in `FoldersReadByName`; a board workbook is read at its spelling ON DISK; a main file listing a
  board whose board file is missing makes the tree INCOMPLETE, and an incomplete tree removes nothing.
- **A served file is REPLACED, never opened for writing, and every folder is checked FIRST**
  (`FileReplacer`, `VerifiedFileCopy`, `TreeWriteAccess`) - the project owner copies data in as root.
  `FileReplacer` tries the rename again for a moment while the target is held open (Windows refuses
  it; CRT reads drafts on the UI thread while they save off it) - `MoveOver`, 8 tries 25 ms apart.
  The highlight file is written only when its CONTENT changes. Every blob is re-verified before a
  publish writes anything.
- **The revision date and the "# Hardware:"/"# Board:" caption are the SERVER's** at both stages
  (`PublishMerge`); `BoardDataCaptionTests` fails on a `new BoardData { ... }` naming neither. Do not
  make the client stamp a date.
- **A file of no bytes is refused, whatever its type** (owner decision, 2026-11-02: "Refuse empty files";
  `SubmissionFileRules`, code `file.empty`, one sentence on every side): at create and at an amendment on
  what a manifest declares (size 0 or the hash of no bytes), at finalise and the publish on the bytes
  (`MatchesWholeFile`; `Matches` still takes an empty text file, for a feedback attachment). CRT says it per
  file before anything is sent (`HashFilesAsync`), on the table's cell (`BoardFileState.Empty`), and as a
  file is added (`DraftPickedFiles.EmptyFiles`, `KiCadFolderScan.Empty`). No published file is empty, so
  nothing is exempted.
- **"The address is verified" is said only of an address that is** (owner decision, 2026-11-02):
  `ReviewContributorFacts.SignedIn` is a verified account found (`ContributorHistory`), as a model's is.
  Anybody can register under anybody's address, and `LoginAsync` asks nothing of verification.
- **WHICH files a submission may carry is `SubmissionFileRules`**, run at create AND in `PublishPlan`;
  `SubmissionRulesShippedDataTests` runs every rule and the publish plan over every shipped board as
  the 4.0 converter writes it - do not loosen a rule to pass it. KiCad data travels by `SubmissionKiCadFiles` (in 4.0: the names under "KiCad data/" a
  board lists). "Already held" includes bytes the published tree stores (imported at create - in 4.0
  found by checksum, `PublishedBoardFiles.FindStored`). **A KiCad STRAGGLER** (a type CRT does not
  read, in "KiCad data" - the converted trees' `KiCad-traces.json` and one `.sch`) is admitted only as
  the published board's own bytes, so finalise's content check passes it (`SubmissionKiCadFiles
  .IsStraggler`, in `MatchesWholeFile`; code review, 2026-10-25 - those boards could not be changed at
  all); the shipped-data test content-checks every KiCad file a submission carries, not only cited ones.
- **A newer submission from the same contributor replaces the older one** (`withdrawn`, only if still
  pending and never amended). **One that stays - amended, or with the first of two approvals - is
  named on its queue row and detail** (`ReviewQueueEntry.NewerSubmissionId`/`NewerSubmissionUtc`,
  `SubmissionReplacementRules.NewerFromSameContributor`; owner decision, 2026-10-13) and said in the
  Maintainer tab's amber panel - the rule is kept, the reviewer is told.
- **A new board's creator becomes its first maintainer** (owner request, 2026-10-13;
  `NewBoardMaintainerFlow`, run by the approve route when `ApproveOutcome.CreatedBoard`): the account
  the submission was SENT WITH (signed in) goes into the pool; a typed address with no account gets the
  administrator's own kind of invitation (worded for the creator); **a typed address that HAS an
  account makes nobody maintainer** (code review, 2026-10-13 - nothing proves the sender owns it, and
  the invitation that would is refused for it); an unverified or locked account is left alone. The approver is named on the pool row and the audit, and
  told the outcome (`ReviewDecisionAnswer.MaintainerNote`). It runs after the publish and never fails
  it. "Create board"'s `NewBoardMaintainerWindow` is the agreement this acts on.
- **A model's queue entry says `SignedIn` only for a verified account found** (code review, 2026-11-02):
  the card says so or says the name and address were typed in, and words the test suite as what the
  sender's CRT reported (`OscilloscopeModelRules.SuiteReportedPassed`, shared with the mail), never green.
  The models on show follow the queue's count quietly, and on screen are read at every full check whatever
  it says - a model replacing another leaves it standing (`FollowWaitingCountAsync`; no card made again for a
  read with nothing new, `OscilloscopeModelsDisplay.SameModels`) - never under a reason being typed.
- **Oscilloscope models sent in** (server 4.2.0, migration 0024, 2026-10-29; `OscilloscopeModelFlows`,
  `OscilloscopeModelEndpoints`, `AdminEndpoints.OscilloscopeModels.cs`): a NEW KIND of submission with a
  table of its own (`oscilloscope_submissions`) - nothing of the board pipeline applies. Sending needs no
  account (a typed address, or the signed-in account's), is under the contributor routes' SENDING limit,
  and at most `MaxPerAddress` (10) a day per INTERNET CONNECTION (`ConnectionOf`: an IPv4 address, an
  IPv6 /64 - all that is stored; the refusal says "internet connection", never "this address"),
  **counted and kept as ONE step** (`SendGate` - the service is one process; third review, Security: a
  burst all counted under the limit), and **at most `MaxWaiting` (100) wait at once from everybody**
  (fourth review, Security: a whole IPv6 prefix holds thousands of /64s, each model mailed to every
  administrator); **names are matched by `OscilloscopeModelRules.SameName`, never the table's collation**
  (utf8mb4_unicode_ci ignores accents - the store selects and filters in C#); a newer model of the same name from the same contributor - the
  same account, or WITHOUT one the same typed address AND connection (an address alone is no proof) -
  marks the PENDING one `replaced`; the administrators are mailed. What is stored is cut by `BoardViewRules.Clip`, the one
  rule (never half a character - utf8mb4 refuses one; the sign-in's user agent too). **The administrator
  alone decides** (the `/api/admin` group). **Publish writes the BETA and the stable main files AT ONCE**
  (owner decision 2026-10-24: "published with one click"), as "Order of boards" does - in place of the
  model of its name (`MasterListing.PutOscilloscope`; only a CRT 4.0 main file takes one), each changed
  tree's manifest rebuilt (`BoardOrderFlow.RebuildAsync`), under the `PublishLock`; BETA first, never
  rolled back. **A stable source that could not take the model leaves it IN BETA**
  (`OscilloscopeModelStates.InBeta`, a state of its own - no migration, the column is a VARCHAR):
  still listed (`inBeta` on its queue entry, the card offering Publish alone), the BETA write audited,
  its contributor not mailed yet, never replaced by a newer copy, and **never rejected** - it cannot be
  taken back out of BETA; Publish again finishes it (BETA, holding it, writes nothing) - code reviews,
  2026-10-29: marked published the retry was refused for good, left pending a Reject or a newer copy
  took it out of the list while BETA kept it. A manifest not rebuilt is said in `Problem`, and the
  contributor's mail then promises no next launch. What was written is recorded published (or in BETA)
  even when a newer copy marked it replaced meanwhile (`DecideAsync(evenIfReplaced)`, `MarkInBetaAsync`).
  **A publish that writes BETA settles every OTHER copy of its name still in BETA** - replaced, out of
  the list (`ReplaceInBetaOfNameAsync`; third review: its "Publish again" wrote older commands back over
  the newer model in both trees) - once both trees are written, guarded (a database failing there fails
  nothing written), and a copy ANOTHER contributor sent is mailed (`OscilloscopeModelSuperseded`). The
  administrator's answer promises the next launch only as the mail does (manifests rebuilt). **Reject needs a reason and takes the `PublishLock` too**, so no
  decision lands in the middle of a publish. The contributor is mailed either way; no mail fails what it
  follows. A reason typed on a card outlives a decision on another (`OscilloscopeModelsView.ShowModels`).
  **The models waiting are the administrator's work**: `GET /api/review/queue` carries
  `OscilloscopeModelsWaiting` (`OscilloscopeModelFlows.WaitingForAsync`, 0 for anybody else), counted in
  the tab's badge and on the Management button - "hide the tab while no work is waiting" hid them. The data
  reset counts, shows, fingerprints and deletes the table like the submissions. `contact_email` is
  VARCHAR(320), as every email column (`EmailColumnsMigrationTests`).
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
  raise is a MAJOR server bump. **While nobody but the project owner has connected to a revision, it is
  not raised** (owner decision, 2026-11-01: "no one (other than me) has ever connected to API version 2,
  so keep it on 2 for now"): its `api-revision-<N>.txt` is brought in line instead, and the turn summary says so. `GET /api/health` and Account > "Server version" show both revisions.
- **CRT shows the server's words** (`ApiRefusal.Read`; `ReviewApiClient.WithServersWords`).
- **An outdated CRT gets "CRT has to be updated" over the Contribute and Maintainer tabs, which cannot
  be closed** (owner request, 2026-10-09; `Main.UpdateRequired.cs`, `AppUpdateRequirement`,
  `UpdateRequiredOverlay` in `Controls/`). Once either is in use - contribution in use (a draft, or
  unread news on something sent) or the Maintainer tab on (`AppUpdateRequirement.AsksServer`), at
  launch or later - CRT asks `/api/health`'s `ApiRevision` once: a higher one covers BOTH tabs. Nobody
  using neither is asked - but a board's first draft (`Main.EnsureDraftAsync`) and "Add a new board"
  ask it FIRST, awaited, when nothing has yet (`AskApiRevisionBeforeFirstDraftAsync`), or a machine's
  first draft opens under the cover that follows. **Whether a tab is covered has ONE record each**: the Maintainer tab reads its overlay
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
  Contribute tab gets nothing new - no first draft is made, and "Add a new board" makes nothing - while
  another editor's change (the label editor's, say) still saves, onto the draft held in memory (`DraftEdits`).
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
series; every version since the second restart holds all of it. The next breaking change is `4.0.0`.

Three edits per bump: `InformationalVersion` in [CRT.Server.csproj](../src/CRT.Server/CRT.Server.csproj),
a row in VERSION.md's history saying what a caller would notice and why, and a section of one-line bullets in
[RELEASE-NOTES.md](../src/CRT.Server/RELEASE-NOTES.md) - the release's text in CRT-Server, said as a reader of
that list meets it. `GET /api/health` reports it,
without the `+<commit>` metadata (the endpoint is public). **Machine-checked:** the `Stop` hook
[hooks/server-version-bump.sh](hooks/server-version-bump.sh) warns when a file the service is built
from changed while the version stood still - "nothing a caller can observe changed" is sometimes the
right answer.

**Say the API version when it MOVES** (owner request, 2026-10-09: "whenever you change anything for
the API, then do comment what is the API version on SERVER and APPLICATION, so I always can track
this"; narrowed 2026-10-13: "It should only show when this changes"). End that turn's summary with
`API changed: Server changed from [N] to [M], Application changed from [N] to [M]`. An API change
that leaves the version alone (an added optional field or route) gets no line. Both are
`ClientVersionContract.ApiRevision` - ONE constant, so in the source they always move together; they
differ only between a deployed server and an installed CRT. Account > "Server version" shows it as
"API server version" and "API application version"; VERSION.md's history carries it per server
version (its **API** column). **Machine-checked:** the `Stop` hook
[hooks/api-version-report.sh](hooks/api-version-report.sh) prints that line once when the number
differs from the last commit, and nothing otherwise.

## Release process

Versioning lives in [CRT.App.csproj](../src/CRT.App/CRT.App.csproj) (`AssemblyVersion`/`InformationalVersion`)
- bump `InformationalVersion` before releasing; it is the only place a release version is entered.
**A MAJOR or MINOR bump is when the NuGet packages get looked at** (owner request, 2026-10-12): the
`Stop` hook [hooks/nuget-check-on-version-bump.sh](hooks/nuget-check-on-version-bump.sh) compares
`InformationalVersion` with the last release (the highest `crt-<version>.txt` below it) and, once per
MAJOR.MINOR, stops the turn: run `dotnet list Classic-Repair-Toolbox.slnx package --outdated`, advise
the project owner package by package, and update nothing until they say which.
**A package reference that only raises what another package brings in carries `PinnedOver="<that
package>"`** (2026-10-14: EPPlus 8.7.1 asks for System.Security.Cryptography.Pkcs/.Xml 10.0.10, which
known vulnerabilities cover, and a known vulnerability stops a Release restore - so CRT.Data pins both
at 10.0.13). `PinnedPackagesTests` fails once that package requires the pinned version itself, no
longer brings it in, or is no longer referenced - then delete the line. A pin LOWER than something
else needs is NuGet's NU1605, a build error, so a stale pin can never hold a version down quietly.
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
must stay sequential (a YubiKey locks after three wrong PINs). Linux builds are not signed.
There is no separate maintainer application to release any more.

**macOS builds are signed AD HOC, and the job builds the `.app` itself** (2026-10-12, after a user's
scope connect failed with "No route to host" although Local Network was switched on for CRT: 3.0.0's
bundle was unsigned, its executable signed by the .NET SDK as `apphost` with no Info.plist bound, and
its Mach-O UUID the stock .NET apphost's, shared with every .NET app). The rules, all in the macOS job:

- **CRT writes the Info.plist** (with `NSLocalNetworkUsageDescription`), and "Compare Info.plist with
  Velopack's" fails when vpk's own plist has a key ours lacks. The bundle ID
  `com.Classic_Repair_Toolbox.Classic_Repair_Toolbox` is what every released CRT carries - never change
  it. **vpk is the Velopack package's version in CRT.App.csproj, in every build job** (owner request,
  2026-10-12: `determine-version` reads it and hands it on, installed with `--allow-downgrade` for the
  self-hosted Windows runner) - updating Velopack through NuGet moves vpk too, and on macOS the plist
  check may then need our Info.plist to follow. Never install vpk unversioned or pin it apart.
- **The data pack is in `Contents/Resources/Data`**, never `Contents/MacOS` (codesign signs every file
  there as code, a non-code file's signature lives in an extended attribute, and the update `.nupkg` is
  a zip that keeps none). CRT seeds from either (`BundledDataLocation`).
- **Each executable gets its own LC_UUID** (`.github/workflows/macos-macho-uuid.py`: uuid5 of bundle
  ID, version and RID), written BEFORE Velopack signs.
- **`VPK_SIGN_APP_IDENTITY: "-"`** makes Velopack sign the bundle (deep, hardened runtime) with
  `Assets/macOS/CRT.entitlements` - Microsoft's .NET set. Do not drop it: without an identity Velopack
  signs nothing. Keep that file a bare plist (no comments - the kernel parses it).
- **`macos-verify-bundle.sh` checks the installer's app AND the update's**, and starts the installed one
  with a Velopack fast-exit hook. Its deep check is required for the installer and only reported for the
  update: the 224 .NET assemblies still sit in `Contents/MacOS` and lose their extended-attribute
  signatures in the zip. Whether macOS's Local Network check minds that is NOT known; a single-file
  publish for macOS would remove them.
- Ad hoc is not Developer ID: Gatekeeper still asks on first open, and each release is a new program
  to macOS, so it asks about the local network again after an update. Developer ID signing and
  notarisation need the paid Apple Developer Program - the project owner's decision, not taken.
