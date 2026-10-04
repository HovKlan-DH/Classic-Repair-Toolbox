# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

Classic Repair Toolbox (CRT) is a cross-platform (Windows/Linux/macOS) Avalonia desktop app that helps
hardware enthusiasts diagnose and repair vintage computers. It presents schematics, component data,
oscilloscope baselines, and interactive KiCad traces for a curated set of hardware (mostly Commodore,
plus Amstrad and ZX Spectrum boards). It also drives real test equipment: SCPI oscilloscopes over TCP
and a MiniPro USB IC programmer/tester.

## One change, every side of it

**This is one system in four parts, and a change to shared behaviour is not done until every
part that shares it has been changed in the SAME session.** The parts are the CRT desktop app
(`src/CRT.App/`, which includes the Maintainer tab - the separate CRT Maintainer application until
2026-09-29 - and, in `src/CRT.App/Controls/`, the controls the Drafts and Maintainer tabs share),
the contribution service (`src/CRT.Server/`), the shared library they both reference
(`src/CRT.Data/`), and the board DATA itself (`Assets/Data/`, and the published tree the server writes).

**The legacy PHP contribution path is NOT one of them, by owner decision (2026-09-23).**
`Assets/Webserver/app-contribution/` is going away when the new pipeline ships, and the old
method will not be supported alongside it - so a shared change does NOT have to be mirrored into
the PHP, and a change is not blocked by it. Do not spend effort keeping it in step, and do not
let its contract constrain the new design. This REPLACES the older "change both sides in the
same sitting" instruction for that pair, which is still recorded further down in the webserver
section and now applies only while something is deliberately still using the old path.
(`app-feedback` was RETIRED on 2026-10-03, owner request: CRT.Server's `POST /api/feedback` does
its job, and Apache forwards the old address for CRTs already installed - DEPLOYMENT.md, "Feedback
from CRT". `app-checkin` went the same way the same day: `POST /api/usage/check-in`, with Apache
forwarding `/app-checkin/` - its server steps were given in the session, not written into
DEPLOYMENT.md, at the owner's request. Their local copies stay in `Assets/Webserver/`, which is
git-ignored as a whole: nothing there has a git copy, so never delete a file there to "retire" it.)

**The compiler covers less of this than it looks.** Both applications reference `CRT.Data`,
so a changed method signature or a renamed property fails the build everywhere at once. That is
exactly why the dangerous surfaces are the ones it CANNOT see:

- **HTTP routes and JSON field names.** CRT (its Maintainer tab, and its submissions) calls
  `CRT.Server` over the wire. A renamed
  endpoint or a renamed JSON property compiles perfectly on both sides and fails at runtime, in
  the user's hands. The Phase 5 password-reset link was a live example: the email named a path
  the server mapped nothing at, and the tests asserted the dead URL, so nothing caught it until
  a real reset was attempted. **The review API's bodies are CRT.Data's `ReviewApiContract`
  records** (2026-09-25): a request is one record both ends use, an answer is a record the server
  returns and the Maintainer tab parses (the queue too, since 2026-09-26: `ReviewQueueAnswer`), and CRT.App.Tests' `ReviewWireContractTests` puts each
  answer through the server's JSON settings and the real parser. A new route gets its records
  there and a case in that test.
- **The workbook schema.** `BoardWorkbookSchema` names the columns the reader reads and the
  writer writes. A column added on one side only is silently blank on the other.
- **On-disk formats.** The data tree's layout, the JSON sidecar, the draft folder, the sync
  manifest. Nothing type-checks a file format.
- **Vocabulary the user reads.** A state name, a status word or a date format that appears in
  more than one application must be changed in all of them, or the same submission describes
  itself differently depending on which window it is shown in.

**So before finishing a change to anything shared, ask which of the four parts also speak it, and
change them together.** If one of them genuinely cannot be changed yet, say so explicitly in the
turn summary rather than leaving the divergence to be discovered later.

**A shared change needs a test that would fail if only one side moved.** A test per side, each
asserting its own half, passes happily while the two halves disagree - which is precisely the
failure this rule exists to prevent.

## Hands off CHANGELOG.md

**Never create, edit, rewrite, reformat or delete [CHANGELOG.md](../CHANGELOG.md) unless the
project owner explicitly asks for it in that message.** It is written by hand, in the project owner's own
words, and it is the body of every GitHub Release — an "improvement" there is not a small edit, it
is words the project owner never wrote going out under their name.

This holds even when a change would normally warrant a changelog entry, and even when the file
already has uncommitted edits in it (those are the project owner's, in progress). Do not touch it as a
"finishing touch" on a feature, do not tidy its formatting, and do not stage, commit, revert or
`git checkout` it. If you think an entry is needed, say so in your summary and let the project owner
write it. "Update the changelog" from the project owner is the only permission — and it covers that
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
[Assets/Wiki/](../Assets/Wiki/) — one `.md` file per Wiki page, named **exactly** as the page is.

### Know who is reading: hobbyists, not professionals

**CRT's users are hobby hardware enthusiasts repairing their own vintage computers.** Some are
deeply technical; most are ad-hoc "I want to see what this is" users. As a rule of thumb **nobody
is doing this for a living and nobody earns significant income from it.**

So documentation and user-visible strings must never assume a job, a workplace, customers, clients,
billing or a trade. "Pick the country you work in" was wrong for exactly this reason — a reader
simply *lives* somewhere. The same goes for "what you bill in", "your customer", "hand the job
over", "at invoice time", and example data like a customer's name in a workbook description.

Write "the country you live in", "the repair", "share the write-up", "export it". Where a page has
to name the thing a workbook represents, it is **a repair**, not a job. If a phrasing is genuinely
ambiguous, ask rather than guess.

Two things this rule does **not** touch: the app's own domain nouns (a workbook holds worklogs -
that is product naming and is fine), and ordinary English "job" meaning a task ("a re-labelling
job", "a data job"). The problem is only prose implying paid work. Code comments explaining design
rationale are also out of scope - the rule is about what a user reads.

**Nothing publishes automatically.** GitHub offers no way to push a folder in this repository to
the Wiki, so the project owner copies changed files across by hand. Editing a file here changes what
the page *should* say; the page itself changes only when they paste it. Never tell the user a Wiki
page has been updated — say which files changed and that they are ready to be pasted.

- **Filenames ARE page names**, capitals included (`MiniPro-programmer.md` →
  `.../wiki/MiniPro-programmer`). Never rename one without the Wiki page being renamed to match.
- **These files use WIKI syntax, not repository syntax**, because they are written to be pasted:
  internal links carry no path and no extension (`[Board JSON](Board-JSON)`), and images stay as
  `<img>` tags pointing at their `user-attachments` URL. A relative `./page.md` link or a
  `./images/` path breaks the moment it is pasted into a page. This is the opposite of the
  convention everywhere else in this repo, and it is deliberate.
- **`Assets/Wiki/images/`** holds reference copies only. The Wiki does not read from it, and
  repointing a page at it would break the image.
- **[Assets/Wiki/!sync-status.md](../Assets/Wiki/%21sync-status.md)** records which pages have been
  pasted. It is a record of the Wiki's real state, not of this folder's.

  **The WHOLE FILE is generated and you must not hand-edit any of it, or add anything to it.** The
  `Stop` hook [hooks/wiki-sync-status.sh](hooks/wiki-sync-status.sh) rewrites everything between
  the `<!-- crt:waiting-start -->` / `<!-- crt:waiting-end -->` markers on every turn, and those
  markers are the first and last lines of the file - there is nothing outside them. Comparing each
  page against the blob hash recorded for it in [.claude/wiki-synced.tsv](wiki-synced.tsv).

  This exists because the file previously recorded sync state only in prose ("Verified against the
  live Wiki on 2026-09-06"), which nothing could check. Six pages had changed while the prose still
  named three, and the project owner had to ask for a manual re-check every time. The point of the
  automation is that they no longer have to: the table can be trusted on sight.

  **It holds ONE table and exactly two columns: the FILE, and where that page lives in the live
  Wiki** (`Workbooks-Daily-use.md` | `Home > At the bench > Workbooks > Daily use`). Both halves of
  that are a owner instruction, not a default to improve on:

  - **No other columns.** It used to carry "Last pasted" and "Changed since", which answer a
    question nobody asks - a page is listed because it needs pasting, and how stale it is changes
    nothing about what to do. What was missing was the only thing that costs real time: *where in
    the Wiki the page actually is*, so it can be found and opened.
  - **No other content.** The mechanism explanation, the "what changed in each page" prose and the
    "not a Wiki page" notes were all removed outright. The project owner opens this file to go and
    paste pages and reads nothing else in it, and that prose pushed the table itself below the
    fold. **Do not write a "what changed" section back into it** - if a change needs explaining,
    say it in your turn summary. Mechanism notes belong here in CLAUDE.md and in
    [Assets/Wiki/README.md](../Assets/Wiki/README.md).

  **The Wiki location is DERIVED from [Assets/Wiki/_Sidebar.md](../Assets/Wiki/_Sidebar.md) at hook
  run time, never hardcoded.** The sidebar is what GitHub renders beside every Wiki page, so it is
  what the project owner actually navigates by, and deriving it means a sidebar edit cannot leave this
  table describing a structure that no longer exists. A page listed twice there (`Workbooks-tab` is
  under both "At the bench" and "The tabs") keeps its FIRST, deeper trail; a page the sidebar does
  not list falls back to `Home`. **If you restructure `_Sidebar.md`, re-run the hook and glance at
  the trails** rather than assuming the parse still fits.

  **The comparison is on CONTENT, not on commits**, and that is load-bearing. Pages are routinely
  pasted while their edits are still uncommitted, and a commit-based check then reports every
  stamped page as dirty against HEAD forever, so the list could never be cleared - which is exactly
  what happened when it was first built that way.
- **When the project owner CONFIRMS a paste** ("I pasted Configuration-tab and Workbooks-tab", or
  "pasted all of them"), run [hooks/wiki-mark-synced.sh](hooks/wiki-mark-synced.sh) with those page
  names (or `--all`). It stamps each at its current content, and they drop off the table on the
  next turn. Never stamp a page the project owner has not confirmed.
- **When asked for a sync diff**, the table already IS the answer - read it rather than re-deriving
  it, then say what changed in each page **in your reply** (not in the file) and leave the pasting
  to the project owner.

**Documentation changes ship in the same commit as the code they describe.** When a change alters
behaviour a mirrored page documents, update that page in the same commit and say so.

**This one is machine-checked.** The `Stop` hook [hooks/remind-wiki-mirror.sh](hooks/remind-wiki-mirror.sh)
maps code paths to the pages that document them, and names any whose code changed while the page
did not. It WARNS rather than blocking - whether a change is user-visible is a judgement call, and
unlike a red suite there is no objective pass/fail - so read the list and decide. When a page starts
documenting new code, add the mapping to that script's `MAP`.

**Six Wiki pages are opened by buttons in the shipped app**, across seven buttons:
`Workbooks-tab`, `MiniPro-programmer` (two buttons), `Controlling-oscilloscope-with-keyboard`,
`Synchronize-oscilloscope`, `Maintainer-tab` and `View-boards-from-online-source`. **They are named
once, in `AppConfig`** (`WikiPage*` plus
`WikiPageUrl`), and the buttons in `TabConfiguration.axaml.cs` and `ComponentInfoWindow.axaml.cs`
read them from there rather than carrying URL literals — a repository rename then changes one
place, not five. Renaming or deleting one of these pages breaks a button in builds already
installed, which no update can fix, so `WikiHelpPageNamesTests` asserts each name still matches a
file in `Assets/Wiki/`; that test failing is the rename being caught. Add an entry there for any
new help button.

**Six Wiki pages are deliberately not mirrored** (orphans superseded by newer pages) — see
[Assets/Wiki/README.md](../Assets/Wiki/README.md). Do not add them back without being asked.

## Build, test, run, publish

There is a unit test suite covering the UI-free logic in `Handlers/`. **Always run it, always add
tests for new logic, and always update the existing tests when you change covered behaviour** — see
[Tests](#tests) below for the full rules. UI behaviour has no automated coverage, so changes to the
tabs and overlays still need the app built and run by hand (see
[Assets/Wiki/Compiling-yourself-from-source.md](../Assets/Wiki/Compiling-yourself-from-source.md)
for full per-OS instructions).

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
    [Handlers/Data/SimulationOptions.cs](../src/CRT.App/Handlers/Data/SimulationOptions.cs). Active in RELEASE
    builds too, on purpose; the startup log shouts about it and the banner says `(simulated)`.
  - All three are parsed the same way (case-insensitive, surrounding quotes stripped, first match
    wins, unrecognised arguments ignored). `--simulate-update` is set for F5 and the `watch` task in
    [.vscode/launch.json](../.vscode/launch.json) and [.vscode/tasks.json](../.vscode/tasks.json).
- **To skip the online data sync while iterating, untick "Check for new or updated data at
  application launch" in the Configuration tab.** It is a normal user setting, not a build or
  command-line concern — it short-circuits the manifest fetch, and with no manifest the board-Excel
  sync and the background image sync skip themselves too.

## Tests

**There are THREE test projects**, all xUnit and all run by the one
`dotnet test Classic-Repair-Toolbox.slnx`. None needs an oscilloscope, a MiniPro programmer, a
display, a database or a network; the whole suite is about 2 minutes in Release including the build
(CRT.App.Tests is the long pole, ~1m45s - its UI tests are ~90 s of it).

**EPPlus forced a full garbage collection on every workbook it disposed** (`DoGarbageCollectOnDispose`,
true by default) - found 2026-10-03, when the suite had crept to ~10 minutes: in the CRT.App test
process, whose heap grows to ~500 MB, each forced collection cost ~0.4 s, about 85 % of the run
(dotnet-counters showed ~90 % of the time in GC pauses; a GC trace put every induced collection in
`ExcelPackage.Dispose`). CRT paid it too - on the UI thread, per board read or saved. **Every
`ExcelPackage` is now made by CRT.Data's `EpplusLicense.NewPackage`/`OpenPackage`, which turn it
off**, and `EpplusPackagesTests` fails on a `new ExcelPackage(` anywhere else in `src/`. If the suite
slows down again, measure before guessing: a class run alone (`CRT.App.Tests.exe -class ...`) against
its time in the full run tells whether the cost is the test or something building up in the process.

**A test that WRITES a board workbook as set-up goes through `CachedWorkbooks.Write`** (CRT.App.Tests,
2026-10-03), not `BoardWorkbookWriter.Write`: the writer is deterministic, so the first write of a
board hands its bytes to every later test writing the same board (it was ~13 s of the run, measured
with a sampled trace). Saves the code UNDER TEST makes stay real, and `CachedWorkbooksTests` holds that
the key covers every field of every entry type. A sampled trace is the tool for the next slowdown:
`dotnet-trace collect --profile dotnet-common,dotnet-sampled-thread-time -- <absolute path to the
test exe>` (a relative path fails to launch), then count busy samples per method.

| Project | Covers | Rough size |
| --- | --- | --- |
| [tests/CRT.App.Tests/](../tests/CRT.App.Tests/) | the desktop app's `Handlers/` (the Maintainer tab's in `Maintainer/`), plus the headless UI tests (the tab's in `Ui/Maintainer/`), including the shared `Controls/` | ~4,300 |
| [tests/CRT.Data.Tests/](../tests/CRT.Data.Tests/) | the shared board-data library | ~940 |
| [tests/CRT.Server.Tests/](../tests/CRT.Server.Tests/) | the contribution service's flows and rules | ~440 |

Counts go stale, so treat them as "what order of magnitude", not as a figure to quote - read the
real numbers off a run. (There were four until 2026-09-29: `CRT.Maintainer.Tests` went with the
separate maintainer application, its tests moving into CRT.App.Tests.)

**The framework is xunit v3 (4.x), as the `xunit.v3.mtp-off` package, on VSTest.** From 4.0 the
plain `xunit.v3` package runs on Microsoft Testing Platform v2, where `coverlet.collector` - the
`--collect:"XPlat Code Coverage"` CI relies on - is not supported, and `dotnet test` on the .NET 10
SDK then needs the new test mode opted into in `global.json`. The `mtp-off` package is the same
framework with that runner left out, so `dotnet test`, CI, the release gate, the Stop hook and
coverage all work exactly as they did on xunit 2. Moving to MTP is possible but is a change to all
of those at once, not a package bump. **xUnit1051** (pass `TestContext.Current.CancellationToken`
to every call that takes a token) is in `NoWarn` in all three test projects: it fired 208 times on
the move from v2, and rule 6 below already rules out the real I/O the token would cancel.

**v3 RANDOMISES test order on every run**, where v2's was effectively fixed - so a test that leaks
state into whichever test follows it now fails (or hangs) at random rather than in one place. The
run header prints the seed (`Starting: CRT.App.Tests (parallel mode = none, ..., seed = N)`), and
running the test executable directly replays that exact order:
`tests/CRT.App.Tests/bin/Release/net10.0/CRT.App.Tests.exe :N -reporter verbose -longRunning 30`
(the last two print each test as it starts and name any test running past 30s - the fastest way to
find a hang, since `dotnet test` shows nothing until the run ends).

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
8. **NEVER hard-code a Windows path in a test** (`@"C:\Drafts\..."`, `"D:\\data"`). The suite is
   written and run on Windows, but GitHub runs it on Linux, where a backslash is not a separator -
   so such a test passes here, passes the Stop hook, and fails only after the push. It has happened
   three or four times; the last one held back every build and both release workflows. Build paths
   with `Path.Combine` (from `Path.GetTempPath()` when one must be rooted). A literal that is safe
   on purpose - a hostile input refused on every OS, command-line text compared as text, a test that
   returns early off Windows - carries a `// windows-path-literal: <why>` comment on its line or up
   to three lines above. **This one is machine-enforced:** `TestPathLiteralTests` (CRT.App.Tests)
   scans every test file and fails - on Windows too, so before the push - on any unexplained one.

**Writing them:** one test file per class under test, named `<ClassName>Tests.cs`. Give each test a
sentence-shaped name saying what must hold (`A_faulty_chip_fails_and_names_the_failing_pin`), and put
a comment above anything non-obvious explaining *why* it matters — a reader should learn the rule from
the test. Use `TempWorkspace` for filesystem work. Cover the failure and edge cases, not just the
happy path: blank input, malformed contributed data, wrong region, wrong board side, locale-specific
number formats.

### No case lists before the code (test-first trial ENDED, owner decision 2026-10-03)

**Do not send the project owner numbered lists of test cases to confirm before writing code.** A
test-first trial ran from 2026-10-02 to 2026-10-03 - every rule change began with "the rule, a
numbered list of cases, questions" and waited for "yes to all" - and the project owner stopped it:
the lists were too much to overview. Implement as before - the code and its tests in one go, per
the rules above - and **ASK when something is genuinely the project owner's to decide** (a
behaviour the request does not settle, a wording, a trade-off), with the answer you would propose,
rather than assuming. A few pointed questions, not a list of cases.

### What is covered

| Area | Classes |
| --- | --- |
| Oscilloscope | `ScopeValueMapper`, `ScopeCommandResolver`, `ScopeCommandPaletteDefinitions`, `ScopeFormatting`, `ScopePayloadParser`, plus `TabOscilloscope`'s command sequencing via `IScopeClient` and a fake |
| IC testing | `MiniproOutputParser`, `IcTestService` (via `MockMiniproRunner` and local test doubles) |
| Security | `ExternalTargetLauncher`, `OnlineServices`' manifest-validation predicates |
| KiCad | `KiCadRawProjectLoader`, `KiCadProjectLoader`, the `KiCadProjectData` model |
| Board data | `BoardDataReader`, `BoardDataWriter`, `BoardComponentHighlightStorage`, `ComponentListBuilder`, `ComponentImageQueries`, `OverviewHtmlBuilder`, `ContactLinkFormatter`, `BoardWorkbookSchema` (both directions), `BoardDataChecks` (the one rule set: where each problem is, and that every table error is a server refusal), `SuppliedFileLookup`/`DiskFileLookup`, `DraftFingerprint` (what a draft holds, for "already sent as it is", 2026-10-03) |
| Draft table editor | `BoardTableDocument`/`BoardTableSheet` (colours, ghost placement, agreement with `BoardDataDiffer`, moving rows, placing new components on save), `BoardTableHistory` (undo/redo of every change kind, the saved state), `BoardTableRowDrag` (live row drag, ghosts never targets, one step per drag, a return trip leaves none), `BoardTableClipboard`, `DraftTableSession` and `DraftWorkbookStore.EditIfUnchanged` (the refuse-a-stale-save rule), `ComponentPlacement`, the checks on the cells (`BoardTableProblemsTests`), `BoardTableRowFilter` (the pills as the filter), `BoardTableProblemWording`, `BoardTableSheet.DeleteRows` and `BoardTableDeletedWith.Combine` (several rows deleted as one step), `BoardTableSearch` (the search box's rows and runs), the document-wide colour-key counts |
| Worklog | `WorklogManager` (including `ResolveActiveWorkbook`, `AddEntryRecord`, `DeleteEntry`, the non-reusing id counters, `IsResolvedState`, `IsWorkbookStatusOpen`, `GetAllWorkbooks`), `WorklogEntryScope`, `WorklogSearchQuery`, `WorklogSearchIndex`, `WorkbookSummary`, `WorkbookExportModel`, `WorkbookPdfExporter.WriteZip` (the archive only), `WorkbookPdfExporter.ConfigureQuestPdf` and the icon font (the QuestPDF settings, not the layout), `WorklogAttachTargets`, `WorklogAttachmentWriter` |
| Text links | `TextLinkFinder` (which runs in a user-typed note are web links) |
| Settings / startup | `UserSettings`, `DataManager` (data-root + master workbook), `DataValidator` (smoke, and its log line - its rules are `BoardDataChecks`'), `SimulationOptions`, `TabBadge` and the Drafts tab's badge (`Ui/DraftsTabBadgeTests.cs`: the unread count on a real `Main`, and when the check while CRT runs asks anything) |
| Updates | `UpdateChannelFilter` (which release stages the ALPHA/BETA checkboxes admit) |
| Maintainer tab (`Handlers/Maintainer/`) | `ReviewApiRoutes`, `ReviewApiParser`, `ReviewApiClient` (its User-Agent, via `AnsweringHttpHandler`), `ReviewWireContractTests` (both ends of the review API), `ReviewSession`/`ReviewSessionStore` (the `"ReviewSessionStore"` collection), `MaintainerSettingsMigration`, `QueueRefreshRules`, `SubmissionPrefetch`, `MaintainerModes`, `ReviewQueueDisplay`, `SystemsDisplay`, `SystemPlacementDisplay`, `ProductionDisplay`, `UnusedFilesDisplay`, `MaintainerAssignmentDisplay`, `ReviewNotInTable`, `ReviewContributorLine`, `ReviewContributorHistory`, `MaintainerModes.EntryToOpen` (which entry a queue opens on), `SubmissionViews` (the three views and the Files count, held to the tree's own headline), `SystemSections` (a system's six views on the Systems screen, and what its table sends), `SystemHistoryDisplay` (the History view's cards, months and change lines), `SystemOrderDisplay` (Account > "Order of systems"), `ReviewDecisionWording`, `ReviewTableWording`, `ApprovalGate`, `ApprovalWording`, `DraftDiscardWording`, `FileRemovalWording`, `MaintainerWaitWording`, `FileTree`, `OpenedFiles`, `ReviewTableFiles`, `ReviewFileComparison`, `ReviewImageComparison`, `ReviewHighlightGeometry`, `ReviewScopeBaseline`, `ReviewSummaryPresenter`, `PoolAction`, `ContactAddress` (which address the Feedback tab and the Submit dialog use while signed in), `MyAccountRules` and `ReviewSession.WithAccount` ("My account" on the Account screen - the "Your account" window until 2026-10-04 - and the "Logged in as" line), `SystemDeletionWording` (Account > "Delete a system"), `DataResetWording` (Account > "Reset contribution data": the button's rule, the counts) and `ApiUsageDisplay` (Account > "API usage": the groups and their order) |
| Headless UI (`Tests/.../Ui/`) | All nine tabs built headlessly, the worklog and Workbooks palettes, component highlight selection and schematics zoom, plus `Main` itself, the label editor's full edit cycle, the worklog area-marking flow, `ComponentInfoWindow`, the oscilloscope's SCPI sequencing, and the Configuration/Overview/About tabs - see [Headless UI tests](#headless-ui-tests) |
| Geometry (`Handlers/Geometry/`) | `PolygonGeometry`, `RectGeometry`, `KiCadLayerGeometry`, `KiCadPadGeometry`, `OverlayCullGeometry`, `KiCadOverlayCacheKeys`, `KiCadOverlayNetCache`, `ViewportMath`, `KiCadNetGraphBuilder`, `KiCadHoverIndex`, `HighlightRectBuilder`, `LabelEditorGeometry`, `LabelEditorSnapGeometry`, `TraceGeometry`, `KiCadCalibrationGeometry`, `WorklogBadgeLayout`, `ExportOverlayGeometry`, `WorklogDefaultAreaGeometry`, `ColumnAutoFitGeometry` (which column a table heading's edge resizes, and the double-click fit's width), `RowDragSlots` (the frozen-slot drag maths the worklog's Photos/Files lists and - through `ListRowDrag` - the Drafts tab's Schematic images window and the Maintainer tab's drop-down placement share) |

`Handlers/` is where the real coverage is; most of the uncovered remainder is `Tabs/` and `Main/`,
Avalonia code-behind that is verified by running the app.

**No coverage percentage is recorded here on purpose** - a figure in a document goes stale silently
and then gets quoted as fact. Measure it when you actually need it:

```
dotnet test Classic-Repair-Toolbox.slnx --collect:"XPlat Code Coverage"
```

and read `lines-covered` / `lines-valid` from the `coverage.cobertura.xml` it writes under
each test project's `TestResults/`. **There is ONE REPORT PER TEST PROJECT - three of them - and none
of them is the total.** Each covers what its own project exercised, and adding them up counts
`CRT.Data` three times. Merge them line by line first, the way CI does, with ReportGenerator
(`reportgenerator "-reports:<all three, ;-separated>" -targetdir:<dir> -reporttypes:Cobertura`), and
read the merged `Cobertura.xml`. **Always quote the denominator alongside the percentage, and say
which build configuration you used** - Debug and Release instrument different numbers of lines
(~26.7k vs ~21.4k), so two bare percentages from different configurations are not comparable.

**CI already computes it on every push**, which is the one place a figure cannot go stale:
[build-and-unittest.yml](../.github/workflows/build-and-unittest.yml) collects coverage alongside
the test run, merges the three reports with a pinned ReportGenerator, and renders the merged totals
into the run's GitHub job summary via [coverage-summary.sh](../.github/workflows/coverage-summary.sh),
with each project's raw Cobertura XML and the merged one kept as a 7-day artifact for per-file
numbers. It used to summarise whichever single report `find | head -1` returned first, so the job
summary quoted one project's figure as the total; if the merge fails it now skips the summary
instead. It is **reported, never enforced** - there is deliberately no
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
`InternalsVisibleTo` entry in [CRT.App.csproj](../src/CRT.App/CRT.App.csproj), so
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
- `ReviewSessionStore.InitialiseAt(path)` — the Maintainer tab's remembered sign-in, pointed at a
  temp file. **Never call `ReviewSessionStore.Initialise()` from a test** (it resolves the real
  AppData file, which may hold a live, DPAPI-protected maintainer token), and point it back at
  nothing (`InitialiseAt(string.Empty)`) when done - a later UI test showing the tab would otherwise
  restore a session from it. Its tests share the `"ReviewSessionStore"` collection.
- `TabMaintainer.UseRememberedChoices(...)` — `Main` hands the Maintainer tab its "Show changes
  only" from `UserSettings` through this; a tab a test builds on its own never touches `UserSettings`.
  `UseRememberedSelections(...)` does the same for the two queues' last entries, and `UseTabBadge(...)`
  for the tab's badge.

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

**`UiTest.RunAsync` must hand the test back OFF the dispatcher thread, and it does so explicitly.**
The session completes the awaited task from its own dispatcher thread with no
`RunContinuationsAsynchronously`, so a plain `await` continues inline THERE. xunit 2 hid this by
running every test under its own `SynchronizationContext`; xunit v3 has none, so the rest of the
test - and xunit, running the next test - stayed on the dispatcher thread, and the next synchronous
`UiTest.Run` queued its body to the thread it was blocking. The whole CRT.App.Tests run hung, idle,
on whichever sync UI test the random order put after an async one. `RunAsync` now hops to the pool
via `ContinueWith(..., TaskScheduler.Default)`; `UiTestTests.cs` pins it, both tests failing against
the plain-await version. Do not "simplify" it back to one `await`.

| File | Covers |
| --- | --- |
| `UiTestTests.cs` | The harness itself: that after `UiTest.RunAsync` the test is no longer on the headless dispatcher thread, that a synchronous `UiTest.Run` completes straight after an async body (asserted with a 30s timeout, so a broken harness fails the test rather than hanging the suite), and that a body changing the shared Application's theme fails with the theme put back |
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
| `BoardTableEditorTests.cs` | The Drafts tab's table editor as painted: one tab per sheet with its change count, the marker column then the schema's columns, and - in a SHOWN window, reading real `DataGridCell.Background`s - a changed cell orange while its neighbour is not, an added row green, a deleted row red with the `BoardTableDeleted` strike-through class, a duplicate row uncoloured and counted by its problem but not on the tab (it was violet "Flagged" until 2026-10-03), a cell's tooltip letting the pointer through to the cell below it, each legend count sharing one pill with its own word, a picked pill (the filter) surviving a sheet switch, a clicked cell unfilled (or keeping its orange) inside a 2px dashed red frame with the grid's own frame and fill handle hidden, the drag grip centred, Ctrl+Z / Ctrl+Y / Ctrl+Shift+Z on real keys (a typed word is ONE step, Ctrl+Z inside a cell being edited is left to it, undo returns to the sheet of the change, works after "Delete row" and from a toolbar button, and greys "Save changes" out again), and a cell edit recolouring once the posted refresh runs (fails if the per-column cell theme binds the wrong column). Plus a REAL pointer drag by the grip (the dashed placeholder row mid-drag, the order live, one undo step; a click is not a drag; past the top edge steps; ghosts and the filtered view refuse), one "Insert row" and no move buttons, zero pills outlined and unpicked ones dimmed, Credits the last sheet tab, the toolbar's enablement per selected cell, insert/delete (there are NO "Restore row" / "Revert cell" buttons - removed by the project owner once undo existed; the model keeps `RestoreRow`/`RevertCell` for the Maintainer tab), a picked pill surviving a MAIN tab switch and never applied to a table it would show nothing of (kept for the next), pills clicked with a real pointer, two pills showing either kind, single-cell paste (Excel's trailing line break dropped, a block refused with a message, a ghost refused), copy quoting, `EditTriggers` not editing on a single click, save and the refused-when-changed-on-disk message, watching the draft file (an outside change reloads when nothing is unsaved, raises the warning bar and greys Save when something is, never reloads under a cell being typed in; the open-in-Excel notice follows the lock file, edits or not, and shows at once when a table is opened on a draft already open in Excel; the outside-change reload is announced; the timer runs only while on screen), and every `BoardTable_*` key in BOTH themes |
| `BoardTableEditorSelectionTests.cs` | Several rows selected and deleted in one go (2026-10-02), with a REAL pointer: Shift+click a range and Ctrl+click one in or out, a plain click (also inside a selection, and after a search) selecting one cell again, "Delete row" deleting every selected row with one Ctrl+Z back, only red rows greying it out, a picked pill keeping hidden rows out of a range, the cursor landing where the first deleted row was, a grip drag moving one row, copy/paste staying on the framed cell, the button saying "Delete 8 rows", the Delete KEY deleting nothing (it did, through the grid's `CanUserDeleteRows`), and the veil over several selected cells |
| `BoardTableEditorSearchTests.cs` | The table's search box (2026-10-02): typing narrows the rows and marks the found text in `Workbooks_SearchHit_*` - only the text found, the rest of the cell in its ordinary colour, each cell keeping its own fill - the clear button, together with a picked pill, moving to the first sheet with a match and hiding the others' tabs, a search matching nothing anywhere keeping only the sheet on screen with its line (cases 1-5 of 2026-10-03), no row moving while it is on, kept across sheet and tab switches and a save, emptied by closing the table or opening another draft. Case 21 (the Maintainer tab's table has it too) is in `Maintainer/TabMaintainerTableTests.cs` |
| `BoardTableEditorChecksTests.cs` | The checks as the table paints them (2026-10-02), in a SHOWN window: a problem cell's corner mark read off its real `LinearGradientBrush` (the level's colour, then the cell's own wash - a changed cell with a warning keeps its orange), the colour key counting the WHOLE DRAFT on every sheet (owner decision, 2026-10-02 - it counted the sheet on screen), picking "Errors" moving to the errors and a fixed row leaving the view, a sheet tab with errors or warnings showing its name alone (its pills went 2026-10-02), the line for a problem in no row, and the new `BoardTable_*` keys in both themes |
| `BoardTableEditorTextWrapTests.cs` | Long texts (owner requests, 2026-10-04), in a SHOWN window: EVERY cell wraps, with no check box ("Wrap text" was one for a day; now the grid's `CellsWrap` class and no `RowHeight`), a long note's row grows while a short one stays `MinRowHeight`, and a narrower column wraps more; pressing a column heading frees every column of `MaxColumnWidth` (420) so a drag can widen it, built capped again for the next sheet; and a real DOUBLE-CLICK on a heading's right edge, or on the next heading's left edge, fits the column to EVERY row - a row the grid has not built included, which the grid's own `CanUserResizeColumnsOnDoubleClick` (turned off) missed; both cases fail against it - never wider than the table and then scrolled wholly into view (it grew off the right edge before), while the marker column's edge fits nothing. The width rule is `ColumnAutoFitGeometryTests`' |
| `UnsavedTableEditsWindowTests.cs` | The table editor's unsaved-edits prompt: Enter and Escape both CANCEL, including with "Discard edits" focused (Tunnel route, as `DeleteWorklogWindowTests`), and neither the reload wording nor the draft-changed-on-disk one offers Save (the latter says why), and the saving-elsewhere notice has Cancel alone and sends you to the Drafts tab |
| `Maintainer/MainMaintainerTabTests.cs` | The Maintainer tab as CRT's window handles it, on a real `new CRT.Main()`: hidden until "Enable Maintainer tab" is on, placed between Drafts and Configuration, selecting it collapses the sidebar (columns 0 and 1 of `RootGrid`, `LeftPanel`, `MainSplitter`) and the worklog bar and leaving it puts back the width that was there, `UserSettings.LeftPanelWidth` never written with the collapsed 0, the worklog bar coming back only when the worklog is on, turning the tab off while selected moving the selection off it, the component filter never taking focus over it, and the table's filter (its picked pills) read from and written back to `UserSettings.MaintainerTableFilter`, the tab's header badge adding up the queue and BETA waiting for this account (gone, with its tooltip, once nothing waits), and that the badge counts as seen only with the tab on and the window not minimised. Its settings file is written as `{}` first: `UserSettings.LoadFrom` on a MISSING file keeps the previous test's settings |
| `Maintainer/TabMaintainerQueueTests.cs`, `TabMaintainerModesTests.cs`, `TabMaintainerTableTests.cs`, `TabMaintainerSubmissionViewsTests.cs` (the three submission views: opening on Board data, the Files count, switching without closing the table, the Contributor view's record, a new submission resetting it), `TabMaintainerOpenOnEntryTests.cs` (both queues opening on the last entry or the first, never over one still open), `TabMaintainerDecisionClickTests.cs` (one click on Approve with the pointer resting on it reaches it - the tooltip trap), `TabMaintainerScreenTabsTests.cs` (the screens a tab strip across the top in CRT's tab red, the views one joined switch), `TabMaintainerInvitationTests.cs` ("I have an invitation" shown alone, read off a shown window, and Cancel back to the form as it was) | The Maintainer tab itself (the separate app's `MaintainerMain*Tests` until 2026-09-29): the queue grouped by board, the four screens and their badges, the table as the submission view, the decisions and their colours, waits under the overlay (through `MaintainerTabHost`, a Grid with one `BusyOverlay` standing in for Main - so the fade lands on the tab, as it does in CRT's window), and quitting CRT's question about unsaved table edits (`ConfirmLeavingTableAsync`, answered by `UnsavedTableEditsAnswerForTests`) |
| `Maintainer/BetaViewTests.cs`, `SystemViewTests.cs`, `SystemPlacementViewTests.cs`, `UnusedFilesViewTests.cs`, `FileTreeViewTests.cs`, `RollBackBetaWindowTests.cs`, `DraftDiscardShownTests.cs` | The Maintainer tab's right-hand panels and dialogs, moved unchanged from the separate application's tests |
| `Maintainer/MaintainerPoolViewTests.cs`, `SystemOrderViewTests.cs` | Account > Maintainers (2026-10-04 - the Systems screen's pool tests, moved with the controls, plus choosing the system) and Account > "Order of systems" (a move offering Save, "Put back as saved", every system sent in order, a refusal keeping the moves) |
| `Maintainer/SystemViewSectionsTests.cs`, `TabMaintainerSystemTableTests.cs` | A system's six views (2026-10-03, History 2026-10-04): the switch and the view staying chosen for the next system, the Board data table (opened unmarked with no description box, read-only for a system not maintained or waiting in BETA, the save's order - check, reason, publish - and what it carries, a cancelled reason or a refusal keeping the change, a published change showing BETA as it is now, a table that cannot be read again closed, the `LeavingSystem` prompt and its Save), the Files listing, and the tab asking before another system replaces a change not published and opening one made but not published under Contributor Submissions |
| `Maintainer/SystemViewBetaCatchUpTests.cs` | A system's table and files following BETA (owner report, 2026-10-04): the owner's case end to end (a highlight naming "hest" warned about, gone once BETA moved), not read again when nothing moved, the sheet kept, a table read while the system waited no longer read-only after a promotion, never read again under unsaved edits (the warning instead, read again once undone), not read again for its own publish, closed when BETA no longer holds the board, and a file list dropped only when BETA's content moved |
| `Maintainer/PublishSystemChangeWindowTests.cs` | The reason asked before a Systems-screen change goes to BETA (2026-10-03): the files it removes named first, Publish off until a reason is typed and the reason handed back trimmed, a reason kept on reopening, and Enter/Escape cancelling with Publish focused (Tunnel) |
| `Maintainer/SystemDeletionViewTests.cs`, `DeleteSystemWindowTests.cs` | Account > "Delete a system" (2026-10-03): every system with its own Delete button, a confirmed delete sending back the shown plan's fingerprint and the reason, a cancelled one sending nothing, a blocked system showing the server's reason with no confirmation; the confirmation listing what goes, asking a reason only when a contributor is mailed, and Enter/Escape cancelling with Delete focused (Tunnel) |
| `Maintainer/DataResetViewTests.cs`, `ApiUsageViewTests.cs` | Account > "Reset contribution data" and "API usage" (2026-10-04), against `AnsweringHttpHandler`: the counts shown and the button off until `RESET` is typed, nothing typable while the server's switch is off, a reset sending back the shown counts' fingerprint and reading them again, a 409 saying so; the usage asking for the days chosen and showing the launches, then each area's routes |
| `Maintainer/MyAccountViewTests.cs` | Account > "My account" (2026-10-03 as the "Your account" window; a panel of the tab since 2026-10-04), shown, against `AnsweringHttpHandler`: every account the server returns passed on at once, NOTHING passed on while a new address only waits for its code, refusals in red under their own section, the two new passwords that differ, Sign out handed to the tab, every box emptied on signing out, and a change made elsewhere followed |

**The Application is built ONCE per assembly, not per test** - `[assembly: AvaloniaTestIsolation(
AvaloniaTestIsolationLevel.PerAssembly)]` in `TestAppBuilder.cs` (2026-09-26). Avalonia's
DEFAULT is `PerTest`, which tears down and rebuilds the `Application` AND its `Dispatcher` before
every test - and that rebuild intermittently threw `InvalidOperationException: The calling thread
cannot access this object because a different thread owns it` inside
`HeadlessUnitTestSession.EnsureIsolatedApplication`, i.e. BEFORE the named test's body ran. It
landed on a different innocent test each time, always passed in isolation, and blocked three
handovers in one day at about 1 failure per 3 full `CRT.App.Tests` runs. **It made the failure much
rarer (about 1 in 11 runs) but did NOT remove it** - `EnsureIsolatedApplication` runs on every
dispatch and only skips the REBUILD, so the window narrows without closing. Treat an empty
"Error Message" on a headless UI test as this, re-run with `--logger trx` to confirm the trace, and
do not declare it gone on a handful of green runs. `PerAssembly` is what this
suite always assumed anyway (`UiTest` holds one session per assembly; every UI test shares the one
"HeadlessUi" collection and its dispatcher thread). **Its condition: no test may mutate the shared
`Application`** - no `RequestedThemeVariant` assignment, no `Styles.Add`; read theme resources with
an explicit variant through `TryGetResource`, and isolate global state through the app's own seams
(`UserSettings.LoadFrom` and friends), never through Avalonia's teardown. **The theme half is
machine-checked** (2026-10-04, after CI failed on it): `UiTest.Run`/`RunAsync` fail a test whose body
changes `Application.RequestedThemeVariant` and put it back - `SmallTabsTests`' theme drop-down
went through `App.ApplyConfiguredTheme` and turned the rest of the run dark, so whichever test the
random order put next and compared a drawn colour with the light theme failed instead. A control
that applies a theme takes a seam (`TabConfiguration.ApplyThemeOverrideForTests`).

**Do NOT add the `Avalonia.Headless.XUnit` package to get `[AvaloniaFact]`.** It was first kept out
because it needed xunit v3 while this suite was on xunit 2 (adding it made every `Fact` and
`InlineData` ambiguous, ~850 build errors). The suite is on v3 now, but at 12.1.3 the adapter is
built against xunit v3's 3.2 extensibility API, which the 4.0 this suite runs on changed, and it
has not been tried against 4.x. `UiTest.Run(...)` drives the same public session API directly and
depends on no test framework at all. Anything touching a control must go through it, or Avalonia
throws for want of a dispatcher.

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
  **`UpdateService` is half excluded for the same reason.** Which release STAGE a user has opted
  into is pure string work and *is* covered, by `UpdateChannelFilterTests`. It lives in
  `UpdateChannelFilter` rather than inside the service precisely so it can be: GitHub's "prerelease"
  flag is all-or-nothing, so the stage has to be read off the version string, and getting that
  classification wrong silently offers a user a channel they never asked for. `StageFilteredUpdateSource`
  (the `IUpdateSource` decorator that applies it to the feed BEFORE Velopack ranks it) stays
  uncovered - it is a thin pass-through over the real network source.
- **`DataValidator`'s findings.** `ValidateAllDataAsync` returns a bare `Task` and reports everything
  through `Logger`, so the tests only prove it walks real data without throwing, and pin its log line
  (`Describe`). **Its rules are CRT.Data's `BoardDataChecks` since 2026-10-02** - the same rules the
  Drafts tab's table marks and the server refuses on - and THOSE are fully tested, in CRT.Data.Tests.

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

The test project lives as a sibling at [tests/CRT.App.Tests/](../tests/CRT.App.Tests/), not inside
the app's own project folder ([src/CRT.App/](../src/CRT.App/)), so
[CRT.App.csproj](../src/CRT.App/CRT.App.csproj) needs no `Tests/**` compile-glob exclusion — the
SDK's default `**/*.cs` glob only ever walks down from a project file, so it cannot reach a sibling
directory in the first place. This replaced an earlier layout where both projects shared one folder
and the app csproj carried explicit `Compile Remove` blocks for `Tests/**` and `Assets/Webserver/**`
to keep the SDK's glob from pulling either into the build.

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
- **`Controls/` holds the controls more than one tab or window uses** - the table editor
  (`Controls/BoardTable/`, shown by the Drafts tab and the Maintainer tab), `BusyOverlay` and
  `ListRowDrag`. They were a separate `CRT.UI` library while the Maintainer tab was its own
  application, and were folded into CRT.App on 2026-09-30; their pure parts went to `Handlers/`
  (`RowDragSlots` in `Geometry/`, `ThemeResources` in `Theme/`, `WaitLimit`/`WaitWording` in
  `Waiting/`). A control used by one tab stays in that tab's folder.
  **They still touch nothing of CRT's own orchestration - no `Main`, `DataManager`, `UserSettings`
  or `Logger`** (the rule the separate library's compiler used to enforce). A host hands them what
  they need - a folder, a document, a remembered choice (reported back through
  `FilterWantedChanged`) - and CRT.Data's `CrtLog` is their logging seam. That is what keeps
  headless tests of them off the user's real settings and log files.
  `SharedControlsIndependenceTests` fails on any such reference in a `Controls/` source file.
- **No button label has "..."** (owner rule, 2026-10-03: 'rename "Account..." button to "Account",
  and please have a rule for this, as this is not the first time I have seen this - a button should
  never have "..."'). A button says what it does in plain words - "Account", "Browse", "Choose image
  files" - with no ellipsis, three dots or the single character, whether or not it opens a window.
  Status lines and placeholders ("Loading, please wait...") are not buttons and are not affected.
  **Machine-checked:** `ButtonLabelTests` (CRT.App.Tests) fails on a `*Button` whose `Content`, or a
  `TextBlock`/`Run` inside it, carries one in any `.axaml` under `src/`, and on a `Content` set to such
  a string literal in any `.cs` there. A Wiki page naming a button uses the same label.
- **Everything only the administrator can use carries Font Awesome's padlock** (owner rule,
  2026-10-04: "Mark all admin locked things (generally everywhere in application) with the Font
  Awesome "fa-lock" icon, so it is clear which things are applicable only for the admin. No
  title/helper text on this icon"). `<TextBlock Classes="AdminLock" Text="&#xf023;" />` right after
  its name - the style is App.axaml's, with the 1px top room the padlock needs (see
  `FontAwesomeGlyphMetrics`) - and NO tooltip on it. Today that is the seven administrator's entries
  on the Maintainer tab's Account screen; nothing else in CRT is administrator-only (searched
  2026-10-04 - the administrator's second approval of a shared-file change is a role in an approval,
  not a locked control). Such an entry is also hidden from everybody else. Machine-checked for the
  Account screen: `TabMaintainerModesTests.Every_administrators_entry_carries_the_padlock_with_no_tooltip_and_nothing_else_does`
  lists the administrator's entries by name, so a new entry has to be decided one way or the other.
- **Every tooltip opens at its control's EDGE, never under the pointer** (owner report, 2026-09-30:
  Approve needed three clicks; `Handlers/Theme/ToolTipPlacement`, registered in `App`'s static
  constructor). Avalonia's default places a tooltip 20 px below the POINTER; with no room it flips
  above and keeps the 20 px offset, which lands it back OVER the pointer - and CRT's tooltips are
  native popup windows, so the next click went to the tooltip, not the control: the bottom rows of
  every table, list and file tree, and the decision buttons. A class handler on `ToolTipOpening` now
  places each at `PlacementMode.Bottom` with offset 0 - under its control, or flipped on top of it,
  never covering it. **Do not set `ToolTip.Placement="Pointer"` or a positive `VerticalOffset` on a
  control** - a control that chooses its own placement opts out of the fix. The defaults could not
  simply be overridden: Avalonia registers them on `Control` itself. `ToolTipPlacementTests`
  reproduces the lost click and fails with the handler switched off.
- **Every wait the user watches runs under `BusyOverlay`** (owner decision, 2026-09-28:
  "I want this method everywhere in the entire project where there is a Wait"). One per window, the
  last child of its root Grid (`Main`, and each dialog that waits:
  `MySubmissionsWindow`, `SystemFilesWindow`, `ComponentContributionWindow`) - the Maintainer tab
  itself has NONE, its waits run under Main's (its file tree window went on 2026-09-30, when the tree
  became a submission's Files view, and its "Your account" window on 2026-10-04, when it became the
  "My account" panel); anything in the window
  finds it with `BusyOverlay.For`/`BusyOverlay.RunAsync(this, ...)`. It blocks clicks AND keys at once,
  dims after 300 ms, and gives up after `WaitLimit.Maximum` (2 minutes) **with nothing happening** -
  work that calls `WaitContext.Report` restarts the clock, so a long upload is never cut off while it
  moves. **A timeout is never "it failed"**: a call that changed something on the server looks again
  and says what it found (`WaitWording.AfterTimeout`; the Maintainer's `ServerWait` turns a timeout
  into `ReviewApiFailure.TimedOut`); local work that cannot be stopped uses `RunLocalAsync`, which
  lifts the overlay and still returns the work's real result. Several steps are one wait through
  `HoldAsync`, or the window flickers between them. `BusyOverlayHostsTests` fails on a host that loses
  its overlay. **Deliberately NOT under it**: background work nobody pressed for (sync, status checks,
  the minute queue check, board views), board switching, the Submit dialog's upload (its own progress,
  and CRT stays usable), the MiniPro IC test and the oscilloscope's live output, and the SYNCHRONOUS
  table-editor LOAD (the overlay cannot draw while the UI thread is blocked). **The table's SAVE is
  under it since 2026-09-30** (`BoardTableEditor.SaveAsync`; owner question - a save of the largest
  board takes 0.6-2 s): the rows are taken on the UI thread (`DraftTableSession.PrepareSave`,
  `BoardTableDocument.TakeSaveSnapshot`), the workbook read/write and the re-read run on the pool
  under `RunLocalAsync`, and the file watch skips while it writes (`thisSaveInFlight`). The
  synchronous `Save()` stays for tests; the button and the unsaved-edits prompts use `SaveAsync`.
- Biggest files right now, for context budgeting: [Tabs/Contribute/ComponentContribution.axaml.cs](../src/CRT.App/Tabs/Contribute/ComponentContribution.axaml.cs)
  (~2,200 lines), [Tabs/Schematics/ComponentInfoWindow.axaml.cs](../src/CRT.App/Tabs/Schematics/ComponentInfoWindow.axaml.cs) (~1,900),
  [Handlers/Data/UserSettings.cs](../src/CRT.App/Handlers/Data/UserSettings.cs) (~1,900),
  [Handlers/Data/WorklogManager.cs](../src/CRT.App/Handlers/Data/WorklogManager.cs) (~1,900),
  [Tabs/Schematics/TabSchematics.Worklog.cs](../src/CRT.App/Tabs/Schematics/TabSchematics.Worklog.cs) (~1,700),
  [Handlers/Data/DataManager.cs](../src/CRT.App/Handlers/Data/DataManager.cs) (~1,700). Read the part you need rather
  than the whole file. `TabOscilloscope`, `Main` and `WorklogEntryEditorWindow` used to head this list
  and no longer do - each has been split into partials, so go via the file map in its `.axaml.cs`.

## Architecture

### Bootstrap (`Main/`)

- [Main/Program.cs](../src/CRT.App/Main/Program.cs) — entry point; initializes Velopack then starts Avalonia.
- [Main/App.axaml.cs](../src/CRT.App/Main/App.axaml.cs) — contains `AppConfig`, a static class holding nearly every
  tunable value in the app (file/folder names, sync URLs, timeouts, zoom limits, debug flags, version
  helpers). **Check here first before hardcoding a new constant elsewhere.** `App` itself wires up theme
  application (including JSON-defined user-preference theme colors), global exception logging, and the
  startup sequence: show `Splash` → `DataManager.InitializeAsync` (loads/syncs hardware data) → open
  `Main` window → fire-and-forget version check-in.
- [Main/Main.ModeHint.cs](../src/CRT.App/Main/Main.ModeHint.cs) — the khaki "what to do next" label in the tab-header
  row, shown while a mode (e.g. worklog area-marking) is waiting for the user to act. `ShowModeHint`/
  `HideModeHint`; it clears itself on the first pointer press.
- [Main/Main.axaml.cs](../src/CRT.App/Main/Main.axaml.cs) — the main window's code-behind. It acts as the central
  controller coordinating board selection, schematics zoom/pan/thumbnails, and cross-tab state. It
  reaches directly into `TabSchematicsControl` members (`currentThumbnails`,
  `highlightIndexBySchematic`), so changes to those ripple here. **Split into partials by area** -
  `Main.BoardSelection.cs`, `Main.ComponentPopup.cs`, `Main.DataSyncStatus.cs`,
  `Main.SchematicsWindows.cs`, `Main.Updates.cs`, `Main.Worklog.cs` and `Main.ModeHint.cs`, each with
  its own header; the file map in the `.axaml.cs` is the index, so go there rather than grepping.
  `TabOscilloscope` and `WorklogEntryEditorWindow` were split the same way at the same time.

### Tabs (`Tabs/`)

One folder per UI tab — `About`, `Configuration`, `Contribute`, `Drafts`, `Feedback`, `Maintainer`,
`Oscilloscope`, `Overview`, `Resources`, `Schematics`, `Workbooks` — each an Avalonia `UserControl` (`.axaml`) with
its logic in the paired `.axaml.cs`. `Worklog/` is the odd one out: it holds the worklog's dialog
windows, not a tab.

**`Drafts/` is the local-first authoring tab** from Phase 2 of
[Assets/NewContributeStrategy.md](../Assets/NewContributeStrategy.md) — a contributor edits a board
locally, sees how it has drifted from the published copy, and submits when ready. Alongside
`TabDrafts` itself the folder holds that flow's dialog windows (`NewSystemWindow`,
`NewSystemMaintainerWindow` - the maintainer agreement "Create system" must pass through,
`SubmitDraftWindow`, `DraftDriftWindow`, `MySubmissionsWindow`, `SystemFilesWindow`,
`DiscardDraftWindow`, `ForgetSubmissionWindow`), the same way `Worklog/` holds the worklog's.

**The Drafts tab carries a BADGE** (owner request, 2026-09-30: "Like the "Maintainer" tab now has a
badge, then I think the "Draft" also should have same functionality"; `Main.SubmissionChecks.cs`):
`SubmissionReceiptStore.UnreadCommentCount`, the SAME number as the "My submissions" button inside
it, drawn from `ApplyDraftsTabVisibility` - where every receipt or draft change already ends up.
Sent submissions are asked about at launch AND every `SubmissionStatusRefresh.PeriodicInterval`
(1 minute, the Maintainer tab's own - owner decision, 2026-10-01, over the earlier 5) while the
window is not minimised and something is still open (`ChecksPeriodically`); each check is followed
by the launch check's own path (retire, then `ApplyDraftsTabVisibility`) on `onFinished`, never
`onChanged` - retirement depends on a state, and a receipt published before its board synced never
moves again.

**What was just sent cannot be sent again** (owner request, 2026-10-03; cases agreed). The receipt
keeps `SubmissionReceipt.DraftFingerprint` - CRT.Data's `DraftFingerprint` of the draft as it was
sent: the workbook's VALUES (not its bytes, so an Excel save with nothing changed, or a change
undone, gives the same one), the JSON sidecar's bytes, and every file in the draft folder by path
and content (never the downloaded data, Excel's `~$` owner file, the marker or hidden files). While
the draft gives the same fingerprint as its LATEST submission, Submit is greyed out with the reason
(`SubmissionReceiptPresenter.IsAlreadySent` / `DescribeAlreadySent`; `ToolTip.ShowOnDisabled`),
whatever that submission's state - unless it never finished sending (no state yet, `uploading`,
`abandoned`). An empty fingerprint (an older receipt, an unreadable draft) never blocks.
`TabDrafts.SubmitAsync` checks again after settling the table's edits, since the row can lag an
Excel edit.

**A submission the server no longer knows reads "No longer on the server"** (2026-10-04, with the
reset at go-live). The server answers 404 for one deleted with its system or by Account > "Reset
contribution data", and the receipt is marked (`SubmissionReceipt.NotFoundUtc`). It used to keep
showing its last state for ever. `SubmissionReceiptPresenter.DescribeReceiptState`/`ClassifyReceipt`/
`DescribeReceiptBetaTry` read the RECEIPT (the Drafts badge, "My submissions"), in the neutral
colour; such a receipt never blocks Submit (`IsAlreadySent`), offers no "try it in BETA", and is
not reported on a discard (`DraftDiscardContract.WhichToReport`).

**"Save to draft" in the Contribute tab's component editor lands on the Drafts tab** (owner
request, 2026-09-24): `ComponentContributionWindow.SetAfterSaved`, which Main sets to close the
window and call `SwitchToDraftsTab` - only after a SUCCESSFUL save, so a refused one keeps the
window and its message. The board refresh (which also shows the Drafts tab for a first draft)
runs before it.

**While the table is on screen it watches its draft file** (`BoardTableEditor.FileWatch.cs`,
owner request, 2026-09-24): every 2 s, `DraftTableSession.CheckFile` - one stat call unless
the file's time or size moved, and judged against the session's CURRENT fingerprint, so the table's
own save never counts. Changed and nothing unsaved: reload in place. Changed with unsaved edits: the
`ChangedOnDiskBar` (with Reload) and "Save changes" OFF. Never reloads while a cell is being edited
(the typing is not in the model yet) or a row dragged. The reload raises `ReloadedFromOutside`,
which `TabDrafts` answers exactly like `Saved` - without it the draft row's "N rows changed" and the
board on screen stayed stale (seen by the project owner). `OpenElsewhereBar` shows the WHOLE time Excel's `~$`
owner file or LibreOffice's `.~lock.<name>#` sits beside the workbook, set the moment a table is
read (`ShowWhetherOpenElsewhere` in `Attach`) - the owner's choice after a round where it only
showed with unsaved edits, since "edit in one place at a time" matters most before editing starts.
"Save changes" is OFF while the workbook is REALLY held open (`DraftTableSession.IsHeldOpen`: an
exclusive open that fails while another program has it, probed only while a lock file exists) -
not on the lock file alone, which a crashed Excel leaves behind and which would then block the
table for good. The close prompt then offers no Save either (`DraftOpenElsewhere`), for the same
reason `DraftChangedOnDisk` offers none. File sharing is advisory outside Windows, so there nothing
is ever held and a save goes ahead; the held tests `Assert.SkipUnless` Windows. The timer runs only while the editor is attached and
holds a table; it replaced `CatchUpWithOutsideChanges`, which checked only on returning to the tab.

**No other editor writes a draft under a table with unsaved edits** (owner design,
2026-09-24). The Contribute window's "Save to draft" and the label editor's save both ask
`TabDrafts.HasUnsavedTableEditsFor(excelDataFile)` first; when the table is open on that board
with unsaved edits they save NOTHING and show the `SavingElsewhere` notice (Cancel alone) sending
the contributor to the Drafts tab. Writing anyway made the table's own save refused and its edits
lost. The project owner rejected offering "save the table first" from there: from another tab you may
not remember what you did in the table. Excel is the one writer that cannot be held back - for it
the refusal, Reload and the `DraftChangedOnDisk` prompt remain.

**"Edit in table format" (2026-09-24) opens a draft's workbook as editable sheets** directly below
its row, hiding the other drafts (`TabDrafts.Table.cs` owns table mode; `BoardTableEditor` is the
control; `UnsavedTableEditsWindow` the prompt). **Both live in `src/CRT.App/Controls/BoardTable/`**,
shared with the Maintainer tab's table (see "Maintainer" below), with their colours in
`BoardTableColors.axaml` - so the editor's files named below are there, and a change to it reaches
the Maintainer tab's table too. Colours are the project owner's: green added, orange
modified (the changed CELL only, published value in its tooltip), red + strikethrough deleted,
shown WHERE THE ROW USED TO BE. Things to know before touching it:

- **The control only paints.** Every rule is in `CRT.Data` - `BoardTableDocument`/`BoardTableSheet`
  (pairing, colours, ghost placement), `DraftTableSession` (open/save), `BoardTableClipboard` - because
  the Maintainer tab shows the same table. Logic added to the control is logic the
  Maintainer tab's host code would have to duplicate.
- **Pairing is `BoardDataDiffer`'s rule, deliberately** (`BoardDraftNaturalKeys`, keys
  case-insensitive, values trimmed + ordinal, first row per key wins). The sheet tabs' counts sit
  right under the draft row's own "N rows changed", and
  `BoardTableDocumentTests.The_change_counts_agree_with_BoardDataDiffer_on_the_board_a_save_writes`
  holds them equal. **A key edit with NOTHING ELSE changed is ONE row modified** (owner decision,
  2026-10-04 - a credit's "Name or handle" given " 2" showed a red ghost and a green row: "This
  seems weird to me"): `BoardDataDiffer.PairRenamedRows` pairs a row that lost its key with one that
  gained one when every non-identifying cell is equal (DifferingFields' rule) and something filled in
  is shared; the most shared key cells win, ties by key, so the table (screen order) and the saved
  board pair alike. The table, the differ and the server's `ReviewSummary` (as `Renamed`) all call
  it. A key AND another cell changed is still an add plus a deleted ghost - the owner chose this
  strict rule over "mostly the same row". `BoardDraftNaturalKeys.PropertiesOf` names each type's key
  properties, held to `ForRow` by a test. Rows identical but for their key DO pair (C10 deleted and
  an identical C51 added reads as a rename) - test fixtures giving every component the same values
  had to be made to differ.
- **A save replaces all nine sheets, so it may only land on the file it was read from.**
  `DraftWorkbookStore.EditIfUnchanged` compares a SHA-256 of the workbook taken BEFORE the read;
  anything else is `ChangedOnDisk` and the table offers Reload. Never route the table's save through
  plain `Edit`: that would silently undo an Excel edit or a label-editor save made meanwhile.
- **The grid is ProDataGrid** (MIT fork of Avalonia's DataGrid, which is deprecated in 12; the
  maintained TreeDataGrid needs a paid Avalonia Accelerate licence, not an option for this GPL
  project). Its assembly keeps the name `Avalonia.Controls.DataGrid`, so its theme is included in
  `App.axaml` as `avares://Avalonia.Controls.DataGrid/Themes/Fluent.v2.xaml`. Cell colours come from a
  per-column `CellTheme` whose `Background` BINDS to the cell's state - never set a cell's colour in
  code, since containers are recycled while scrolling. `EditTriggers` is set explicitly (double-click,
  typing, F2): the grid's default edits on a single click, which makes "select, then Ctrl+V" impossible.
- **The table must never sit inside the list's ScrollViewer** - a grid measured with unlimited
  height realises every row. `DraftsBodyGrid` holds them as siblings, and in table mode the list is
  hidden and the open draft's row shows alone in `OpenDraftRow` (the row template is the shared
  `DraftRowTemplate` resource). `ApplyTableMode` changes the two rows' `Height`s IN PLACE:
  assigning a new `RowDefinitions` collection left the Grid measuring against the old row kinds and
  drew the open row at zero height - pinned by
  `TabDraftsTests.In_table_mode_the_open_drafts_row_is_really_drawn_above_the_table`.
- **Style selectors match exact types.** The grid draws cell text with `DataGridSearchTextBlock`, a
  `TextBlock` subclass, so the deleted-row strike-through needs `:is(TextBlock)`; a plain `TextBlock`
  selector silently matched nothing while the row class looked right.
- **Data columns set `IsReadOnly = false` explicitly.** Left unset, the grid infers read-only-ness
  from the binding path, and `Cells[i]` indexes an `IReadOnlyList` - so every column came out
  read-only and nothing at all could be typed (reported by the project owner). Tests that set cell text
  through the MODEL cannot see this; `BoardTableEditorTests` now types with real key input too.
- **The table owns the keyboard while it is open.** The always-on component filter pulls focus back
  after every click in the window (`Main.ShouldReturnFocusToComponentSearch`, extracted from that
  pointer-release handler so it could be tested), and a click on a grid cell focuses the grid - not
  a TextBox - so typing went into the filter. It now backs off on the Drafts tab while a table is
  open, and `TabDrafts.FocusTableIfOpen` puts focus in the grid when a table opens or the tab is
  shown again.
- **The sheets are real `TabItem`s in a `TabControl`** (they were buttons at first), with no content
  of their own - the one grid shows the selected sheet. Their font and padding repeat Main.axaml's
  `TabItem` setters, scoped to the control, so they match wherever the editor is hosted.
- **Rows move through the TABLE, never the grid - and the drag is the EDITOR'S OWN, with a
  placeholder** (2026-09-24, "like moving an image in the worklog"). ProDataGrid's row drag is OFF:
  it only draws a line and moves the row once on the drop, so it cannot show a travelling slot.
  `BoardTableEditor.RowDrag.cs` does it the worklog's way: pressing the grip arms, 4px of movement
  starts `BoardTableSheet.BeginRowDrag`, and the row then moves LIVE onto the row under the pointer
  (`BoardTableRowDrag.MoveOnto`, or `Step` past the top/bottom edge), drawn as a red dashed slot -
  the row's `BackgroundRectangle` restyled, its cells faded out, the class following the ROW across
  recycled containers. The GRID captures the pointer, not the pressed header, whose container can
  change rows mid-drag. **The target is read off a FROZEN layout** (`CaptureRowSlots`, the worklog's
  `CapturePhotoRowBoundaries` idea): the first version hit-tested the LIVE grid and scrolled the
  moved row into view, and over the half-visible bottom row that scroll slid a new row under a still
  pointer - the row ran away or flickered (reported). Only a step past the edge scrolls, and the
  slots are captured again after it and after a mouse-wheel scroll;
  `Holding_a_dragged_row_still_over_the_half_visible_bottom_row_does_not_run_away` fails against the
  live version. The rules are in CRT.Data: a red ghost is never a drop target (the refresh
  would put it straight back and the row would flicker), a whole drag is ONE undo step (the
  history's group), and a drag back to its start leaves no step and no "moved by hand" mark. There
  are no Move up / Move down buttons any more (owner request); Alt+Up / Alt+Down remain. The
  grid template hides the grip; the editor's styles show it, with the `RowsDraggable` CLASS in the
  selector - a trigger is needed to outrank the template's Template-priority value.
- **The unsaved-edits prompt has THREE wordings** (`UnsavedTableEditsPrompt`): Leaving (Save /
  Discard / Cancel), Reloading, and DraftChangedOnDisk - leaving a table whose draft file changed
  after it was read (typically "Save to draft" in the Contribute tab). That one offers NO Save: the
  save would be refused, and offering it made "Close table" a loop (reported). Asked from
  `TabDrafts.AskAboutUnsavedTableEditsAsync`, which checks `BoardTableEditor.HasDraftChangedOnDisk`.
- **Pointing at a FILE cell opens a hover card AT ONCE** (owner requests, 2026-09-26,
  `BoardTableEditor.FilePreview.cs`): the picture, the published and the new one side by side when
  it changed (by BYTES too - a picture replaced under its own name is not coloured), or a link
  that opens a PDF. It follows the pointer from file cell to file cell and closes the moment the
  pointer is on neither its cell nor the card. It is a `Popup` in the window's OVERLAY layer with
  light dismiss OFF, beside the cell: a flyout closed only ~100 px away from itself (read as a
  delay), a light-dismissed popup spent the next grid click on closing, and a card UNDER the cell
  covered the next row's file. `BoardTableFilePreviewTests` drives all of it with a real pointer.
  A side is labelled ("Published" / "Your draft", "Before (published)" / "After (submitted)", and
  for a new system in the Maintainer tab "As submitted" / "Your change" - both of whose sides are
  read from the SUBMISSION, `ReviewTableFiles.HashToRead`, since nothing of it is published) only
  beside another one, where it says which is which - a picture on its own is unlabelled (owner
  request, 2026-09-26).
  **Every cell's TEXT tooltip is instant too** (the value it replaced, a problem the checks found, the
  marker's meaning): `ToolTip.ShowDelay` is `BoardTableEditor.CellToolTipDelay` (0) in both cell
  themes, pinned by `BoardTableEditorTests.A_cells_text_tooltip_opens_the_moment_the_pointer_is_on_it`.
  **And the pointer goes THROUGH it** (owner report, 2026-10-02: "the mouse should follow the cell
  below, and not the tooltip"): a cell's tooltip opens under the cell, over the next row, and
  Avalonia keeps a tooltip open while the pointer is on it - so the next row could not be reached.
  The cell themes (and the row grip's style) set `ToolTip.ShouldUseOverlayLayer`, and the editor's
  `ToolTip` style makes it not hit-testable: a tooltip that is a window of its own cannot let the
  pointer through. Pinned by `Moving_down_from_a_cell_with_a_tooltip_reaches_the_cell_below_it`,
  which fails with the style off. Which cells are files is
  CRT.Data's `BoardTableFileCells`; the bytes are the HOST's `IBoardTableFileSource`
  (`DraftTableFileSource` here, `ReviewTableFileSource` in the Maintainer tab) - a host that sets
  none gets no card. A file column's text tooltip gives way to the card.
- **ONE "Insert row" button** (below the selected row; `BoardTableSheet.InsertRow`) - "Insert row
  above" went on 2026-10-02 (owner request: "that can be moved afterwards"); the model keeps
  `InsertRowAbove`. **The legend pills have three looks** (owner request, 2026-10-03): counting
  nothing, OUTLINED in the kind's `BoardTable_*_Edge` colour with no fill, at 0.4 opacity (the
  `Empty` class); counting rows, FILLED with the kind's wash at full strength; PICKED, the 2px
  outline (`Selected`), at full strength even at 0. Pills with rows were dimmed too until picked,
  briefly, and were too hard to read (owner request, 2026-10-03). A pill counting nothing cannot be picked (no hand
  cursor; `BoardTableDocument.CanToggle`), but a picked one whose count drops to 0 can be put back. Each pill names its kind by class (`Added` ...
  `Warnings`) so the fill is a style's - a `Background` set on the pill would outrank `Empty`'s.
- **THE CHECKS ARE IN THE TABLE** (owner request, 2026-10-02: "integrate the existing validation check
  into the "Draft" system table view ... All error should be fixed before submission can be done").
  ONE rule set, CRT.Data's `BoardDataChecks`: the server's `SubmissionValidator` row rules (moved
  there, run in `BoardCheckScope.Submission` - same codes, subjects, messages and order) plus, in
  `Everything`, the server's file-name rules (`SubmissionPathRules.IsSafelyShaped`,
  `SubmissionFileRules.TryCheckName` - CALLED, not restated) and the WARNINGS `DataValidator` used to
  log alone. **An error in the table is a refusal at submit** -
  `BoardDataChecksTests.Every_error_the_table_finds_is_one_the_server_refuses` fails otherwise; a new
  local rule is a warning. Each problem carries its sheet, ENTRY index and column, and
  `BoardTableDocument.RefreshProblems` (after every sheet refresh - the rules look across sheets)
  puts it on the cell; a problem in no row (a highlight) is `ProblemsOutsideSheets`, a line above the
  table. Files are looked for by the host's `IBoardFileLookup`: `DiskFileLookup` (the draft folder,
  then the downloaded data held to EXACT spelling - the server refuses a case variant of a published
  path) for the Drafts tab via `DraftTableSession.Open(..., dataRoot)`; none for the Maintainer tab.
  A problem cell's mark is a TRIANGLE in its top-left corner, red error / amber warning, drawn as ONE
  background brush (a diagonal gradient with a hard stop at half way to (`CornerMarkSize`,
  `CornerMarkSize`), the state's wash after it) - so no cell template change, and it shows on an empty
  cell. Its tooltip leads with the problems; a FILE cell's tooltip is its problems alone
  (`ProblemToolTip`), beside the file card. Submit on a draft with errors sends nothing and opens its
  table on them (`TabDrafts.StopForErrorsAsync`); warnings never block. `DataValidator`'s launch log
  stays (owner decision) and runs the same rules. **Each draft's ROW counts them too** (owner report,
  2026-10-02: "It must check for errors when creating the draft, and if the board changes
  "offline""): `DraftStatusReader.CountProblemsCached`, stamped on the workbook and sidecar like
  the change count, and the list is read again when the Drafts tab is shown or CRT's window is
  activated (`TabDrafts.OutsideChanges.cs`), which is when an Excel edit arrives.
- **A NEW DRAFT IS SEEDED FROM THE PUBLISHED FILE AS IT IS ON DISK, never from the board cache**
  (`DraftSeeder.SeedFromPublishedFile`, 2026-10-02). The Contribute window and the label editor
  passed `BoardDataReader.LoadAsync`'s CACHED board, so a published workbook edited in Excel while
  CRT ran was seeded WITHOUT the edit (owner report). Every seeding save path calls this now.
- **THE COLOUR KEY IS THE FILTER** (owner request, 2026-10-02: "remove the "Show changes only" so it
  works in the same unified way"). Five pills - Added, Modified, Deleted, Errors, Warnings (a sixth,
  Flagged, became warnings on 2026-10-03) - each a toggle (`Tapped`); picked ones show rows of ANY
  picked kind (`BoardTableRowFilter`, `BoardTableRowKinds`), outlined 2px (`Selected` class). "Show
  changes only" is `BoardTableRowFilter.Changes` (Added | Modified | Deleted). A new BLANK or
  incomplete row is always shown. Moving rows is off while anything is picked. **The pills count the WHOLE DRAFT** (owner
  decision, 2026-10-02: "Added, Modified and Deleted should work per system like Error and
  Warning"): `BoardTableDocument.AddedCount` and its four siblings, every sheet together, whichever
  is on screen; a pill is outlined only when the draft has none. A sheet's own change count is its tab's.
- **THE SEARCH BOX** (owner request, 2026-10-02: "the same search/filter as we have in the
  "Workbooks" tab, where it highlights what I search for"; `BoardTableEditor.Search.cs`). The rule is
  CRT.Data's `BoardTableSearch` over the Workbooks tab's own `WorklogSearchQuery` (moved to CRT.Data
  for it): words ANDed, each within one cell, quotes, -minus, any case - and EVERY cell, numbers
  included (unlike Workbooks). **Decided once, when typed**: the rows not matching THEN are hidden,
  so a row edited out of the match stays until the search changes (agreed case C) - a save or Reload
  reads new row objects and works it out again. It combines with the pills (a row both show; the
  tabs of sheets with none hidden; `BoardTableDocument.SheetsShown(kinds, search, current)`).
  **A search finding nothing in ANY sheet keeps only the sheet on screen**, empty, with
  `BoardTableSearch.NothingFoundLine` above the table (owner request, 2026-10-03 - every tab stayed,
  which read as a search of one sheet); picked pills alone showing nothing still keep every tab. It
  stops row moves, and is emptied by `Clear` and by `Load` of ANOTHER draft only (agreed case D). **The
  marks are ProDataGrid's own**: its cell text block marks the runs in the grid's search model's
  results, which the editor computes and hands over through `BoardTableSearchAdapter` - the stock
  adapter recomputes them with its own grammar on every view refresh and WIPED them (a spike showed
  it), so the subclass answers every recompute with the editor's results. Three traps, each found
  by a render: the model's `HighlightMode` must be `TextAndCell` (the default marks no text), the
  mark colours are written into the editor's own `Resources` in code from `Workbooks_SearchHit_*`
  (the grid's control lookup does not reach a theme-dictionary key - in a theme dictionary they came
  out in its default blue), and the grid's row tint (`DataGridRowSearchMatchOpacity`) is set to 0
  and a matching cell's `Foreground` bound back to its row's (it gave the whole cell the mark's text
  colour, unreadable in dark).
- **SEVERAL ROWS DELETED IN ONE GO** (owner request, 2026-10-02: "mark multiple rows and then
  delete those in one go"; `BoardTableEditor.Selection.cs`). The grid is `SelectionMode="Extended"`
  over cells: Shift+click a range, Ctrl+click one in or out; a row is selected when any cell of it
  is (`SelectedRows()`, sheet order, rows on screen only). "Delete row" calls
  `BoardTableSheet.DeleteRows` - ONE undo step, ghosts skipped, a component's rows on other sheets
  taken for all of them (all selected rows are removed FIRST, so two regional twins deleted together
  count as both gone), `BoardTableDeletedWith.Combine` saying it in one sentence - and reads "Delete
  8 rows". **`SelectCell` makes its cell the whole selection**, so a selection left behind by an
  undo, a delete or Tab is never what the next delete takes. Several selected cells carry a
  translucent veil (`BoardTable_Selection_Fill`, the cell template's `CurrencyVisual` filled, under
  the grid's `SeveralSelected` class) over their own colour; one cell keeps the dashed frame alone.
  **The Delete KEY deletes no row** (owner answer: "No") - `CanUserDeleteRows="False"`: it defaulted
  ON, and Delete on a selected cell took the row out of the sheet's list behind the model's back (no
  red row, no undo step) until this change.
- **Deleting a component deletes everything that is its own** (owner, 2026-09-25): its rows on
  the image, local file and link sheets at once (`BoardTableDocument.DeleteRowsOfComponent`, shown
  red there), its highlights at save (`ApplyTo`), all ONE undo step - `BoardTableHistory` steps span
  sheets for this. A RENAME keeps them: the document remembers the component rows it opened with,
  and a row object still live under a new label is a rename. The status line says what else went
  (`BoardTableDeletedWith`). The Contribute window's "Delete this component" drops highlights too.
- **Row order is what the user sees, and new components are PLACED (2026-09-24).** The main
  window's component list, category list and Overview all show the Components sheet in its own
  order. `ComponentPlacement` is the one rule: a NEW component goes into its category in natural
  label order (C1, C2, C10), a new category goes last, and an EDITED component stays put. The
  Contribute window's save (`ComponentBoardWriter`) used to remove and re-append a component at
  the bottom on every save - so editing C1 after adding C2 put C2 first, which was reported. The
  label editor's new components and the table's inserted rows (on save, unless moved by hand)
  follow the same rule. Existing rows are never re-sorted.
- **A component row's identity is its label PLUS its region** (`BoardDraftNaturalKeys.ForComponent`,
  2026-09-24) - a regionalised component is one row per region. It used to be the label alone, so
  an added U1/NTSC beside U1/PAL was flagged a duplicate in the table, counted nowhere, and never
  shown to a maintainer. Rows with no region key exactly as before. Giving a component a region
  changes its key, like a label change - one row modified when nothing else changed (2026-10-04,
  pinned by `ReviewFieldDiffTests`). **An IMPORTANT SIGNAL's identity is its display name PLUS its KiCad net**
  (`ForKiCadImportantSignal`, 2026-09-26): one display name covers several nets ("9VAC" is the 9VAC
  and the 9VAC~ net - the KiCad Wiki page says rows may share a display name), and keyed on the
  name alone every second one was flagged a duplicate (reported from the maintainer's table). A
  signal therefore has no non-key field: a new net is one row modified while its display name stays,
  and a removal plus an addition once neither half does. `BoardDraftSummary` takes the label as the key's FIRST part for the Draft chips.
  `ComponentPlacement` puts a new regional variant straight after its twin (a blank category used
  to send it to the end of the sheet, where the project owner thought it had vanished).
- **Excel keys:** Tab / Shift+Tab move right/left and wrap rows (handled on the tunnel route, since
  left alone Tab is the window's focus navigation and left the table); Enter moves down (the grid's
  own). The pill filter goes through a `DataGridCollectionView` over the sheet's rows and
  switches moving off. **The grid's `FilteringModel.OwnsViewFilter` is set FALSE** - left true,
  the grid writes its own empty predicate over the view's `Filter` whenever it takes the view or
  is re-attached, so the filter (then "Show changes only") was lost with the box still ticked on
  every sheet switch and then on every MAIN tab switch (both reported). A table the user's pick would
  show NOTHING of opens unfiltered (`document.HasRowsShownBy`): a filter left on from another draft
  (the editor is reused) hid every row - reported as an "empty" Board schematics sheet, a draft with
  nothing published. **The user's PICK is kept apart from the filter** (`thisFilterWanted`,
  `FilterWanted`), so the next table with such rows has it back on; `ApplyFilter` changes the filter
  without touching the pick. **A pick also HIDES the tabs of sheets it would show nothing of**
  (owner request, 2026-09-26) - `BoardTableDocument.SheetsShown` (the current sheet
  keeps its tab while on screen; nothing to show anywhere keeps every tab) and `SheetToShow` (the
  sheet to open on when the wanted one's tab is hidden). The grid's scroll bars stay FULL SIZE (`ScrollViewer.AllowAutoHide="False"` on it - it scrolls
  through a ScrollViewer of its own; owner request, 2026-09-26): the theme's hairline hid that a
  wide sheet has columns past the right edge. The header row is set apart WITHOUT weight (owner:
  a black band "steals way too much focus"): semi-bold names on a soft band and a firm 2px line
  under the row (`BoardTable_Header_*`, the table's own keys - not CRT's `Table_Header_Bg`, which
  also colours the Contribute window's and About tab's tables). **A value the grid's TEMPLATE sets
  outranks a plain style** (Template priority is above Style), so the line under the header is
  styled through the grid's `HeaderStyled` class (a class makes the style a trigger), and the
  column separators through the header's `SeparatorBrush`. Also: the corner above the grips draws
  its own `TopLeftHeaderRoot` Grid, and CRT's app-wide `TextBlock` style outranks the header's
  inherited text colour. All pinned by `The_header_row_is_set_apart_from_corner_to_corner`.
  **Every key the table's markup names must resolve in CRT's App.axaml** - `SharedTableColourKeysTests`, light and dark,
  fails on a key App.axaml does not define (a missing DynamicResource draws nothing, silently). The text size is set on EVERY column (`CellFontSize`) - a text column carries
  its own size, so the grid's `FontSize` alone shrank only the headers - and cells get `MinHeight`
  `MinRowHeight` (26, the `CellsWrap` style) instead of the theme's taller default, so the denser rows
  do not clip the current cell's frame. **Every cell wraps and a row is as tall as its tallest cell**
  (owner request, 2026-10-04 - a "Wrap text" check box, remembered in `UserSettings`, came and went
  the same day), and **a double-click on a heading's edge fits the column to the widest text in every
  row shown**, never wider than the table (`BoardTableEditor.TextWrap.cs`, the rule
  `ColumnAutoFitGeometry`) - the grid's own double-click fit is off because it measures only the rows
  it has built.
- **Order is NOT a change `BoardDataDiffer` counts** (it pairs by key). A reorder is saved into the
  draft and published with a submission - `PublishMerge` takes rows as submitted - but a draft whose
  ONLY change is a new order reads "0 rows changed" and cannot be submitted, and a maintainer is not
  shown it. Making order count would touch the differ the server and the Maintainer tab share; that is a
  owner decision, not yet taken.
- **No "Restore row" / "Revert cell" buttons** (removed at the owner's request, 2026-09-24,
  once undo existed). Undo only reaches back to the last save, so a row deleted or a cell changed
  BEFORE it is put back by typing - the red row and the orange cell's tooltip still show the
  published values. `BoardTableSheet.RestoreRow`/`RevertCell` stay in the model, tested, for the
  Maintainer tab's table.
- **Undo and redo (Ctrl+Z / Ctrl+Y, the platform's own gestures) live in the model**, in
  `BoardTableHistory` (`BoardTableDocument.History`). Each step is a SNAPSHOT of one sheet's live
  rows - the row OBJECTS, their values and their two placement flags - taken just before the
  change, and undo puts it back and lets `Refresh` rebuild colours and ghosts. So every change
  kind is undone by the same code, and **a new kind of change is undoable only if it calls
  `History.Record` before it mutates** - the cell `Text` setter and every row operation do. The
  history dies with the document, so it reaches back to the last SAVE (a save reloads). While a
  cell is being edited, Ctrl+Z is its text box's own. Undoing back to the saved state clears
  "unsaved" again (`IsAtSavedState`).
- **The row and cell buttons hand focus back to the grid** (`ThenFocusGrid`). "Delete row" leaves
  the cursor on the red ghost, which disables the button, and a focused button that disables
  drops the focus to NOTHING - so Ctrl+Z straight after reached no handler at all. Undo is also
  handled on the whole editor, not just the grid, for focus on the check box beside it.
- **The current cell has no fill and a 2px dashed red frame** (owner request). The grid
  theme's selected fill looked like "added" and covered the cell's own state colour, so the cell
  theme re-binds `Background` in a `^:selected` trigger of its own (added after the grid
  theme's, so it wins). The frame is the template's `CurrencyVisual` RECTANGLE restyled - a
  `Rectangle` can dash its stroke, a `Border` cannot - and the grid's `FocusVisual`, its
  selection overlay's `PART_SelectionOutline` and the `PART_FillHandle` (Excel's drag-to-fill,
  which writes a run of cells the table was never built for) are hidden. The drag grip spans
  both of the row header's columns (`Grid.ColumnSpan`), or it sits centred in the first, 16px one.
- **Three colours; a duplicate or a row the save drops is a WARNING** (owner request, 2026-10-03:
  "Should flagged now be treated as warnings?"; cases agreed). They were a fourth colour, violet
  "Flagged", with `!` and a pill of their own, until then - the colours now mean only what changed.
  A row whose natural key repeats one above is `BoardDataChecks`' `row.duplicate` warning on EVERY
  row of the set, on the first identity column (`BoardDraftNaturalKeys.ColumnsOf`, held to `ForRow`
  by a test), in `Everything` only - except on Components and Board schematics, where a duplicate
  is already the server's error - refused on the later row(s), and marked by the TABLE on every row
  of the set (owner decision, 2026-10-03: "it should show all rows, and not only last, as the
  maintainer/contributor then has some context"), the first one in `Everything` only, so what the
  server refuses is unchanged. A row the save drops (an Important signals
  row missing a half - `BoardWorkbookSchema.RequiredColumns`, held to the mapper by a test) is
  `row.incomplete`, put on by `BoardTableDocument.RefreshProblems` itself, since the checks only see
  rows a save writes; it is coloured like a new row, marked `+`, and always shown by a filter. A
  later duplicate is uncoloured and unmarked. Neither is a change - `ChangeCount` and the tab's
  number leave them out, agreeing with `BoardDataDiffer`, and the first row of a duplicate still
  counts as whatever pairing made it. A remembered "Flagged" filter parses as Warnings. Each pill holds its count AND its word ("[ 2 Added ]"): a count badge beside a
  separate word read as belonging to either neighbour (reported). **A sheet TAB says its name and
  change count and nothing else** (owner request, 2026-10-02: "the tabs ... should not show "2
  flagged" and "8 error" - the tabs will be obvious when you click those badges"): the flagged
  (2026-09-26), error and warning pills it carried are gone with `BoardTableSheetTabHeader`, since
  picking a pill hides every tab without such rows. A tab's header is that plain string.

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
  [Handlers/Data/WorkbookSummary.cs](../src/CRT.App/Handlers/Data/WorkbookSummary.cs) (pure, unit tested), which
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
  ([DeleteWorklogWindow](../src/CRT.App/Tabs/Worklog/DeleteWorklogWindow.axaml.cs) — a near-copy of
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
  [WorkbookExportModel](../src/CRT.App/Handlers/Data/WorkbookExportModel.cs) decides WHAT goes in and in what
  order (grouped per schematic, entries by id, missing attachment files dropped, an entry with no
  schematic filed under `(no schematic)` rather than silently lost) and is unit tested;
  [WorkbookPdfExporter](../src/CRT.App/Handlers/Data/WorkbookPdfExporter.cs) only paints it, and its LAYOUT is
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
  PROPORTIONAL: [ExportOverlayGeometry](../src/CRT.App/Handlers/Geometry/ExportOverlayGeometry.cs) turns a pixel
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
  The query language lives in [src/CRT.Data/WorklogSearchQuery.cs](../src/CRT.Data/WorklogSearchQuery.cs)
  (pure, unit tested by `WorklogSearchQueryTests` in CRT.Data.Tests; moved from CRT.App on 2026-10-02
  so the board table's search box uses the same grammar - `BoardTableSearch`): space-separated terms are ANDed, `"a phrase"`
  quotes a run containing spaces, a leading `-` excludes, matching is case-insensitive substring
  throughout (so `p c u` finds `CPU`, and `"full text"` finds `Afull textB`). An empty box is not a
  filter and matches everything. Which fields are searched is
  [WorklogSearchIndex](../src/CRT.App/Handlers/Data/WorklogSearchIndex.cs): every user-typed TEXT field on the
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
[Handlers/Theme/WorklogInfoPillBuilder.cs](../src/CRT.App/Handlers/Theme/WorklogInfoPillBuilder.cs): a `Form_Bg`
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
why [WorklogAttachCaptureWindow](../src/CRT.App/Tabs/Worklog/WorklogAttachCaptureWindow.axaml.cs) is ONE modal
(image, workbook, worklog, comment) rather than a picker followed by the editor's Add photo dialog.
It resolves the choice and returns it; the caller performs the write, the same division
`WorklogAddPhotoWindow` keeps. The popup asks `Main.ResolveActiveWorkbookForBoard` (made `internal`
for this) rather than re-deriving "which workbook" a third time.

**The entry list is RANKED, and the first row is preselected** -
[WorklogAttachTargets](../src/CRT.App/Handlers/Data/WorklogAttachTargets.cs) (pure, unit tested). The rule is
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
[WorklogDefaultAreaGeometry](../src/CRT.App/Handlers/Geometry/WorklogDefaultAreaGeometry.cs) (pure, unit tested),
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
[WorklogAttachmentWriter](../src/CRT.App/Handlers/Data/WorklogAttachmentWriter.cs), which the editor's own Photos
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
[Tabs/Schematics/TabSchematics.axaml.cs](../src/CRT.App/Tabs/Schematics/TabSchematics.axaml.cs):

| File | Owns |
| --- | --- |
| `TabSchematics.axaml.cs` | Construction, one-time wiring in `Initialize`, fullscreen/splitter layout, the user-drawn trace colour palette, shared parse/theme helpers |
| `TabSchematics.Types.cs` | Private data types shared by the other parts (cache records, hover candidates, undo state, `EditableComponentHighlight`) |
| `TabSchematics.Viewport.cs` | Zoom, pan, the transform matrix, matrix clamping, content/viewport rects |
| `TabSchematics.Input.cs` | Pointer, wheel, gesture and keyboard handlers — these only dispatch |
| `TabSchematics.Thumbnails.cs` | Thumbnail list, selection, thumbnail bitmaps, drag-to-reorder |
| `TabSchematics.ThumbnailsDetach.cs` | Detach-to-window mode: collapses/restores the inline thumbnail column, mirroring `EnterFullscreenMode`/`ExitFullscreenMode`. The detached window and its own 2D gallery/drag logic live in `SchematicsThumbnailGallery`/`SchematicsThumbnailsWindow`/`ThumbnailGalleryPanel` |
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
`ComponentLabelEditorOverlay`, `IcTestPanel`, `SchematicsThumbnailGallery` (the detached-window
thumbnail gallery), `SchematicsThumbnailsWindow` (the window that hosts it), and
`ThumbnailGalleryPanel` (its auto-fit layout panel — the maths itself is
`Handlers/Geometry/ThumbnailGalleryGeometry`).

#### Maintainer (`Tabs/Maintainer/`)

**The role it serves is "maintainer"** (owner decision, 2026-09-25). The tab was the separate
"CRT Maintainer" application until 2026-09-29, and "Review" / "reviewer" before that; everything was
renamed before the app's first release -
code, the API, the database (migration 0010 renames the `reviewers` table back to `maintainers` and
the stored approval roles; `MaintainerRoleMigrationTests` holds the code's role words to it) and
every text. **Two words, two meanings:** "maintainer" is ONLY the role - a person in a system's
pool - and the person who owns this project is "the project owner" in every document and comment.
The ACTIVITY is still "review": a maintainer reviews a submission, the queue is the review queue,
and names like `ReviewEndpoints`, `/api/review/...` and `ReviewSummary` stay as they are.

**A TAB IN CRT SINCE 2026-09-29, NOT AN APPLICATION** (owner decision - Phase 8 of
NewContributeStrategy.md, and `Assets/MaintainerTabMergePlan.md`). It was a separate Avalonia desktop
app (Phase 5) with its own version, workflow and release repository; its code now lives in
`src/CRT.App/Tabs/Maintainer/` (UI, namespace `CRT`) and `src/CRT.App/Handlers/Maintainer/` (logic,
namespace `Handlers.MaintainerHandling`), its tests in `tests/CRT.App.Tests/Maintainer/` and
`tests/CRT.App.Tests/Ui/Maintainer/`. What the window did and a tab cannot, and where each went:

- **Shown only when "Enable Maintainer tab" is ticked** in the Configuration tab
  (`UserSettings.EnableMaintainerTab`, default OFF - it needs a maintainer account), between Drafts
  and Configuration. `Main.Maintainer.cs` owns that and the LAYOUT: while the tab is selected the
  sidebar (columns 0 and 1 of `RootGrid`) and the worklog bar collapse, and leaving it puts back the
  width that was there. `OnMainSplitterPointerReleased` never saves `LeftPanelWidth` while collapsed
  (it would save 0). Hiding the tab never signs out or closes its table.
- **The remembered sign-in is restored AT LAUNCH, quietly, for the tab's badge** (owner request,
  2026-09-30; `TabMaintainer.RestoreInBackgroundAsync`, started by `Main.StartMaintainerBadge` from
  `StartAsync` once the window is up, or when the setting is ticked later) - no overlay, and only the
  badge's two lists. Failing that, on the tab's first attach, under the overlay as before. NEVER in
  the constructor. `ReviewSessionStore.Initialise()` runs at CRT's start-up;
  no test runs that, so a tab a test builds finds no stored session. The token is DPAPI-protected on
  Windows and not stored anywhere else (`ReviewSessionProtection` - its entropy string still says
  "Review" and must not change). See `ReviewSession`'s header for the API's `refreshToken` naming
  trap: that field IS the bearer token, and there is no access-token exchange to go looking for.
- **The tab carries a BADGE in CRT's row of tabs** (owner request, 2026-09-30: "until it is fully
  processed, including if it is awaiting in "BETA to PROD" queue"): the "Contributor Submissions"
  and "Beta > Prod" attention badges ADDED UP (`MaintainerModes.TabAttention` - a sum, so the tab
  says what opening it shows), drawn by `Main.ShowMaintainerTabBadge` through
  `TabMaintainer.UseTabBadge`. A 401 clears what it counts (`ShowSignInPanel` empties
  `thisQueueEntries`), so a signed-out tab never keeps a stale number. The number is
  `Handlers/Theme/TabBadge`'s, which the Drafts tab's badge shares. **No tooltips** on it or on the
  tab's buttons (owner request, 2026-09-30: "no need to see it").
- **The minute check** (`TabMaintainer.QueueRefresh.cs`, `QueueRefreshRules.MinuteCheck`): with the
  tab on screen AND CRT's window in front, EVERYTHING, as before (a tab not on screen is DETACHED,
  which clears the window it listens to; returning checks at once if due, against when everything
  was last read - `thisEverythingAskedUtc`, not the badge check's time). Otherwise, while the badge
  can be SEEN (tab on, window not minimised - `Main.MaintainerTabBadgeCanBeSeen`), only the queue
  and the BETA list - never the Systems overview. **The price, accepted with the request:** every
  request slides the 30-day session, so a maintainer now stays signed in while CRT is used at all,
  not only while reviewing.
- **The entry "Contributor Submissions" opens on is READ AHEAD while the tab is away** (owner
  request, 2026-10-02: no "please wait" on the first look; `TabMaintainer.Prefetch.cs`, the rule
  `SubmissionPrefetch`): its detail and table, quietly, last at launch and after each off-screen
  badge check when that entry's queue row changed. Used once, only for EXACTLY the queue row it was
  read for; opening it shows it at once and re-reads both quietly (a newer table version replaces
  the one shown unless something was typed). Coming back to the tab opens the entry BEFORE the
  full check when the queue is already known.
- **Quitting CRT asks about unsaved table edits** through `HasUnsavedTableEdits` /
  `ConfirmLeavingTableAsync`, from `Main.OnWindowClosing` - after the Drafts tab's table, in ONE
  posted continuation with one settled flag.
- **No "please wait" overlay of its own**: its waits run under `Main`'s (`BusyOverlayHostsTests`
  fails if one comes back). Tests stand a Grid with one overlay in for Main (`MaintainerTabHost`).
- **The table's filter - its picked pills - is `UserSettings.MaintainerTableFilter`** (kind names,
  `BoardTableRowFilter.Format`; it was the "Show changes only" bool until 2026-10-02, and a saved
  `true` reads as `Changes`), handed over by Main through `TabMaintainer.UseRememberedChoices` and
  written back through the table editor's `FilterWantedChanged` (raised for the user's PICK only).
  The separate app's own settings file is carried over once and deleted
  (`MaintainerSettingsMigration`); its window placement is not.
- **The component filter never takes focus back over the tab** (`ShouldReturnFocusToComponentSearch`),
  and F11 does not open the schematics fullscreen while it is selected.
- **The server address is `AppConfig.CrtServerRootUrl`** (no "/api" - every review route appends the
  full path), while `CrtServerBaseUrl` is built from it WITH "/api" for `SubmissionClient`. The two
  differ on purpose; `ReviewApiRoutesTests` pins the review side. Requests send
  `User-Agent: CRT <version>`.

**"Contributor Submissions" and "Beta > Prod" OPEN ON AN ENTRY** (owner request, 2026-09-30;
`TabMaintainer.OpenOnEntry.cs`, the rule `MaintainerModes.EntryToOpen`): the one looked at last
while it is still listed (`UserSettings.MaintainerLastSubmissionId` / `MaintainerLastBetaSystemId`,
so across restarts), else the first - on choosing the screen's button, after the lists are read at
sign-in, and on coming back to the tab. Never over something still open (a submission decided
elsewhere keeps its table), and never from the badge's off-screen check.
**With nothing waiting, the tab OPENS ON SYSTEMS** (owner request, 2026-10-04: "if there is no queue
awaiting, when opening the "Maintainer" tab, then go to "Systems" and show the last selected
system"; `MaintainerModes.ScreenOnOpening`, `TabMaintainer.OpenOnEntry.cs`). Opening is showing the
tab or signing in on it; the screen is decided ONCE per opening, as soon as it is known whether
anything waits (`WaitingIsKnown`: something counted in either list settles it, "nothing" needs both
read), and nothing is opened before then. "Waiting" is the badge's own sense (`TabAttention`), so a
submission waiting only for the other approver does not count. Only a queue screen gives way, and
never with a submission or BETA system open on it; a screen chosen meanwhile settles it. Systems
then opens on `UserSettings.MaintainerLastSystemId`, else the first, once its list is read - the
Systems BUTTON still opens nothing by itself. The read-ahead (`TabMaintainer.Prefetch.cs`) skips an
opening that will go to Systems (`ScreenOnNextOpening`).
**Four screens, chosen by a TAB STRIP across the top** (owner request, 2026-09-27; a strip since
2026-10-01 - they were grey buttons wrapped onto two rows of the list column, identical to the
submission's view buttons, and the owner called it "bad UX design");
`TabMaintainer.Modes.cs`), in this order and with these labels (owner request, same day; RENAMED
2026-10-04 to "Queue: Contributor submissions", "Queue: Awaiting push from BETA to stable" and
"Administrator activities", the last renamed again the same day to "Account" - this file still says
"Contributor Submissions", "Beta > Prod" / "BETA > Stable" and "Admin" for them in older notes). **The labels are CRT.Data's
`MaintainerScreenWording`**, which the tab strip reads (x:Static) and every message naming a screen -
the tab's own, `OneSubmissionInBeta` and the server's refusals - is built from, quoted
(`*Quoted`); never write a screen's name as a literal:
**Systems** (every system, in five views - below; `TabMaintainer.Systems.cs` beside
`SystemView`), **Contributor Submissions** (the Review screen:
the queue and its table, below - `MaintainerMode.Review` and the `Review*` control names stayed),
**Beta > Prod** (systems waiting for production - `TabMaintainer.Beta.cs` beside `BetaView`, the
"Publish to production" window until then; `MaintainerMode.Beta`) and **Account** (`MaintainerMode.Account`,
`TabMaintainer.Account.cs`; "Admin", the administrator's alone, until 2026-10-04 - owner request:
"rename the "Administrator activities" tab to "Account". Then make this tab available to all
maintainers"). **Every maintainer** has its first two entries: **"My account"** (`MyAccountView` - the
name, address and password of the "Your account" window, and **Sign out**; both were buttons under the
lists until then) and "Server version". **The administrator's own** follow - "Maintainers", "Order of
systems", "Unused files", "Rebuild checksum manifests", "Delete a system", "API usage" and "Reset
contribution data", beside `MaintainerPoolView`, `SystemOrderView`, `UnusedFilesView`,
`RebuildManifestsView`, `SystemDeletionView` (whose confirmation is `DeleteSystemWindow`),
`ApiUsageView` and `DataResetView` - each with the class `AdminOnly` and the padlock (see "Code layout
conventions"), and **taken OUT of the list** for anybody the server's `isAdministrator` does not name
(`ShowAdministratorEntries`; they start out until the first queue answer). **Not `IsVisible = false`**:
the list's `VirtualizingStackPanel` sets `IsVisible` back to true on each container it realises, so a
tab never shown - every test - kept them hidden while a real window showed them all to a maintainer
(seen in a render; `In_a_shown_window_a_maintainer_is_drawn_only_the_two_entries_that_are_theirs`
fails against it). Account opens on "My account", which reads nothing. **"API usage"** and **"Reset contribution data"** (owner requests, 2026-10-04 -
see the server's bullets) read when chosen and never on the minute check; the reset is LAST in the
list, worded by `DataResetWording`, and its red button works only once `RESET` is typed and only while
the server's switch is on (`DataResetWording.CanReset`); the usage is worded and ordered by
`ApiUsageDisplay` (least-called first in each area, the never-retired routes last). **"Server version"** (owner requests,
2026-10-04 - a line above the account row until the second) is three lines, each value bold in
brackets: "Server version [4.6.0]", "API server version [1]" and "API application version [1]" -
the deployed version and API revision from `GET /api/health`, asked with no session whenever the
entry is shown (chosen, or Account shown again with it chosen), and this CRT's own revision; worded by
`ServerVersionDisplay`, the answer CRT.Data's `HealthStatus`). Sign-in opens on Contributor Submissions -
or on Systems when nothing waits (above). At 800 wide (CRT's minimum) with two-digit badges the four
tabs still do not fit on one line and "Account" wraps onto a second, inside the tab (rendered
2026-10-04); with no badges they fit. **Under the lists the tab says only who is logged in** (owner
request, 2026-10-04): "Logged in as:" and, under it, the name in bold and the address in brackets
(`MyAccountRules.LoggedInAs`) - no buttons.
**An administrator sets maintainers under ACCOUNT > MAINTAINERS** (owner request, 2026-10-04: "'Send
invitation' and 'Add as maintainer' gets moved to the 'Admin' tab ... as this is something only the
admin should be able to do" - they were on the Systems screen from 2026-09-27, and under Admin
before that). `MaintainerPoolView`: a system chosen from every system (`GET /api/admin/systems`,
`ReviewApiClient.GetSystemsAsync`, its count beside each), then a Remove button per maintainer, the
open invitations with Withdraw (the system's detail carries them), an account list to add from, and
**"invite somebody new" by email** - the server (`MaintainerInvitationFlows`, migration 0012's
`maintainer_invitations`) mails a one-time code, and the invitee redeems it with "I have an
invitation" on the tab's sign-in screen (`TabMaintainer.Invitation.cs`), which creates their account
VERIFIED and puts it in every pool that address was invited to. The invitation is shown INSTEAD of
the sign-in form, never under it, with Cancel back to the form as it was left (owner request,
2026-10-04 - the email and password boxes above it were confusing). That is the ONLY way a new
maintainer's account is made - CRT cannot register one. An address that already has an account is
refused as an invitation and chosen from the list instead (the list says why an account cannot be
granted). A system's own Maintainer view on the Systems screen LISTS its maintainers and nothing more
("The stuff that should be visible in here, is just the selected maintainer(s)"), for everybody.
**Account > "Order of systems"** (owner request, 2026-10-04: "sort the list of systems, which then gets
saved to both sources (BETA + stable) after my save"): `SystemOrderView` drags BETA's drop-down list
(`ListRowDrag`, the placement's own drag) and sends EVERY system in its order to `POST
/api/admin/systems/order` (`SystemOrderFlow`), which refuses a list not naming exactly BETA's systems
(409), writes BETA's newest main Excel data file and then the stable source's in the same order
(`MasterListing.Reorder`/`ArrangeAs` - a row the order does not name stays after the row it
followed), and rebuilds each changed tree's manifest. A stable list or manifest that could not be
written after BETA's was is said in the answer (`SystemOrderAnswer.Problem`), never rolled back.
**EVERY WRITE OF A MAIN EXCEL DATA FILE dates both sheets and freezes their panes** (owner request,
2026-10-04; `MasterListing.Finish`, called by its one `Save`): the "# Revision date:" line of
"Hardware & Board" AND "Oscilloscope" set to the day, the first frozen under its header, the second
under its header and after its second column. A sheet with no date line is not given one; a rich-text one (the shipped file's: a plain label,
the date bold) keeps its runs, only their text changing. **The Systems list flags what is off with the
drop-down lists** (owner request, 2026-10-04: "I do not expect there should be cases where something
can only be listed in stable? If so, it must be flagged"): `SystemOverviewEntry.ListedInBeta` /
`ListedInStable` (`SystemOverviewFlow.ReadListings`, null when a list could not be read) worded by
`SystemsDisplay.ListingParts` as to-do - listed only in the stable list, or missing from it while the
stable source holds the board - and the list now also holds what only the stable source has (tree or
list), which used to be missing from the screen altogether. Each is a
list on the left and the chosen item on the right. **Switching
screen HIDES, it never closes**: an open table, unsaved changes and all, is where it was on coming
back, so nothing is asked. Review and BETA carry a badge counting the SYSTEMS waiting for this account
(`MaintainerModes`, over the server's `awaitsYou`), Systems a discreet count of all systems, and
Account's administrator's entries are in its list only when the server says the account is an
administrator. The BETA list and the
drop-down listing are read on the queue's own minute check; the Systems OVERVIEW only while its
screen is shown (`QueueRefreshRules.ReadsSystemsOverview`, code review 2026-09-27 - the server walks
both data trees for it). Each keeps the rule below: nothing open is reloaded unless its row changed
(a system's view count alone is not a change, `QueueRefreshRules.SystemChanged`) - the "I have
checked this in BETA" tick survives a check. **Approve's block (`ApprovalGate`) is applied again
whenever either list it reads is replaced** (`TabMaintainer.ReapplyApprovalGate`), not only when a
submission is opened. **The
Systems screen shows every system's contributors, addresses included, to EVERY maintainer** - an
owner decision ("Everything for everyone"), enforced on the server by `SystemOverviewFlow`
(`CanReviewAnything`), not a default to tighten quietly. Its state words are CRT's own
(`SubmissionReceiptPresenter.DescribeState`), so the Drafts tab and the Maintainer tab describe a submission identically.
**A system has SIX VIEWS, like a submission's three** (owner request, 2026-10-03: "all the same
functionalities, as the 'Contributor Submissions' has"; `SystemSections`, the `SystemView.*`
partials): **Board data** (BETA's board in the shared table editor), **Files** (every file it uses
in BETA - `FileTreeView.ShowListing`: no "Show only changed files", a count of files, its own
folder open), **Contributor**, **Maintainer** (its placement while it needs one, and its maintainers
as a list), **History** (2026-10-04, below) and **Statistics** (the board views; graphs are to come).
**The History view is a CARD PER SUBMISSION** (owner request, 2026-10-04: "the full history of what
has happened with this board, in an 'easy to overview' way"; `SystemHistoryDisplay`,
`SystemView.History.cs`): by month, newest first, the day in a column of its own; a card gathers a
submission's sending, amendments and decision, its state now in CRT's words, what the contributor
was told - and **WHAT IT CHANGED IN BETA**, CRT.Data's `SubmissionChanges` on its
`SystemSubmissionEntry.Changes` (server 4.3.0). That summary can only be made AT THE PUBLISH - no
board history is kept, so the board before is gone once written over - so `ApprovePublishFlow`
compares BETA's board with the one replacing it (`ReviewSummary`, the review's own comparison: rows,
fields, highlights, calibrations; files added or replaced asked of the tree before the write, removed
ones after) and stores it (migration 0017's `submission_changes`, `ISubmissionStore.SetChangesAsync`),
never failing the publish for it. Bounded: `SubmissionChangeFacts.ListedPerKind` names per list, the
counts carry the rest. A submission published before 4.3.0 has none, and its card says so. Every
other event is a line of its own (`SystemsDisplay.HistoryWhat`). The view chosen STAYS for the next system (a submission's resets); Board data and Files
are read from the server when first shown for a system, and READ AGAIN WHEN BETA MOVES (owner report,
2026-10-04: a fix approved and promoted kept showing its warning until a restart):
`SystemOverviewEntry.BetaContentHash` (server 4.1.0, the `systems` row's `content_hash`) is held
against the entry the table was read at - `QueueRefreshRules.SystemTableChanged` for the table (also
whether it may be edited, which a promotion moves), `BetaBoardChanged` for the files - in
`SystemView.CatchUpWithBetaAsync`, on every re-read of the system's row. Never under unsaved edits
(`SystemSections.BetaMovedUnderChange` says so instead). **A Board data edit GOES STRAIGHT TO BETA**
(owner decision, 2026-10-03: "it should go directly to the next queue, 'BETA > Stable', so it can
directly be tested in BETA" - it replaced the same day's "Becomes a submission"): "Save changes"
asks the server what the change would remove (`edit/check`), then asks for a REASON in
`PublishSystemChangeWindow` beside that list (owner request: "do ask for a change reason when
clicking the 'Save changes' button" - the description box above the table is gone), then sends it to
`POST /api/review/systems/edit` with BETA's fingerprint; the server makes it a submission from the
maintainer's account and approves it at once (see the server's bullet). The table is then read again
and the BETA > Stable list and the systems re-read (`AfterSystemChangePublishedAsync`), staying on the
Systems screen. Made but not published, it opens under Contributor Submissions
(`OpenSentSubmissionAsync`). The table is the editor's READ-ONLY mode (`BoardTableEditor.IsReadOnly`,
from the server's `MayEdit`) for a system the account does not maintain AND for any system waiting
under BETA > Stable (owner decision: "it should simply disallow it, even if this is coming from a
maintainer"). Choosing another system, signing out and quitting CRT ask about a change not published
(`UnsavedTableEditsPrompt.LeavingSystem`; `TabMaintainer.HasUnsavedTableEdits` covers both tables). **The list line says only what is OFF** (owner request, 2026-10-03: "only show
where there is something odd/off"): "In BETA and the stable source" (and "Published" with no stable
tree) is said by nothing, and the view count left the line for the Statistics view.
**Each level of choice has its own look** (2026-10-01): CRT's tabs; the four screens as tabs
under them, across the whole Maintainer tab above both panels (`Button.ScreenTab` - still
Buttons, so `OnModeClick` and the `Selected` class are unchanged - in CRT's tab colours: the
chosen one underlined in `Main_TabUnderline_Selected`, the others' LABELS greyed so a red badge
keeps its colour, no bolding since a wider label would move the tabs after it); and a
submission's Board data / Files / Contributor as ONE joined switch (`Button.Segment`, neighbours
1px-overlapped to share a border, the chosen one filled in `Mode_Selected_*`). The strip is a
WrapPanel so a narrow window never clips a tab: since the renames of 2026-10-04 the four need ~910
px with the real font, so at CRT's 800 minimum the last wraps onto a second line, inside the tab
(`In_a_narrow_window_the_screen_tabs_wrap_inside_the_tab`). `TabMaintainerScreenTabsTests` pins all of it, read off the
templates' ContentPresenters. The list column below stays 370 wide.

**A system says WHERE IT IS, and can be looked at in the stable source** (owner request, 2026-10-04: "where
does this sit now, as I do not think it is in BETA nor stable? Shouldn't there be somewhere a possibility
to see what we actually do have in BETA or stable"). Under its name the STAGE LINE (`SystemStagesDisplay`,
`SystemView.Stages.cs`; it replaced the revisions line): Submitted (the newest submission, CRT's state
words), BETA, Stable - each said, "Not there" included. Board data and Files carry a **Data: BETA | Stable**
switch (`SystemView.Stable.cs`; the rule `SystemSections.ShowsStable` - the pick, where the system is,
kept for the next system): the stable source has its OWN read-only table and file tree beside BETA's, so
switching never touches a change in BETA's table; read through `SystemDetailRequest.Tree`
(`DataTreeNames.Production`) by `SystemEditFlow.ReadStableTableAsync` and `SystemFilesFlow` with
`SystemFileSource.Production`, re-read when a promotion moves it (`QueueRefreshRules.StableBoardChanged`).
The stable table shares the picked pills (`TableEditorsForSharedChoices`, the Maintainer tab's three
tables).
**Every file tree says each file's SIZE** (owner request, 2026-10-04: "Ideally the 'Files' actually also
states its size (everywhere)"): `SystemFileEntry.SizeBytes`, the size of what opens from the row -
filled by the server's `TreeFileSizes` (through `SubmissionPathRules`) for a system's Files and a
submission's, and from the production plan's `FileSizes` for Beta > Prod (`SystemFileEntries.WithSizes`);
null shows nothing. Worded by `FileSizeWording`, which Unused files shares. **Account > "Unused files"
is the same tree** (same request: "the exact same tree-view ... including visualization of images and
opening of files"): `FileTreeView.ShowListing(..., openAll: true)` over `UnusedFilesDisplay.TreeEntries`,
opened from the list's `PublicDataUrl`.
**Both queues show the system's FILES AS A FOLDER TREE** (owner request, 2026-09-28: "a file-structure
for all existing files in BETA or PROD, and then a highlighting of files changed (added, removed,
changed) ... 'show only changed files' ... including their parent folder"). One control,
`FileTreeView` (the model and words are `Handlers/FileTree`): the system's own folder plus every
shared file it uses, each new / changed / removed in the table's own colours, "Show only changed
files" on by default. **Beta > Prod** draws production after the publish under the plan, from
CRT.Data's `SystemFileEntries.ForPromotion` over the plan's copies, removals and the new
`unchangedFiles`; **Contributor Submissions** shows it as the submission's **Files view**
(`TabMaintainer.Files.cs` - a window from a "Files..." button until 2026-09-30, which "can easily
be missed"), from `GET /api/review/submissions/{id}/files` (`SubmissionFileTreeFlow` over
`ForApproval` - read the first time the view is shown for a submission, since it walks the board's
folder and builds the publish plan). Its headline leaves the workbook and highlight file the
approval writes FROM THE TABLE out of its count and names them apart (`FileTreeWording.Summary`),
so it gives the Files button's number. **Pointing at a file shows
the TABLE'S hover card** (`FileTreeView.FilePreview.cs`: the table's `BoardTableFilePreview`, one side,
`showHeadline: false` - the owner asked for "only the relative path" and no change text, which the
row's pill already gives), and **double-clicking opens it**, for every maintainer, through the
host's `IFileTreeFiles` (`FileTreeFiles`, each file fetched once per tree shown): BETA's and
production's files from their PUBLIC data addresses, exactly as CRT downloads them (the plan and
the tree answer carry them; the bearer token is never sent there), a submission's own by hash
through the review API, opened from the temp folder through `OpenedFiles`, which also admits the
workbook, highlight and KiCad types. The workbook and highlight file are generated on approval, so
before then the card says it opens BETA's current copy (`FileTreeWording.Note` - the one thing it
says beside the path). **Only what a row SAYS reacts** - its icon, name and marker, the template's
`RowContent` panel - not the empty width of the list beside it (owner request: "It gets confusing when I
am in the middle of an empty screen"); the hover shading is moved onto that panel too, as a STYLE,
since a `Background` set on the panel in the template outranks it. Folders show Font Awesome's
OUTLINED plus/minus box (the Regular face, linked for this), close in both views, and the controls
sit on the LEFT: "Expand all", "Collapse all", then "Show only changed files". There is
no local-folder button any more (removed the day it came, owner request: "with this new folder view
you can scrap that"). The tree is never inside a ScrollViewer - a list measured with unlimited
height builds every one of a board's ~2,000 rows.

**A NEW system is PLACED in CRT's drop-down lists on the Systems screen before it can be approved**
(owner request, 2026-09-27). `SystemPlacementView` (inside `SystemView`) shows BETA's main Excel data
file in CRT's order with the new system as the one panel that moves - the SAME drag as the Drafts
tab's schematic images, `Controls/ListRowDrag` - plus its hardware name, board name and notes; Save
sends "after this row" (`GET`/`POST /api/review/systems/listing`, `SystemListingFlow`, the placement
kept in migration 0011's `systems.listing_*` columns). The approval refuses an unplaced new system
(`ApprovePublishFlow.ListingForPublishAsync`), the publish inserts the row (`MasterListing.Insert`,
CRT.Data - only the NEWEST master generation, one row, nothing else touched), promotion inserts it
into production's file after the nearest row production also lists, and pushing a never-promoted
system back out of BETA removes it. **Two systems may never share a hardware name + board name**
(`MasterListing.NamesTakenBy`): CRT keys a board by that pair, so it would be one board twice. The
tab words and checks all of it through `SystemPlacementDisplay`, which calls the SAME
CRT.Data name rule the server does. The Systems list puts the systems waiting for a place first,
marked, and the Systems badge turns to an attention badge counting those this account can place; the
review screen warns above the table. The panel is never rebuilt under unsaved changes by the minute
check.

**The submission view IS the table** - the Drafts tab's table editor, from
`src/CRT.App/Controls/BoardTable/`, in DOCUMENT mode (`BoardTableEditor.Open(BoardTableDocument)`: no draft file,
Save raises `SaveRequested`, the host saves). **Choosing a submission opens its table in the panel**
(`TabMaintainer.Table.cs`). **The queue is GROUPED BY BOARD** (`TabMaintainer.QueueItems.cs`,
words and grouping from `ReviewQueueDisplay`): a heading per board ("Commodore / C64 / 250407",
"New system" when nothing of it is published), and each submission as its comment and one grey
line. A submission NOT waiting for this account is dimmed and says it is "with the other
approver" - only the exception is marked. **Headings are DISABLED `ListBoxItem`s, so the list is
read by ITEM (`SelectedQueueRow` / `SelectQueueRow`), never by index** - a heading shifts every
position. **There is no Refresh button: the queue checks itself** (`TabMaintainer.QueueRefresh.cs`,
`QueueRefreshRules`, 2026-09-26) every minute while the tab is on screen with CRT's window in FRONT,
and on coming back to either - not in the background, since each check extends the session. A check updates the LIST only:
it never reloads the open submission's table, re-reads its detail only when its queue row changed,
and an open submission decided elsewhere stays on screen with its decisions off
(`ReviewDecisionWording.DecidedElsewhere`) instead of closing. A NEW SYSTEM's table is compared with
the SUBMISSION ITSELF as it was opened (owner decision, 2026-09-26): its rows start white and only
the maintainer's own inserts, edits and deletions are coloured; the change pills picked show
nothing until there is one. **So its baseline IS a board and `HasBaseline` is TRUE** - which is why a
changed cell's tooltip cannot work its own wording out and is TOLD it
(`BoardTableDocument.Create`'s `baselineLabel`): there it reads "As submitted: (empty)", the file
card's own word for that side, instead of "Published value: (empty)", which named a board that does
not exist (owner report, 2026-09-26). A draft names nothing and keeps the default.
**The decision buttons say which way the decision goes** (owner request, 2026-09-26): Approve is
CRT's green (`Button_Ok_*` - a missing `DynamicResource` draws an unstyled button in silence, so
`SharedTableColourKeysTests` checks every key the tab's markup names against CRT's App.axaml, where
the tab's ten keys of its own - `Badge_*`, `Mode_Selected_*`, `Queue_Divider` - now live too), "Request changes" and
"Reject" the red. Colour is not a licence to move it: Approve stays last and separated. (Compared with nothing it showed a lone "0 Flagged"; compared with an empty
board everything was green - tried and turned down.) A save makes the edits the submission, so the
reopened table is white again.
The table's picked pills are remembered between runs in CRT's own settings (see the list above).
The badges are the SERVER's answers, in the queue's `ReviewQueueEntry` (`ReviewQueueFlow`:
`PublishedBoardLocator.LocateSystem`, and `ApprovalStatus.CanApprove`); the opened row takes the
detail's answers, which judge shared files against the tree as it is now. **There is no change
summary any more** (owner request, 2026-09-26: "just
scrap that information") and no "View in table format" button.
**A submission has THREE VIEWS, chosen by buttons above it** (owner request, 2026-09-30;
`TabMaintainer.SubmissionViews.cs`, words in `SubmissionViews`): **Board data** (the table - first,
and what every newly chosen submission opens on), **Files** (the file tree, above) and
**Contributor**. Switching HIDES, never closes, so the table and its unsaved changes are untouched;
the same submission keeps its view through a minute check. The decision bar stays under all three.
**What the table cannot show -
highlight and calibration-point changes, the automatic checks' findings - is listed in a few lines
above it** (`ReviewNotInTable`); both are published by an approval, so do not drop those lines.
Each section is a counted heading and one change a line, set in under it (owner request, 2026-10-04,
in the owner's words: "Component highlights have [1] change:" / "Removed component [hest] from
schematic "Board layout"") - `ReviewNoteLine.Indent`.
**Who sent it and how their OTHER submissions went** is the Contributor view's (there was a one-line
"From dh@hinet.dk - [5] other submissions: ..." beside the view buttons until 2026-09-30, removed at
the owner's request: "this should be available under Contributor"): the server's `ContributorHistory`
builds CRT.Data's `ReviewContributorFacts` (the detail's `contributor`, pinned by
`ReviewWireContractTests`), `ReviewContributorLine` words the name and counts. Same contributor = same account or
same email, as `SubmissionReplacementRules`; a rejection counts only when a maintainer made it,
and replaced or never-finished submissions not at all. **The Contributor view is their whole
record** (owner request, 2026-09-30: "the trustworthiness history of the contributor"; server 3.7.0):
whether this submission came from an account (a verified address) or was typed in without one, and
every OTHER submission the counts count - exactly those (`ContributorHistory.Counted`), newest
first, at most 50 - with its system, description, state in the contributor's own words, and what
they were told. `ReviewContributorHistory` words it. **The counts read "[5] submissions in total
whereof [1] published to stable and [4] rejected"** (owner request, 2026-10-01): the total is still
the OTHER submissions, and "published" splits into stable and BETA only from the server's
`PublishedToStable` (3.9.0) - an older server's stays plain "published". **No text a user reads
calls a contributor "their", "them" or "they"** (owner request, 2026-10-01) - "the contributor",
or no pronoun at all; the two wording classes' tests assert it.

**A maintainer signed in on the tab is signed in for the whole of CRT** (owner request, 2026-10-01:
"I want to use that email address everywhere in the CRT app"). `TabMaintainer.SignedInChanged`
(every change of `thisSession` goes through `UseSession`) reaches `Main.ShareMaintainerSignIn`,
which hands the session to the Feedback tab (the account's address, read only) and the Drafts tab,
whose Submit dialog shows it read only and sends the session's token with the create request
(`SubmissionClient.CreateAsync`'s `bearerToken`) - so the submission is the ACCOUNT's and the
Contributor view says "Sent with an account". `ContactAddress` is the one rule; the account's
address is never saved as `UserSettings.ContactEmail`, so signing out brings the typed one back.
**Shared whatever the Configuration tab says, until signing out** (owner request, 2026-10-02): the
tab turned off or hidden for want of work changes nothing, and the remembered sign-in is restored
at launch with the tab off too (`TabMaintainer.RestoreSignInQuietly`, no request). The server's decision mails go to the account
for such a submission (`SubmissionNotifier.NotifyDecisionAsync(SubmissionRecord, ...)`, 3.9.0) -
they read the typed address alone before, which a signed-in submission does not carry.
**A maintainer edits their own account under Account > "My account"** (owner request, 2026-10-03;
`MyAccountView` - the "Your account" window, opened by "Account" beside "Sign out" under the lists,
until 2026-10-04; the server's `AccountSelfServiceFlows`,
`POST /api/accounts/me/name`, `/me/email`, `/me/email/confirm`, `/me/password`, migration 0016,
server 3.12.0; the session alone since 3.13.0). **The session is enough - NO current password anywhere** (owner decision, 2026-10-03:
"as I see it as you are already logged in"; an ACCEPTED RISK, recorded in NewContributeStrategy.md's
Threat 2 - do not put it back without asking). A new ADDRESS still needs a code mailed to it - it
changes only when the code comes back (the capitals alone: at once, same mailbox), typed in after "I
have received a code" (the owner's wording, as are "Name or handle" and "New email address"). A new
address another account holds answers EXACTLY like a free one, and its owner gets a "somebody tried"
mail instead of a code (owner decision - registration's anti-enumeration rule). An address or
password change signs out every OTHER session, spends reset codes already mailed, and mails the old
address (owner decision); both spend the per-address mail budget, which is also what keeps the
password change's Argon2 hash from being callable without limit. Every account the server returns
goes through `TabMaintainer.ApplyAccount` (the session, the remembered sign-in, `SignedInChanged`),
and the remembered name and address are read again from `GET /api/accounts/me` once a launch - only
on the paths that already talk to the server, never from `RestoreSignInQuietly`.
**A FILE REPLACED UNDER ITS OWN PATH COLOURS NOTHING** - without something saying so, a replaced
PDF or scan is approved unseen, the blind spot the retired change summary's file list used to
close. Lines above the table said it until 2026-09-30 ("Files: [12] included ([1] replaced under the
same name)", "KiCad data included: [3] files"); **the Files button's COUNT says it now**
(`SubmissionViews.ChangingFiles`, through `SystemFileEntries.ForApproval` - the tree's own rule, so
`SubmissionViewsTests` holds the button and the tree's headline to one number). Do not drop the
count: it is what those lines' tests became. The lines left do not say "(not in the table)" - that
read as if the COMPONENTS were missing - and each
count is "[2]" with the number bold, so a line comes as `ReviewNoteRun` pieces; a new system's
lines count ("Highlights included for [2] components") instead of naming every row. Its file card reads
both sides from the submission and says no "Unchanged" (`IBoardTableFileSource.SaysUnchanged`).
Each submission's table reopens on the sheet last looked at IN IT (`thisSheetBySubmission`,
passed as `Open`'s `preferredSheet`; one not opened yet starts on its first sheet with a change). Every way of leaving the table asks about unsaved changes (choosing
another submission, signing out, quitting CRT), a decision is refused until the table is saved, and the queue keeps its
selection across a refresh. **A test or render may show the tab in a plain window** (`new Window { Content = tab }`, as
`TabMaintainerQueueTests` does) - it restores a session only from `ReviewSessionStore`, which only
CRT's start-up points at the real file. **Never call `ReviewSessionStore.Initialise()` from a test**,
and a test pointing it at a temp file (`InitialiseAt`) joins the `"ReviewSessionStore"` collection and
points it at nothing again afterwards. It sends the table as an AMENDMENT,
which the server decides (`AmendSubmissionFlow`): authority over the board, still undecided, still
at the version the maintainer opened, only the table's nine sheets
(`SubmissionRowsBoard.WithTableSections`), files rebuilt from the rows (kept, or taken from the
published tree, else refused), the same validation as a new submission, and approvals already given
cleared. It runs under the `PublishLock`, and the store re-checks the version and the state inside
its own transaction - the flow's checks are only the early answer. Migration 0009 keeps what each amendment replaced - the contributor's original first. The
contributor is told (`SubmissionStatus.AmendedByMaintainer`, the BETA mail). **Anything the table
covers is `Controls/BoardTable`'s and `CRT.Data`'s, not either tab's** - the same rule as "the control only
paints" above, across the Drafts tab and the Maintainer tab.

### Data layer (`Handlers/Data/`, split across `src/CRT.Data/` and `src/CRT.App/Handlers/Data/`)

**Phase 1 of [Assets/NewContributeStrategy.md](../Assets/NewContributeStrategy.md) split this
folder across two projects.** `src/CRT.Data/` (a plain, Avalonia-free `net10.0` library, no
namespace change — everything stayed in `Handlers.DataHandling`) holds the board-data schema and
read/write logic shared with the server and the Maintainer tab; everything that orchestrates the
app itself (data-root resolution, sync, settings, logging) stayed in `src/CRT.App/Handlers/Data/`.
**Before moving anything else here, check which side it actually belongs on** — `DataValidator` was
in the original move list and was pulled back out because it calls `DataManager.HardwareBoards`/
`LoadBoardDataAsync`/`DataRoot` directly, which would have meant either moving `DataManager` (a
1600+ line app-orchestration class, clearly out of scope) or rewriting `DataValidator`'s signature
(a public API change the strategy doc's own coverage table already flags as a owner decision,
not something to do in passing). A class with a hidden dependency like this is not a "move."

**In `src/CRT.Data/`:**

- `BoardData` / `BoardDataReader` / `BoardDataWriter` — the schema (in
  [src/CRT.Data/BoardData.cs](../src/CRT.Data/BoardData.cs)) and read/write logic (via EPPlus) for
  per-board Excel files: schematics, components, component images/highlights, local files, links,
  credits, and KiCad signal mappings.
- `BoardWorkbookSchema` — the sheets and columns, `BuildRows` (BoardData -> cells) and, since
  2026-09-24, its inverse: the `Map*` row mappers, `MapRows`, `EntriesOf` and `WithRows`. They moved
  out of `BoardDataReader` so the table editor saves through the same mapping the reader loads
  through; `BoardWorkbookSchemaTests` round-trips every field of every sheet by reflection.
  **`AllSheets` is the sheet order of every workbook the app or server WRITES and of the table's
  tabs, and it ends in "Credits"** - as all published workbooks do ("Important signals" before it;
  it was the other way round until 2026-09-24, reported). Readers find sheets by name, so the order
  never affected reading.
- `BoardWorkbookWriter` / `BoardWorkbookStyle` - how EVERY workbook CRT writes looks: a new or saved
  draft, and every board the server publishes to BETA (`PublishExecutor`). The look is read off the
  shipped references (C64 250407, C128 310378), and was misread once: **the title band ("Components")
  is BLACK with white text on every sheet** - in the styles part `theme="1"` is the DARK colour - and
  the header row is 57.6 pt only on Board schematics, 28.8 on Components and Component images, one
  ordinary line elsewhere (owner report, 2026-09-30). Check a change against the references' XML, not
  against memory; `BoardWorkbookWriterTests.The_shipped_reference_has_a_dark_title_band...` reads the
  C128 file itself (finding the band by its text). **Since 2026-10-02 every sheet has the same
  preamble** (owner request): hardware, board and REVISION DATE (on every sheet - the reader still
  takes Board schematics' one), a blank line, the title band, the headers on row 6 - with the panes
  FROZEN under them. The "Documentation of columns ... availble here:" line and its link to the old
  Wiki are gone: the project owner removed them from every published board by hand, while the copies
  in `Assets/Data` still carry them.
- `BoardTableDocument` / `BoardTableSheet` / `BoardTableRow` / `DraftTableSession` /
  `BoardTableClipboard` — the Drafts tab's table editor model (see the Drafts paragraph under Tabs).
  Avalonia-free so the Maintainer tab's table uses the same rules.
- `BoardTableSearch` / `WorklogSearchQuery` — the table's search box (which rows a search shows,
  which runs of a cell it found), over the Workbooks tab's grammar, which moved here for it.
- `BoardTableHistory` — the table editor's undo/redo: one sheet snapshot per step, back to the
  last save. In `CRT.Data` so the Maintainer tab's table gets it too.
- `BoardTableRowDrag` — one row being dragged in the table: `MoveOnto` the row under the pointer,
  `Step` past an edge, `Finish`. Ghosts are never targets; one undo step per drag.
- `ComponentPlacement` — where a NEW component's rows go (its category, natural label order) and
  that an edited one stays put. Every writer that adds components goes through it.
- `ComponentListBuilder` — the main window's component list: region filter, category filter, search,
  and the `ComponentListItem` rows it produces. Also owns `IsSupportedKiCadRawFile`.
- `ContactLinkFormatter` — classifies a contributor's contact string and builds its href.
- `TextLinkFinder` — which runs inside a user-typed free-text field are web links, so the UI can
  render them as clickable. Deliberately conservative: only `http://`, `https://` and `www.` at a
  word boundary count. A bare `example.com` does NOT, because that is the shape repair prose
  collides with — `74LS08.pin3`, `5.0V`, `notes.txt` — and a false link is worse than none. Pure
  string work, unit tested; the Avalonia half (`Handlers/Theme/TextLinkRenderer.cs`) stayed in
  `CRT.App`, since rendering a clickable run needs a control.
- `ICrtLog` / `CrtLog` — the logging seam these classes call instead of the app's own `Logger`
  (which is a static singleton tied to Velopack's AppData folder and has no place in a library
  the server will also reference). `CrtLog.Sink` defaults to null (writes nothing, which is exactly
  what every test in `CRT.Data.Tests`/`CRT.App.Tests` relies on — no test may call
  `Logger.Initialize()`, and by construction none needs to for this reason either); `CRT.App`
  installs `AppLoggerAdapter` as the sink from `App.OnFrameworkInitializationCompleted`,
  immediately after `Logger.Initialize()`, so a moved class logs to the exact same file it always
  did.

**Stayed in `src/CRT.App/Handlers/Data/`, deliberately:**

- `DataManager` (static) — resolves the data root (default location or `--data-root=`), loads the master
  Excel workbook, and drives sync against the online checksum manifest. App orchestration, not schema.
- `DataValidator` — validates board/contribution data. See the note above for why it stayed.
- `PublishedDraftRetirer` — the sequencing around deleting a draft whose work is published TO
  PRODUCTION (since 2026-09-27; "merged" into BETA no longer counts, because "Push back to queue" can
  take it out of BETA again - owner request), with the folders it leaves empty removed too. The same
  status check drives `Main.SourceSwitchNotice.cs`: a user downloading from the BETA source is told,
  once, when their work reaches production, to switch back to the normal source. The
  check (`DraftRetirement`, in `CRT.Data`) runs off the UI thread, each folder is re-stamped just
  before it is deleted, a draft with unsaved table edits is kept, and the board cache is cleared for
  every folder touched. `Main.RetirePublishedDraftsAsync` runs it after EVERY launch status check
  (`SubmissionStatusRefresh`'s `onFinished`, not `onChanged` - retirement depends on a state, and a
  receipt already stored as merged never changes again) and when "My submissions" closes.
  **A receipt names its system by ID ("Commodore/C64/250407"); everything in the drafts flow is
  keyed by the board's WORKBOOK path.** `DraftStatusReader.ResolveForSystem` maps one to the other
  (through `SystemDescriptorRules.SystemIdFromExcelDataFile`, the function that built the id), and
  `RetirableDraft.ExcelDataFile` carries the workbook key onward. Calling `Resolve` with the id
  looks one folder too high and finds nothing - which is why no draft was ever retired until
  2026-09-25. **A NEW system's draft is retired too** (2026-09-25), once its published board is on
  the machine - which needs the master to list it, so the contributor can then open it instead.
  Its draft and published workbooks have different names (the published one carries the
  generation), so the draft is found by its folder and keeps the marker's key; Main passes only
  the boards the master lists, and refreshes the draft-only entries after retiring one.
- `UserSettings` — JSON-persisted user preferences (theme, window placement, MiniPro path override, etc.).
- `KiCadProjectData` / `KiCadProjectLoader` / `KiCadRawProjectLoader` — parse raw KiCad PCB/schematic
  files into a normalized bundle so the Schematics tab can highlight matching copper and wire geometry.
- `Logger` — writes to the app's log file in the AppData folder alongside settings.
- `ComponentImageQueries` — which component images and entries the popup shows for a region, and
  which of them carry an oscilloscope baseline.
- `OverviewHtmlBuilder` / `OverviewModels` — bill-of-materials grouping and the printable HTML for
  the Overview tab, plus the `OverviewRow`/`OverviewLink` models it renders. **The Overview AXAML
  binds these through an `xmlns:data="clr-namespace:Handlers.DataHandling"` mapping** — if you move
  or rename them, update `TabOverview.axaml` too.

**Three Avalonia-based geometry helpers were also left behind on purpose** despite an earlier plan
naming them for this move: `Handlers/Geometry/HighlightRectBuilder.cs`, `PolygonGeometry.cs` and
`RectGeometry.cs` are all built on `Avalonia.Point`/`Rect`/`Matrix` and are called from 21 files
across the Schematics/Worklog rendering pipeline. `CRT.Data` must never carry an Avalonia
dependency — a server has no business depending on a desktop UI framework — so converting these to
a portable representation is a real refactor of its own (touching all 21 call sites), not a file
move, and belongs in whichever later phase actually needs a portable geometry library.

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
  conditional `ItemGroup` in [CRT.App.csproj](../src/CRT.App/CRT.App.csproj).
- `Security/ExternalTargetLauncher` — the only sanctioned way to open an external link or local file from
  the UI; it restricts targets to HTTP/HTTPS/mailto URIs or local paths that resolve inside the current
  data root *and* carry a document/image/data file extension from its allowlist (it hands files to the
  OS shell, which would run a `.exe`/`.bat`/`.lnk` instead of displaying it), rejecting anything else.
  Use this rather than shelling out directly when opening user/data-supplied links or files.

**QuestPDF** (the workbook PDF export) is the app's one non-Avalonia UI dependency worth knowing
about. Three things about it:

- **Its licence is a condition on this project, not just a package reference.** The Community
  licence QuestPDF is used under is free for individuals and for organisations under $1M USD annual
  revenue — which this project is. An organisation above that threshold shipping a fork would need
  its own commercial licence from QuestPDF.
- **`QuestPDF.Settings.License` must be set before it generates anything**, or the first export
  throws. It is set by `WorkbookPdfExporter.ConfigureQuestPdf()`, called once in
  `App.OnFrameworkInitializationCompleted`, not at the export call site, so a missing call fails at
  launch in development rather than in a user's hands. The tests call the same method rather than
  setting the licence themselves, so they export under the app's settings.
- **QuestPDF 2026.9.0 flipped its font defaults, and `ConfigureQuestPdf` deliberately flips them
  back.** The new defaults use no system fonts and THROW on any glyph or font family they cannot
  find, so one emoji or CJK character typed into one repair note failed the whole export
  (reproduced by `WorkbookPdfExporterTests`, which fails without it). So `UseSystemFonts` is true,
  `ThrowOnMissingTextGlyphs` and `ThrowOnMissingFontFamilies` are false, and
  `EnableDetailedLayoutErrors` is false (its message quotes document content into the log). **Check
  QuestPDF's release notes on every bump** - this one changed behaviour without breaking the build.
  The same release made registering a font under a custom name obsolete, so the icon font is now
  referenced by the family stored in the .otf (`WorkbookPdfExporter.IconFontFamily`,
  "Font Awesome 7 Free") at weight 900 (`.Black()`), which keeps a system-installed Font Awesome
  REGULAR face from being picked instead.

### Contribution service (`src/CRT.Server/`) — the PHP's replacement, not yet its retirement

An ASP.NET Core minimal-API service on the project owner's own AlmaLinux box, built by phases 3-5 of
[Assets/NewContributeStrategy.md](../Assets/NewContributeStrategy.md). **Read that document before
touching this project** — it carries the phase status, the decisions already settled, and the traps
per phase. It is the handoff document between sessions; nothing else records that state.

**It runs alongside the PHP rather than having replaced it.** Retiring the PHP is Phase 7 and has
not started, so both exist and the legacy path below is still live.

**The shape to keep: endpoints are a RIM, flows hold the decisions.** `*Endpoints.cs` reads a
request into plain values, calls one method on a `*Flows` class, and maps the verdict onto a status
code — nothing else. Every rule lives in a pure class taking its store and mailer as arguments, so
it tests against fakes with no database and no network. A rule that migrates into a handler body is
a rule no test can reach, the same defect this file describes for logic trapped in a `UserControl`.

Things that are load-bearing and easy to undo by accident:

- **The status codes are the security design, not presentation.** Registration and password reset
  always answer 202 whether or not the address is known, and login answers 401 with no reason for
  every failure. A 409 for a taken address is an account-enumeration oracle.
- **The rate limiter runs BEFORE the Argon2 hasher**, which allocates ~128 MiB per verification.
  Checking the password first makes the limiter decoration and leaves a memory-exhaustion vector.
- **A request body's size limit is written on its ROUTE** (`.WithBodyLimit(...)` where it is mapped),
  and everything else gets 64 KB. A new route that posts rows or a file list needs one, or it
  answers 413 to the first real board; `RequestBodyLimitsTests` builds the real route table and
  fails on any body-reading route missing from its list.
- **Only token HASHES are stored**, for sessions, verification links, reset links and the
  submission capability token alike.
- **Every mail is HTML with a plain-text copy** (owner request, 2026-10-03). A template writes its
  mail ONCE as `MailBody` blocks (paragraph, quote, code, steps; runs styled plain, bold, italic,
  `Named` - a system as `[<b>id</b>]` - or link), and `MailBody` renders both bodies, HTML-encoding
  every run - so **no template writes markup**, and nothing a person typed becomes markup in somebody
  else's mail. `SmtpEmailSender` sends multipart/alternative (the plain part is what spam filters
  expect). The contribution mails carry the project owner's own wording; each greets by the
  account's name (`MailRecipient`, "Hi there," without one) and ends with the contact line
  (`EmailTemplates.ContactName`/`ContactAddress`), except the production notice to the
  administrators. The BETA check box is named through CRT.Data's `ConfigurationWording`, which the
  Configuration tab's markup also reads (x:Static), so the mail cannot name a box CRT does not have.
- **Feedback from CRT's Feedback tab is the service's since 2026-10-03** (owner request: "retire the
  old Feedback backend PHP"): `POST /api/feedback` (`FeedbackEndpoints` -> `FeedbackFormReader` ->
  `FeedbackFlow`), the PHP page's form and its plain-text `Success` answer UNCHANGED - CRT.Data's
  `FeedbackContract` is the one copy, and Apache forwards `/app-feedback/` here for installed CRTs, so
  changing a field name, the zip's file name or the answer breaks every older CRT. CRT's own text
  files (`FeedbackContract.InlineFiles`, held to AppConfig's names by `FeedbackAttachmentNamesTests`)
  are SHOWN in the mail in a `MailBody.Preformatted` block ("logfile should still show as monospace");
  everything else is saved under `FeedbackRoot/feedback-<random>`, the PHP's folder, which the
  project owner opens - and DELETES - from a network share, so every folder and file of it is made
  group-writable (`FeedbackFlow.ShareWithGroup`; the folder is the service's own, group `crt-server`
  - the group `useradd --system` made beside the user, owner decision: no group of its own - and setgid). At most
  250 MB packed (owner decision), which CRT streams from a temporary file and the server writes
  straight to disk - neither end may hold it in memory. The mail is FROM the service with Reply-To the sender
  (the PHP sent it from the sender's address, which fails SPF), and a mail postfix refused is a 502
  the user sees - `IEmailSender.SendAsync` answers whether it went for exactly this - and its saved
  files are deleted again, since no mail names them. The saved folders together are capped at
  `FeedbackMaxStoredBytes` (20 GiB; code review, 2026-10-04): anonymous senders must not fill the
  disk down to the reserve the blob store shares.
- **CRT's launch check-in is the service's since 2026-10-03** (owner request: retire the
  `app-checkin` PHP page): `POST /api/usage/check-in` (`CheckInEndpoints` -> `CheckInFormReader` ->
  `CheckInFlow`), the PHP page's URL-encoded form and plain-text answer UNCHANGED - CRT.Data's
  `CheckInContract` is the one copy, the version is the User-Agent, and Apache forwards
  `/app-checkin/` here for installed CRTs, so a renamed field stops every older CRT counting. It
  writes ONE row into `crt_update` exactly as the PHP page did (`MySqlCheckInStore`: `NOW()`,
  `versionMajor` 2, "" rather than NULL) - **including the sender's address**, unlike board views,
  because every Fun facts chart counts `DISTINCT ipaddr`; never from a local network. `apiJson` holds
  only the country (the PHP stored ip-api.com's whole answer, which nothing reads).
- **`SubmissionPathRules` is the single path-containment rule** and it RESOLVES then checks
  containment rather than pattern-matching for `..`. Every write path and the maintainer's file-read
  path go through it. Do not add a second way to turn a submitted string into a path.
- **`ApprovePublishFlow` is one of the two irreversible operations in the system** — no revision
  history is retained, so a publish overwrites in place. Its header explains why the order of its
  checks is the design. The other is deleting a system, below.
- **The administrator can DELETE A SYSTEM completely (owner request, 2026-10-03; Account > "Delete a
  system", `SystemDeletionFlow` over `SystemDeletionFiles`/`SystemDeletionRules`).** Its folder in
  BOTH trees, its row in both NEWEST main Excel data files, both manifests rebuilt, then the `systems`
  row - the schema's ON DELETE CASCADE takes every submission, approval, maintainer and invitation -
  and its board views (owner decision: "Delete them"). Shared files stay (AutomaticRemovalScope).
  Open submissions go too and their contributors are MAILED with the reason, which is then required
  (owner decision: "Delete them and mail the contributors"). **The plan is held to a fingerprint**
  (both trees' files, both rows, the record, every submission's state): anything changed since the
  confirmation is a 409. **Refused** for a system an OLDER main Excel data file lists (frozen - and a
  master naming a missing workbook makes DataTreeUsage incomplete for good) or with a file another
  board uses. **The order is the crash story**: folders probed, rows (stable first), files, manifests,
  and the record LAST - a file left keeps the record, so pressing Delete again finishes. Do not move
  the record's deletion earlier.
- **The administrator can RESET THE CONTRIBUTION DATA for going live (owner request, 2026-10-04:
  "When I go-live with this, it should not have old data visible ... the real sources of BETA and
  stable must not be touched, but all contributor and maintainer data should go away"; Account >
  "Reset contribution data", `DataResetFlow` over `IDataResetStore`; server 4.7.0).** ONE
  transaction (`MySqlDataResetStore`, DELETE never TRUNCATE - ids must not start again) empties
  `submissions` (and everything cascading from it), `production_approvals`, `maintainer_invitations`,
  `maintainers`, `systems`, every account with `is_administrator = 0` (sessions and tokens cascade),
  `audit`, `crt_board_views` and `crt_api_calls`, and writes the reset as the new history's first
  row; then every deleted submission's partial uploads and every stored blob nothing needs go
  (`SubmissionFlows.CollectUnreferencedBlobsAsync`). **Never touched**: the data trees, their
  manifests and main Excel data files, `crt_update`, `crt_board_view_batches`, the feedback folders,
  the administrators. Owner decisions: accounts and maintainer lists go ("I need real data anyway"),
  board views go, and NOTHING refuses it - not even work waiting under BETA > Stable (the owner
  copies stable over BETA and rebuilds the manifests around it). **The system records go too** -
  every flow makes one on demand (`EnsureSystemAsync`) and reads a missing one as "never touched by
  the pipeline", which is what every system is after a reset. **Only while `AllowDataReset` is true**
  in appsettings (default false; DEPLOYMENT.md "Going live") - a stolen administrator session cannot
  wipe the database. The counts shown are held to a fingerprint (submissions and accounts by count
  AND highest id, pools, invitations, approvals, records - NOT the history, board views or usage,
  which grow by themselves): 409 when they moved. Under the publish lock. Nobody is mailed.
- **The Systems screen is for EVERY maintainer, every system (owner decision, 2026-09-27).**
  `GET /api/review/systems` and `POST /api/review/systems/detail` (`SystemEndpoints` over
  `SystemOverviewFlow`) answer to anybody `CanReviewAnything` - contributors' addresses included,
  for systems they do not maintain. That widening was asked and decided ("Everything for
  everyone"); do not narrow it to `CanReview(system)` without asking. `GET /api/review/production`'s
  optional `awaitsYou` (`ProductionPromotionRules.AwaitsAccount`) is false once this account has
  approved that BETA state - the BETA badge's count.
- **A maintainer's edit of a SYSTEM goes STRAIGHT TO BETA - through the ordinary approval (owner
  decision, 2026-10-03: "it should go directly to the next queue, 'BETA > Stable', so it can directly
  be tested in BETA"; it reversed the same day's "Becomes a submission").** `POST
  /api/review/systems/table` reads BETA's board for any maintainer (with a `fingerprint` and
  `mayEdit`), `POST /api/review/systems/files` lists the files it uses, `POST
  /api/review/systems/edit/check` (`SystemEditFlow.CheckAsync`) answers what a publish would remove,
  making nothing, and `POST /api/review/systems/edit` (`PublishAsync`) turns the table edit into a
  submission from the maintainer's ACCOUNT through `SubmissionFlows.CreateAsync` and `FinaliseAsync`
  and then APPROVES it as that maintainer with `ApprovePublishFlow.ApproveAsync` - so every approval
  rule (removals shown, shared files, one submission in BETA, the publish lock) still stands between it
  and BETA. **Do not give it a path that writes BETA by itself**; the approval is what keeps those
  rules. Only the system's own maintainers and the administrator may send one (`CanReview(systemId)`),
  not to a system closed to contributions, **not while the system waits under BETA > Stable** (owner
  decision, same day: "it should simply disallow it, even if this is coming from a maintainer" -
  `OneSubmissionInBeta.NoChangeMessage`, only where production publishing is configured), and NOT
  while the account's last change of the system still waits in the queue (it would replace that one -
  `SubmissionReplacementRules`). Its order is the design: BETA must still hold the board the table was
  read from (the fingerprint, 409 otherwise - publishing a table read before another publish would
  undo that publish silently, as rows are replaced whole), only the nine sheets come from the request
  (`WithTableSections`; highlights and calibrations are BETA's, read from its sidecar), every cited
  file must already be in BETA and is imported from there, an edit that changes nothing is refused, a
  reason (the description) is required, and `expectedRemovals` must be the list the check answered
  (409). Published: BETA's manifest is rewritten (`ReviewEndpoints.RewriteBetaManifest`) and nobody is
  mailed. A publish the approval refuses after all leaves the submission PENDING in the queue, its
  approvers mailed - nothing is lost.
- **The production panel names WHOSE WORK a promotion carries (owner request, 2026-09-27).** The
  right-side panel was the file copy list alone, which cannot answer the question the button asks.
  It now reads: the merged submissions this carries (contributor, their own description, when it was
  accepted), then what needs a second look (approval, removals, shared files), then the files
  COLLAPSED behind a counting header. The facts come from `ProductionPromotionRules.Carrying` over
  **the same query and window `AfterPublishAsync` uses to mail those contributors**, so the screen
  and the mail cannot disagree about who is affected; it never fails the plan. It is deliberately
  NOT a diff of the data - the submissions were reviewed in the table and the tick is the check.
- **One submission in BETA per system (owner decision, 2026-09-27).** `ApprovePublishFlow` step 3b
  refuses any approval of a system that `ProductionPromotionRules.IsAwaitingProduction` - a push-back
  can only be per system (below), so two contributors' work in BETA at once could only be taken back
  together. Worded once, in CRT.Data's `OneSubmissionInBeta`, which the Maintainer tab's `ApprovalGate`
  shows beside a greyed-out Approve. **On only when production publishing is configured**
  (`ServerOptions.IsProductionPublishingConfigured`, passed to the flow's constructor): without it
  nothing leaves BETA and every system would close after its first approval. **A board copied to
  production BY HAND is not waiting** (code review, 2026-09-29): it was never PROMOTED, so its record
  said "waiting" for ever and blocked the system. When the record says waiting, the approval and the
  Beta > Prod list ask the trees (`ProductionPromotionFlow.RecordIfProductionAlreadyHoldsAsync`):
  nothing to copy, nothing to remove and no listing row to add means production already holds it -
  recorded as in production and audited (`production.found_in_place`, in the system's history). The
  list caches a "differs" answer per BETA/production state for ten minutes, and reads its two facts
  (`AwaitsYou`, `CarriesDiscardedDraft`) for the whole list in one query each
  (`ProductionPromotionFlow.ListEntriesAsync`) - it is polled every minute by every open tab.
- **A system's HISTORY** (`SystemHistoryRules`, 2026-09-27) rides on `POST /api/review/systems/detail`:
  its submissions' sending and decisions plus the audit rows naming the system or one of its
  submissions ("#id"), newest first. **An event only appears if it is audited under the system's id**
  - a new kind of system event needs its `AuditEntry` subject to be the system id, and its action
  added to `SystemHistoryRules.ShownActions` and worded in the maintainer's `SystemsDisplay.HistoryLine`.
- **Board views (owner request, 2026-09-27): which boards CRT users look at, and in which country.**
  CRT counts a VIEW each time a PUBLISHED board has been on screen for 10 s - every time, not once
  a day (`BoardViewTracker`; a reload of the board already on screen is not a new one) - and sends
  batches (`BoardViewReporter`, an outbox file beside the settings) to `POST /api/usage/board-views`
  (`BoardViewFlows`). A report is sent again UNCHANGED, same `BatchId`, until an answer finishes it
  (`BoardViewContract.DeliveryFor`: only 2xx and a refusal of the report itself - 400/413/415/422 -
  finish it; any other 4xx is a server set up wrong and keeps it), and the server stores a batch
  once - that is what makes a retry after a lost answer safe. While CRT is open, anything still
  waiting is tried again every 15 minutes (`AppConfig.BoardViewRetryInterval`). One row per view in `crt_board_views` (migration 0013; its column
  names follow `crt_update`, as agreed with the project owner): names from the PUBLISHED main Excel
  data file (`BoardViewNameDirectory` - an unlisted id is ignored, which keeps made-up text off the
  public Fun facts page), version/OS/CPU as the check-in sends them, `fromBeta`, and the country from
  ip-api.com. **No address and no identifier is stored.** A batch from a local-network address is
  the project owner's own machines: stored "for now" (owner request, 2026-09-27, to check the
  numbers) with the SERVER's own country and marked `fromLocalNetwork`, until
  `ServerOptions.CountLocalNetworkBoardViews` (default true) is set false - then it stores nothing,
  the check-in's rule. **Mandatory, no opt-out**
  (owner decision); the Wiki's `Information-collected` page says what is sent. The Systems screen
  shows the counts (`BoardViewStatisticsRules`: whole UTC days, BETA-source views apart; a count that
  cannot be read leaves the numbers out, never the screen). The Fun facts page reads the table
  directly (`Assets/Webserver/funfacts/funfacts_board_views.php`; `helligsoe` has a table-level
  SELECT). The launch check-in writes `crt_update` (the service since 2026-10-03, below); it and the
  other two usage tables were moved into `crt_review` by hand the same day and are no migration's.
- **"Push back to queue" rolls a BETA board back to production's state (owner decision,
  2026-09-27).** `BetaRollbackFlow` over CRT.Data's `BetaRollbackPlan`, the promotion's mirror:
  production's bytes go back over BETA through `VerifiedFileCopy` (never a bare copy onto a served
  path), BETA-only files are deleted, the BETA checksum manifest is rebuilt at once, and ONE store
  transaction (`RecordRollbackAsync`) returns every submission merged since the last promotion to
  `pending` with the maintainer's comment, CLEARS their approvals, and makes the system's recorded
  BETA state follow the tree. **It works because a MERGED submission keeps its blobs** -
  `SubmissionCollectionStates.Live` contains `Merged` - which is the fact an earlier session got
  wrong when it called this impossible.
  **A rollback is PER SYSTEM and cannot be otherwise**: `PublishMerge` replaces rows wholesale, so
  one contributor's work cannot be picked out. The confirmation therefore NAMES every submission it
  takes back; do not remove that list or the "ALL N submissions" sentence, which is what stops a
  maintainer discarding other people's accepted work unknowingly. A system never promoted has its
  folder REMOVED from BETA instead (there is nothing to restore from). **"Never promoted" is the
  system's RECORD, never an empty production folder** (code review, 2026-09-27): a promoted system
  whose production folder cannot be read plans as `BetaRollbackKind.ProductionUnreadable` and is
  refused with 409 - it used to be wiped out of BETA.
  **Paths are ORDINAL** - the server's trees are case-sensitive, so two spellings are two files.
  **A SHARED file is touched only when BETA still holds exactly the bytes a returning submission
  carried** (its `submission_files` hash): then it goes back to production's version; one production
  never had STAYS (2026-09-27, `AutomaticRemovalScope` - it used to be removed through
  `UnusedFileRemover` unless another board cited it). One written again since belongs to somebody else and is left alone. Leaving the submission's
  shared change in BETA (the first version) leaked it to production with the next promotion of any
  board citing it. A failed record answers `NotRecordedMessage` rather than a 500; pushing back
  again finds the tree level, moves nothing, and records it.
  **CRT hears it as `returned`** - not a submission state, but RECORDED: the rollback writes the
  moment it returned each row into `submission_beta_returns` (migration 0015, code review
  2026-09-29), and `ContributorFacingState` reports `returned` for a `pending` row whose return is
  not older than its latest decision - so any later decision ends it. It used to be INFERRED from
  "pending and carrying a decision comment", which any other path leaving a comment on a pending
  row would have turned into a rollback that never happened. CRT shows "Taken back out of
  BETA - waiting for review again", coloured NeedsAction, and a merged receipt is now re-asked
  weekly past its 30-day window rather than never, so a late rollback still arrives.
  Submissions merged BEFORE the last promotion are already in production and cannot be pushed back
  at all; a corrective submission is the only route there. **Known gap, not handled:** a file in
  THIS board's folder that the rolled-back submission added and another board has since cited is
  deleted with the rest of the board's BETA-only files.
  **"Reject" on Beta > Prod is the SAME rollback** (owner request, 2026-09-28: "Then there is no
  need to push it back and then reject it"): the request's `reject` makes `RecordRollbackAsync` set
  the submissions to `rejected` instead of `pending`, the contributor gets the queue's own rejection
  mail (`SubmissionNotifier.NotifyTakenOutOfBetaAsync`), and the audit row is `beta.rejected`. Do
  not give it a code path of its own - BETA must move exactly as for a push-back. The answer's
  `rejected` lets the Maintainer tab say so when an older server pushed back instead.
- **The workbook's "# Hardware:" / "# Board:" CAPTION is the SERVER's too (owner report,
  2026-09-28)** - every sheet of a workbook approved into BETA had lost its first two lines. The
  rows never carry it: `PublishMerge.Build` keeps the replaced board's, and a board with none (a new
  system, or one published before the fix) takes the names it is listed under
  (`PublishMerge.CaptionedAs`, all or nothing). `BoardDataCaptionTests` reads the source and fails on
  any `new BoardData { ... }` that does not name both - the caption had been dropped by a hand-listed
  copy three times.
- **The REVISION DATE is the SERVER's, at both stages (owner decision, 2026-09-26).** "Server always
  wins, and what is typed by user is not important": `ApprovePublishFlow.BuildPlan` stamps
  `BoardWorkbookStyle.FormatRevisionDate(nowUtc)` into the plan and `PublishExecutor` writes THAT
  into the workbook rather than computing its own, so the board and `systems.current_revision` carry
  one value. They used to disagree - the workbook was stamped, the plan's descriptor was built from
  the SUBMITTED date - and that row is the base a contributor's next draft is diffed against.
  It is also what lets a NEW SYSTEM publish at all: its seeded workbook carries no revision date
  (nothing to inherit one from, and CRT never asks), so the old rule refused `publish.no-revision`
  at the approval. **Do not make the client stamp one** - `PublishExecutor`'s header explains that a
  locally stamped draft immediately looks newer than the board it came from, so every draft would
  report drift.
- **Two roles, and authority is PER SYSTEM (owner decision, 2026-09-25).** An
  ADMINISTRATOR (`accounts.is_administrator`, granted by hand - see DEPLOYMENT.md) reviews and
  publishes everything and assigns maintainers; a MAINTAINER is an account in one or more systems'
  pools (the `maintainers` table, the Systems screen in the Maintainer tab) and reviews AND publishes exactly
  those systems. There is no recommend-only role and no `is_reviewer` flag any more. The question
  is always "administrator, or in THIS system's pool" - `ReviewAuthority`, given a `ReviewAccess`
  (the account plus its pool, loaded once per request in `ReviewEndpoints.AuthenticateAsync`).
  Every route naming a submission checks it against that submission; the queue is filtered by it.
  A submission that REPLACES an existing SHARED file with different bytes
  (`SubmissionRecord.TouchesSharedFiles`, decided by `SubmissionSharedFiles`) needs TWO approvals -
  a maintainer of the board AND the administrator - for the BETA publish and again for production
  (owner decision, 2026-09-25; narrowed to REPLACEMENTS 2026-09-27 - adding a new shared file needs
  one approval, since no other board cites it yet, and production's flag is
  `ProductionPromotionPlan`'s shared copies with `Change == Replaced`). **The approval decides it
  from the tree as it is NOW, not from the stored flag** (`ApprovePublishFlow.TouchesSharedFilesNow`),
  and stores its answer both ways (`SetTouchesSharedFilesAsync`) - the flag set under the old
  add-or-change rule would otherwise have held every older queued submission for a second approval.
  **The create-time flag is only a floor**: the approval and the detail screen re-check against the
  tree as it is NOW (`ApprovePublishFlow.TouchesSharedFilesNow`), because a shared file cited
  unchanged at create can be changed by another publish before this one is approved.
  **`ApprovalRules` (CRT.Data) is that rule, and the server sends its `ApprovalStatus` to the review
  app as the same record**, so the Approve button's wording (`ApprovalWording`) cannot promise a
  publish the server would not perform. The first approval publishes nothing (state `approved`,
  migration 0008's approval tables); a production approval is tied to the BETA content hash it
  was given for. Whoever must approve is mailed (`SubmissionRouting`). A pool row counts as the
  board's maintainer only when its account can approve as one
  (`ReviewAuthority.CanGiveMaintainerApproval`: verified, not locked, not an administrator) - a row
  left for an account later made administrator by hand made a shared-file change wait for ever.
  **TOTP is deliberately NOT built** - the accepted risk is recorded in NewContributeStrategy.md's
  security review.
- **WHICH files a submission may carry is `SubmissionFileRules` (CRT.Data), run at create AND again
  in `PublishPlan`** (security review, 2026-09-25): allowed types only, no dot-segments, every file
  cited by a row, and nothing outside the board's own folder and the two shared folders unless it is
  byte-identical to the published copy. **`SubmissionRulesShippedDataTests` runs every such rule over
  every board in `Assets/Data`** - it is what found the cross-board citations and the misnamed images
  a stricter rule would have refused. Any new submission rule must pass it; do not loosen it to pass.
- **A blob is re-verified before a publish writes anything**, and every copy is hashed and renamed
  into place only on a match (`VerifiedFileCopy`, which `BlobStore.TryCopyToAsync` and the
  production promotion both use). Never copy a blob into the tree directly.
- **A served file is REPLACED, never opened for writing, and every folder is checked FIRST**
  (owner report, 2026-09-28). The project owner copies data into the trees by hand as root, and such a
  file refuses an open but not a replace-by-rename, which needs only its folder: an approval
  wrote the new workbook and then answered 500 at the root-owned highlight file beside it. So
  board files go through CRT.Data's `FileReplacer` (the sidecar and workbook writers), copies
  through `VerifiedFileCopy`, manifests through their own temporary - **a new writer into a
  served tree must do the same**. A FOLDER copied in as root refuses everything and the service
  cannot fix it, so the publish, the promotion and the push-back ask `TreeWriteAccess` about every
  folder they will touch - after the link check, since a probe is a real file - and refuse with
  the folders named and the fix command in the log. Running the service as root was raised as
  the alternative and advised against: it is the internet-facing process, and it would end step
  3's interlock. **The highlight file is also written only when its CONTENT changes** (owner
  request, 2026-09-28; `BoardSidecarWriter.WriteIfChanged`, compared through CRT's own readers):
  the server's layout differs in bytes from a hand-made one (LF, text order, an empty calibration
  root), so every board's first publish "replaced" its `.json` and "Publish to production" listed
  it with nothing in it changed. A file with a root nothing reads is still rewritten.
- **Publishing is TWO stages (owner decision, 2026-09-25): Approve writes BETA, and
  "Publish to production" in the Maintainer tab copies a SYSTEM from BETA to Production.** Only bytes
  already in BETA; per system, not per submission; a shared-file change needs the maintainer AND the
  administrator. `ProductionPromotionPlan` (CRT.Data, pure, shown to the maintainer and then
  performed) decides the copies; `ProductionPromoter` copies; `ProductionPromotionFlow` sequences.
  The publish request carries the BETA content hash the maintainer was shown and is refused if BETA
  moved - that is what "only after he has checked it in BETA" means in code. Both stages take the
  one `PublishLock`. **Off until the three `Production*` settings are set** (DEPLOYMENT.md step
  13), which also reverses step 3's kernel-level interlock for the production data folder. A
  contributor's "merged" reads as "Published to BETA source" and becomes "published" ("Published
  to source") once its system is
  promoted (`ProductionPromotionRules.ContributorFacingState`); CRT keeps asking about "merged"
  receipts at every launch until then for 30 days after the decision
  (`SubmissionReceiptPresenter.MergedRecheckWindow`), and weekly after that
  (`MergedLateRecheckInterval`) - a BETA rollback can turn one "returned" at any time. A new system's row in the master
  workbook is added by the publish and the promotion since 2026-09-27, from the place a maintainer
  gave it on the Systems screen (see "Maintainer" under Tabs).
- **Automatic removal stays INSIDE THE SYSTEM'S OWN FOLDER (owner decision, 2026-09-27):**
  `AutomaticRemovalScope` (CRT.Data) filters what a publish, a promotion and a push-back may remove
  to `<Manufacturer>/<Hardware>/<Board>/...`. A shared file, or another system's, that a board stops
  citing stays - an orphan for Account > Unused files, which is not limited by this. It is the
  condition that made dropping the second approval for ADDED shared files safe.
- **No orphan files (owner decision, 2026-09-25): `DataTreeUsage` (CRT.Data) is the one rule
  for what a data tree uses**, and a publish or promotion REMOVES what the board stops citing that
  nothing else uses (inside its own folder - see above). **Always from a list the approver was shown:** the submission detail and the
  production plan carry `removals` (CRT.Data's `FileRemovalPreview`), the approval sends it back,
  and the server refuses (409) if the list it would now remove differs. `UnusedFileRemover` does
  every removal, re-reading the tree first and failing closed. Files that were ALREADY orphans are
  the administrator's "Unused files" entry, under Account (`UnusedFileFlows`). `DataTreeUsageShippedDataTests`
  fails on any file in `Assets/Data` that nothing uses. **A folder CRT reads by NAME** (the MiniPro
  IC tests) must be in `DataTreeUsage.FoldersReadByName`, or the server calls its files unused -
  `DataFoldersReadByNameTests` guards it. The PREVIEWS read workbooks through `WorkbookReadCache`
  (once per file version, 2026-09-25); the removal never does. Details: NewContributeStrategy.md,
  "Orphan files". **A board workbook is read at its spelling ON DISK** (2026-10-04, owner request to
  prove the Unused files list): a master naming it in another case was opened by the master's
  spelling, not found on the Linux server, and read as citing nothing - so everything only it cites
  was listed unused while Windows CRTs used it. Also counted since: the legacy master's "KiCad data
  file" column, and a cited path as the OS resolves it (`./`, `//`, `..`, trailing dots -
  `DataTreeUsage.Resolved`). Any change to the rule may only make it keep MORE.
- **A contributor discarding their own draft is shown to maintainers** (owner request,
  2026-09-28). CRT's Discard reports every submission of that board still open (CRT.Data's
  `DraftDiscardContract.WhichToReport`) to `POST /api/submissions/{id}/draft-discarded` with the
  submission's token - recorded on the receipt first, so an offline discard is sent at the next
  launch (`DraftDiscardReporter`). The server keeps the first notice (`submission_draft_discards`,
  migration 0014; `DraftDiscardFlow`) and audits it for the system's history; the review queue, the
  opened submission, "Beta > Prod" (list and plan) and the Systems screen all show it
  (`DraftDiscardWording`). It never withdraws the submission - the wording says so.
- **A newer submission from the same contributor REPLACES the older one** (owner decision,
  2026-09-26), when it is queued: the same contributor's older submissions of the same system are
  `withdrawn` with a comment - only if still `pending` and never amended (a maintainer's edits or a
  first approval keep it), checked in the withdrawing transaction (`SubmissionReplacementRules`,
  `ISubmissionStore.WithdrawReplacedAsync`). Same contributor = same account, or the same contact
  email - UNVERIFIED, an accepted risk. `withdrawn` has no other producer, so CRT reads it as
  "Replaced by a newer submission" (neutral, not refused); a contributor-initiated withdraw would
  need a state of its own.
- **A board's KiCad data travels in its submission** (owner decision, 2026-09-26).
  `SubmissionKiCadFiles` is the ONE rule: only the types CRT reads, only inside the board's own
  "KiCad data" folder, exempt from "every file is cited by a row" at create AND in PublishPlan -
  and an AMENDMENT must carry them over, since its file list is rebuilt from rows, which cite no
  KiCad file. The client sends the union of the draft's and the synced official folder
  (draft wins); the maintainer sees them in the submission's Files view and counted on its button
  (`SubmissionViews.ChangingFiles` - a "KiCad data included" line above the table until
  2026-09-30). A new file type in a submission needs a signature in
  `SubmissionContentRules` (its default arm fails closed).
- **"Already held" is the blob store OR the published tree at the same path** (2026-09-25). A file
  published unchanged is IMPORTED into the store at create (`BlobStore.TryImportAsync`, hashed as
  it copies) instead of being asked for. Without it, the first submission to every board
  re-uploaded the whole board. Do not change it to "read such files from the tree later": review,
  finalise and publish all read the store, and a shared file in the tree can change while a
  submission waits for review.
- **The Maintainer tab's `submittedFiles` field is CRT.Data's `SubmittedFileFact`**, serialised by the
  server and read by the Maintainer tab as the same type, so the two cannot drift. The full list of
  what the review closed and what it left open is in NewContributeStrategy.md's security model.

### Contribution webserver (`Assets/Webserver/`) — the LEGACY path

A working copy of the PHP deployed at `classic-repair-toolbox.dk/app-contribution/` — the server
side of CRT 2.x's Contribute tab, split into two entities: `api/index.php` receives the uploads
CRT 2.x posts there (CRT 3.0.0 posts nothing to it - its constant, `AppConfig.ContributionUploadUrl`,
and the parser for its "too old" answer were removed on 2026-10-04), and `review/index.php` + `review/functions.php` are the
review queue (admin-only, IP-restricted via `review/.htaccess`) that diffs each submission against
the live server data and merges or rejects it. `api/index.php` requires `review/functions.php` for
its shared helpers. It is deployed by hand and never ships with the app (Assets are whitelisted
per-file in the csproj).

**This whole path is being retired and is NOT kept in step any more** (owner decision,
2026-09-23) - see "One change, every side of it" above. The new pipeline replaces it, the old
method will not be supported alongside it, and a change to shared behaviour does not have to be
mirrored here. What follows describes how it works and stays accurate for as long as it is
deployed; it is no longer an instruction to update it.

**The payload is a two-sided contract.** `ComponentContributionPayload` in CRT 2.x's
`Tabs/Contribute/ComponentContribution.axaml.cs` (read it at the `2.5.0` tag - it was removed from
CRT on 2026-10-04, nothing having built it since 2026-09-25) is what `review/functions.php` parses,
and the PHP deliberately mirrors the app's Excel-reading
rules (case-insensitive headers, exact board-file resolution, `.json`-beside-the-Excel highlights,
the `# Revision date:` marker). When you change either side, change the other in the same
sitting, and run the PHP suite too: `php Assets/Webserver/tests/run-tests.php`
(~1s, no webserver needed — PHP is installed on this machine; the tests sit outside
`app-contribution/` so that folder stays an exact mirror of what is deployed). **When the payload contract
changes, also bump `$minimumContributionVersion` in `api/index.php`** to the first released
version carrying the new contract — older apps are rejected at upload with an update-required
message. Full details, file inventory (including which files are legacy) and known caveats:
[Assets/Webserver/app-contribution/README.md](../Assets/Webserver/app-contribution/README.md).

## Installed CRTs keep working

**Every CRT ever released stays installed, and must keep working** (owner request, 2026-10-04: "It
is important that all older versions will continue to work, including the checkin and feedback").
The server is deployed once and is always the newest party, so the SERVER bends: there is one API,
no `/api/v1` beside `/api/v2`, and it only ever grows. Where a CRT truly cannot be served any more,
it is TOLD so in words - never left to fail in a way nobody can read. The full plan is
NewContributeStrategy.md's "Installed CRTs keep working".

**It protects what a RELEASED CRT uses - never a pre-release** (owner decision, 2026-10-04: "in
this development time it is of course okay that the API changes for the 3.0.0 version, since this
has not been released yet - no need to do new API end-points because I e.g. bump from
"3.0.0-alpha.2" to "3.0.0-beta.7""). The project owner moves CRT to the next version (2.5.0 to 3.0.0)
and works through `-alpha.N` / `-beta.N` builds until the release. Anything only those pre-releases
use is changed IN PLACE - renamed, removed, given a new meaning - with CRT's side changed in the same
session: no new route beside the old one, no field kept for an older alpha or beta. **So the first
question about any API change is "does a RELEASED CRT use this?"**: is it in a `crt-*.txt`, or is it
one of the routes CRT 2.x calls (the check-in, feedback, the legacy contribution address - released
before the frozen files existed, held by their own contract tests)? No: change it freely. Yes: the
rules below. A tester left on an older pre-release that such a change breaks is told to update by
the API REVISION (below) - nobody has to know which build has which shape. A `crt-<version>.txt`
written for a version not yet PUBLISHED (bumped to `3.0.0`, still fixing before the release
workflow runs) may be deleted and written again by the next test run; it is untouchable from the
moment that release is published.

**The rules for every change to a route, a JSON body, an answer or a form a released CRT uses:**

- **Add, never rename or remove.** A new optional request field, a new field in an answer, a new
  route - and the server treats a missing new field as the old behaviour. A field a CRT needs that
  the server can no longer fill stays, with its old meaning, beside the new one.
- **A new meaning is a new route.** 4.0.0's change to what `systems/edit` does was only allowed
  because nothing was released; after CRT 3.0.0 it would be `systems/publish-edit` beside it.
- **A new REQUIRED request field, a narrower rule, a new member of an enum an answer carries** - each
  breaks installed CRTs too (an older CRT reading an unknown enum name fails its parse), even though
  nothing is removed. Treat it as a break.
- **The server ships before the CRT release that needs it.** A CRT that calls a route the deployed
  server lacks gets a 404 it cannot explain.
- **Readers ignore what they do not know.** `ReviewApiContract.WireSettings` must keep skipping
  unknown fields (`ClientVersionContractTests` pins it) - never switch on
  `UnmappedMemberHandling.Disallow`.
- **The data trees are an API too.** Older CRTs read the published workbooks, the JSON sidecar, the
  main Excel data file and `dataChecksums.json`, which the server now writes: add sheets and
  columns, never rename or drop what an older reader looks for. Nothing machine-checks this one.

**The machinery, all built 2026-10-04:**

- **Every request names its CRT**: `User-Agent: CRT <version>` (`OnlineServices.VersionForServer`,
  now on `SubmissionClient` too - submissions sent none before). CRT.Data's `CrtVersion` reads and
  orders it (SemVer: `3.0.0-alpha.2` < `3.0.0`).
- **`ClientVersionPolicy`** (CRT.Server, `Handlers/Compat/`) can refuse a CRT older than a minimum
  for `/api/submissions` or the Maintainer tab's routes (`/api/review`, `/api/admin`,
  `/api/accounts`) with **HTTP 426 and CRT.Data's `ClientOutdatedAnswer`** ("please update CRT to
  version [x] or newer"). **No minimum is set.** Raising one is the last resort for a break that
  cannot be avoided: a MAJOR server bump whose VERSION.md row names the versions it turns away. The
  check-in, feedback, board views and health check are **never** refused, whatever the minimums -
  `ClientVersionPolicyTests` lists them and fails on a new top-level route until it is classified.
  A request naming no CRT version (a browser on a mailed link) is always let through.
- **The API REVISION tells a CRT it is too old - nobody picks a version** (owner question,
  2026-10-04: "how do I know which version of app uses which version of server? ... I do not have an
  overview on what kind of specific changes are done to API and if those changes are compatible or
  not"). A CRT version does not move when the API does, so a minimum version needs somebody to know
  which build has which shape. CRT.Data's `ClientVersionContract.ApiRevision` IS the shape: CRT's
  review and submission clients send it as `X-CRT-Api-Revision` (`OnlineServices.NameApiRevision`),
  and `ClientVersionPolicy` answers a LOWER revision on the gated routes with the same 426 ("please
  update CRT to the newest version", `minimumVersion` null). A newer revision, or none, is served.
  **`ApiCompatibilityTests` raises it for you**: `api-revision-<N>.txt` holds what a CRT built for
  revision N may use - written the first time N is current (that run fails once: commit the file),
  GROWN as the API grows - and any change that breaks it fails until `ApiRevision` goes up by one.
  **Raise it by HAND** for what the surface cannot see: a route given a new meaning with the same
  shape, or a request field the server starts to require. Each raise is a MAJOR server bump whose
  VERSION.md row names the revision. `GET /api/health` reports the server's revision, and the
  Maintainer tab's Account > "Server version" shows it beside this CRT's ("API server version" and
  "API application version" - `ServerVersionDisplay`). After 3.0.0's release it should not move: a break of a released
  CRT is the `crt-*.txt` check's, below.
- **CRT shows the server's words.** `ApiRefusal.Read` takes the sentence out of any refusal body
  (`message`, else `error`); `SubmissionClient.RefusalMessage` and `ReviewApiClient.WithServersWords`
  use it on every path, so a refusal a newer server invents still reads as a sentence, and a 426
  becomes `ReviewApiFailure.ClientOutdated`. A 401 keeps "Sign in to continue.".
- **`ApiCompatibilityTests` is the check** (CRT.Server.Tests, `ApiCompatibility/`). `ApiSurface`
  turns the API into lines - every route, every JSON request body per route, every answer record's
  fields with their JSON kinds, every enum member as it travels, and the wire names of the forms
  and headers. **When CRT.App.csproj's `InformationalVersion` is a release (no `-`), that surface
  must be frozen as `crt-<version>.txt`**: the first test run without the file writes it and fails
  once, saying so - commit it with the release. Every frozen file is then compared with the surface
  on every run; a line a released CRT relied on that has gone fails, named. **Never edit a
  `crt-*.txt`**, and never "fix" a failure there by changing the expectation: put the field or route
  back. A break the project owner has decided on goes in `allowed-breaks.txt` with its reason.
- **Every answer with data in it is a named CRT.Data record**, because the surface only sees
  records. The sign-in answer, a submission's detail and four smaller ones were anonymous objects
  until 2026-10-04 (same JSON now). An endpoint may write an anonymous object only to carry
  `message`, `error` or `errors`; `Every_answer_with_data_in_it_is_a_named_record` fails otherwise.
  The surface's answer records are CRT.Data's `*Answer`/`*Response` types plus a short list in
  `ApiSurface.OtherAnswers` - a new answer type with another name goes in that list.
- **CRT 2.x's contribution upload is answered, not 404'd**: `POST /api/legacy/contribution`
  (`LegacyContributionEndpoints`) gives the PHP page's own 426 `OUTDATED_VERSION 3.0.0 - ...` text,
  which 2.5.0 and later turn into "please update" (tested against a verbatim copy of 2.5.0's parser).
  Apache forwards `/app-contribution/api/` to it once the PHP page is retired.
- **The server COUNTS which CRT version calls which route** (owner request, 2026-10-04: "how about
  tracking the API end-points, to see if it is possible to retire any, if almost no versions uses it
  any more"; server 4.7.0). Every request reaching a route is counted in memory by Program's pipeline
  - after routing, BEFORE the version gate and the rate limiter, since a CRT turned away still
  called the route - per UTC day, route PATTERN (never a path, so no id) and the version its
  User-Agent names (`ApiUsageRules`, `ApiUsageCounter`), and written into `crt_api_calls`
  (migration 0018) every five minutes and at shutdown (`ApiUsageFlusher`). No address, no account.
  Account > "API usage" (`ApiUsageFlow`, `GET /api/admin/api-usage?days=`) lists EVERY mapped route
  (`ApiRouteList`, the live route table - a route nobody called is the answer most worth seeing)
  with its versions, least-called first in each area, and the launches per version from the
  check-ins.

**Retiring a route a released CRT uses** (owner decision, 2026-10-04 - this loosens the rule above
from "forever" to "until shown to be unused"; nothing can be retired before 3.0.0 is released):

1. **Account > "API usage" is the evidence**: which versions still call it in the last 90 or 365
   days, and how many installations still launch those versions.
2. **The project owner decides, per route.** Never automatically, and **never a route in
   `ClientVersionPolicy`'s Forever area** - the check-in, feedback, board views, the health check and
   the 2.x contribution address, which every CRT ever released sends (`NeverRetired`;
   `ApiUsageRulesTests` pins the set).
3. **The route stays MAPPED and answers HTTP 426** with CRT.Data's `ClientOutdatedAnswer` naming
   the first version that no longer needs it (`ClientVersionContract.Outdated`) - never a 404, which
   an installed CRT cannot explain. Every CRT from 3.0.0 shows that as "please update CRT".
   (Nothing has been retired yet, so the small helper that maps a retired route to its 426 is built
   with the first one.)
4. **What its body and answer carried goes in `allowed-breaks.txt`**, each line with the reason
   ("retired 2027-..., API usage showed no call from ... in 365 days") - the frozen `crt-*.txt`
   files are never edited.
5. **A MAJOR server bump**, its VERSION.md row naming the route and the versions turned away.

## CRT.Server's version is YOURS to bump

**The project owner bumps CRT by hand. They do NOT bump CRT.Server** - they
asked (2026-09-26) for it to be handled for them, because it only has to be visibly versioned so a
deployment can be identified. So when you change the server, you decide the bump and make it.
The full policy, the judgement calls and the history table are in
[src/CRT.Server/VERSION.md](../src/CRT.Server/VERSION.md); read it before bumping.

The short form: the version is SemVer 2.0.0 judged on **the HTTP API a client already installed
depends on**, not on the size of the diff. MAJOR when a client that worked can now fail (a route or
JSON field renamed or removed, a field made required, a rule that starts refusing what it accepted).
MINOR for new behaviour that breaks nobody (a new route, a new optional field, a field ADDED to a
response). PATCH for a defect fix or anything no caller can observe (a comment, a test, a pure
refactor, a reworded mail body). **A change only to `CRT.Data` still bumps the server** when the
server's behaviour moves with it - `CRT.Data` has no version of its own and ships inside the same
publish output. The version reached a plain `1.0.0` on 2026-09-26, so the rules above apply
literally - the next breaking change is `2.0.0`, and the old "bump the alpha counter instead" escape
hatch is gone.

Two edits per bump: `InformationalVersion` in
[CRT.Server.csproj](../src/CRT.Server/CRT.Server.csproj), and a row in that VERSION.md history table
saying what a caller would notice. `GET /api/health` then reports it
(`{"status":"ok","version":"...","utc":"..."}`), with the `+<commit>` metadata the SDK appends
stripped - that endpoint is unauthenticated and public, so it must never publish the source commit.

**This one is machine-checked.** The `Stop` hook
[hooks/server-version-bump.sh](hooks/server-version-bump.sh) fires when a file the deployed service
is built from changed (`src/CRT.Server/` or `src/CRT.Data/`, excluding docs and tests) while
`InformationalVersion` stood still. It WARNS rather than blocking, because which bump a change
deserves is a judgement call - and "nothing a caller can observe changed" is a legitimate answer,
so ignoring it is sometimes right. It goes quiet the moment the literal moves.

## Release process

Versioning lives in [CRT.App.csproj](../src/CRT.App/CRT.App.csproj)
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

**Windows signing is done INSIDE `vpk pack`** (`VPK_SIGN_TEMPLATE` running jsign against the
YubiKey), in both release workflows, followed by a step that fails the release unless every
`.exe`/`.dll` in the full `.nupkg`, and `Setup.exe`, carries a valid signature. Do not go back to
separate jsign steps around the pack: that signed only the main exe and `Setup.exe`, and a step
after the pack cannot reach the `Squirrel.exe` (installed `Update.exe`) and launcher stub Velopack
adds inside the package. The step's comment in `build-and-release.yml` records the details that
were verified by dry run, including why the template must stay sequential (a YubiKey locks after
three wrong PINs). The macOS and Linux builds are not signed: macOS needs an Apple Developer ID
certificate (the YubiKey's Authenticode certificate cannot sign for it), and Linux has no
operating-system check of an AppImage's signature to satisfy.

**There is no separate maintainer application to release any more.** CRT Maintainer was merged
into CRT as its Maintainer tab on 2026-09-29, so it ships with CRT's version and CRT's workflow. Its
own workflow (`build-and-release-maintainer.yml`), the release repository
`HovKlan-DH/Classic-Repair-Toolbox-Maintainer` it published to, the `REVIEW_RELEASES_TOKEN` secret and
`MaintainerReleaseSeparationTests` (which kept its releases off CRT's update feed) were all retired
with it.
