# New contribution and maintainer strategy

A staged plan to replace the Contribute tab and the `app-contribution` PHP backend with a
local-first authoring experience in CRT, a shared C# data library, an ASP.NET Core service on the
project owner's own AlmaLinux server, and a maintainer surface - a separate desktop application until
2026-09-29, a Maintainer tab inside CRT since - with per-system maintainers.

**Status: Phases 0-4 are DONE, except that nothing PUBLISHES yet - a submission is queued for
review and stops there, which is by design until Phase 5 builds the maintainer application
(2026-09-21).**
The CRT.Server service is deployed and running on the project owner's AlmaLinux box: it answers over
HTTPS, creates its own schema, and accounts can register, verify their address, log in, refresh and
log out. The submission endpoints exist and are tested but have not yet been deployed or exercised
against the live service.

**Phase 5 is largely built** (the maintainer application, sliding-expiry sign-in, the maintainer
round trip) and **Phase 6a is DONE (2026-09-23)**: drafts are now stored as real board
folders rather than as row deltas, which reverses a Phase 2 decision - read that phase
before touching anything under `Drafts/`. **Phase 6's ROLES are DONE (2026-09-25)** as two
roles rather than the four planned - read the phase before assuming the four-role table - and
its two-stage publish, two-person approval for shared files, orphan removal and the maintainer's
table are DONE too (2026-09-25). Phase 7 is not started. **Phase 8 is DONE (2026-09-29)**: the
separate CRT Maintainer application is now CRT's Maintainer tab - read it before touching anything
maintainer-side.

This file is a handoff document. It is written to be picked up by an agent (or a person) who has
not been part of the conversation that produced it, across many sessions, with no memory of
earlier ones. Read [How to use this document](#how-to-use-this-document) first.

---

## Table of contents

- [How to use this document](#how-to-use-this-document)
- [Why this is being done](#why-this-is-being-done)
- [The target architecture](#the-target-architecture)
- [Decisions already made](#decisions-already-made-do-not-relitigate)
- [Access the project owner must provide](#access-the-project-owner-must-provide)
- [Phase 0 - Repository restructure](#phase-0---repository-restructure-done-2026-09-20)
- [Phase 1 - Extract CRT.Data](#phase-1---extract-crtdata-done-2026-09-20)
- [Phase 2 - Local-first drafts in CRT](#phase-2---local-first-drafts-in-crt-done-2026-09-20)
- [Phase 3 - Server service and accounts](#phase-3---server-service-and-accounts-done-2026-09-21)
- [Phase 4 - Submission pipeline](#phase-4---submission-pipeline-done-2026-09-21)
- [Phase 5 - Maintainer application, single user](#phase-5---maintainer-application-single-user-started-2026-09-21)
- [Phase 6a - Drafts as real board folders](#phase-6a---drafts-as-real-board-folders-done-2026-09-23)
- [Phase 6 - Maintainers](#phase-6---maintainers)
- [Phase 7 - Retire the PHP contribution path](#phase-7---retire-the-php-contribution-path)
- [Phase 8 - Maintainer tab inside CRT](#phase-8---maintainer-tab-inside-crt-done-2026-09-29)
- [Security model](#security-model)
- [Cross-cutting concerns](#cross-cutting-concerns)
- [Open questions for the project owner](#open-questions-for-the-project-owner)

---

## How to use this document

**One phase per working session, at most.** Several phases are themselves multi-session. Each phase
below has: a goal, a definition of done, the exact tasks, and the traps specific to it. Do not start
a phase before its predecessors are complete - the ordering is not arbitrary, and each phase is
written assuming the previous one's output exists.

**Work in the order given.** Phase 0 and Phase 1 in particular must come first, and are worth doing
even if the rest of the plan is abandoned.

**Record progress in this file.** When a phase is finished, change its heading from
`## Phase N - Title` to `## Phase N - Title [DONE yyyy-mm-dd]` and add a short note under it saying
what was actually built and anything the next phase should know. This file is the memory between
sessions; nothing else carries it.

**The existing project rules still apply in full.** In particular, from
[.claude/CLAUDE.md](../.claude/CLAUDE.md):

- **Never touch [CHANGELOG.md](../CHANGELOG.md).** Not as a finishing touch, not to record a phase.
  If an entry is warranted, say so in the session summary and let the project owner write it.
- **Tests arrive with the code, never after.** Run `dotnet test` before reporting any change done.
  A `Stop` hook enforces this.
- **Never commit or push.** Leave changes in the working tree.
- **Plain ASCII punctuation.** No em dashes, no smart quotes.
- **Write for hobbyists.** No prose implying paid work, customers or billing.
- **Wiki pages ship with the code that changes their behaviour**, and are pasted by hand by the
  project owner - never claim a Wiki page has been updated online.

---

## Why this is being done

The current design works well for one person and roughly a dozen boards. Four properties of it do
not survive growth to hundreds of systems and several maintainers. Each one drives a decision later
in this document, so they are worth stating precisely.

**1. A contributor cannot use their own work.** An edit is packed into a zip, posted, and then
invisible to its author until it is merged. Nothing is stored locally, because storing it in the
board Excel would be destroyed by the next sync. This is worst for the highest-value case: someone
documenting a whole new board must do the entire job blind, through a form, never seeing it render.

**2. A new system is not a contribution at all.** It arrives as an email or a GitHub issue and is
imported by hand. The single most valuable kind of contribution is the one the software does not
support.

**3. Review does not distribute.** `review/index.php` is a one-user page locked to a single IP. It
has no concept of who submitted something, who may judge it, or what was decided before. It cannot
be opened to a maintainer without opening everything to that maintainer.

**4. The board-reading rules exist twice.** Once in C# (`BoardDataReader`), once in PHP
(`review/function_xlsx-read.php`). [The webserver README](Webserver/app-contribution/README.md)
carries a parity table whose entire purpose is to remind a human to change both. That is affordable
now. At scale it becomes the thing that breaks, and it breaks silently - as a merge that quietly
corrupts a board.

Problem 4 is the most dangerous and, as it happens, the cheapest to fix. See
[Phase 1](#phase-1---extract-crtdata-done-2026-09-20).

---

## The target architecture

```
                              CRT desktop app
                  (contributors, users - and, in its Maintainer tab,
                   maintainers; a separate app until 2026-09-29)
                                     |
                 authors into Drafts/, submits manifest+blobs;
                 Maintainer tab: approve / reject / request changes
                                     |  HTTPS
                                     v
                      +-------------------------------+
                      |   CRT.Server (ASP.NET Core)   |
                      |   AlmaLinux, behind existing  |
                      |   web server on /api/         |
                      +-------------------------------+
                          |                        |
            review state  |                        |  published data
                          v                        v
                    +-----------+        +-----------------------+
                    |  MariaDB  |        |  app-data/  (prod)    |
                    | accounts  |        |  app-data-BETA/       |
                    | queue     |        |  plain files, as now  |
                    | audit     |        +-----------------------+
                    +-----------+

         Both C# programs compile against ONE shared library: CRT.Data
```

The key property: the app (its Maintainer tab included) and the server both read and write board
data with the same code. The PHP parity problem stops existing.

---

## Decisions already made (do not relitigate)

These were settled with the project owner. An agent picking this up should treat them as given, and
raise a concern only if implementation reveals one to be genuinely unworkable.

**TWO decisions on this list HAVE been reversed by the project owner.** The second, on 2026-09-29:
**the review tool is no longer a separate application** - it is the Maintainer tab inside CRT
([Phase 8](#phase-8---maintainer-tab-inside-crt-done-2026-09-29)). A maintainer is a CRT user, and two
applications, two versions and two release repositories for a handful of people was cost with no
benefit. The table's "Review tool" row below records the original choice.

**The first, on 2026-09-23: the draft file format.** Phase 2 stored a draft as row deltas in `draft.json`; a draft is now a real
board folder with its own workbook. The reasoning, what it cost and what it bought are in
[Phase 6a](#phase-6a---drafts-as-real-board-folders-done-2026-09-23). Every other decision
below stands.

| Decision | Choice | Rationale |
| --- | --- | --- |
| Contributor sees own edits before approval | **Yes, local overlay layer** | The central UX win; makes authoring a new system possible at all |
| Submission scope | **Whole system, transported as a delta** | Contributor never classifies their own work; hashes keep it small |
| Maintainer authority | **Approval publishes directly** | Any design where the owner presses every button recreates the bottleneck |
| Who owns a system | **A pool of maintainers, no single owner** | Work continues when one person goes quiet |
| Administrator scope | **Implicitly a maintainer of every system** | No escalation path to steal or mis-implement |
| New systems | **Always approved by the administrator** | Highest-risk item in the pipeline; no maintainer exists yet |
| Global-scope approvers | **Rejected** | Highest privilege for the lowest-frequency task; the recommend-only Reviewer (since retired) covered the need with no publish rights |
| Guiding principle | **Secure by design, not merely by default** | Insecure states must be unreachable, not just discouraged |
| Server stack | **ASP.NET Core on existing AlmaLinux box** | Enables the shared library; no new machine needed |
| Review state storage | **MariaDB** (already installed) | Queue/permission/audit data; hardware data stays as files |
| Published data storage | **Plain files, exactly as today** | Backup and comparison by copying folders |
| Review tool | **Separate Avalonia desktop app** - REVERSED 2026-09-29: a tab inside CRT (Phase 8) | Review is visual; reuses CRT rendering; audiences barely overlap |
| Repository layout | **One repo (monorepo)** | `CRT.Data` as ProjectReference, not a versioned NuGet package |
| GitHub as backend | **Rejected** | Owner considers it too troublesome; contributors must never need it |
| Maintainer server access | **Never** | Maintainers get a maintainer account, used from CRT's Maintainer tab, nothing more |
| Per-row UUIDs | **Retire them** | Base-revision diffing plus natural keys replaces them; see Phase 4 |

---

## Access the project owner must provide

The agent doing Phases 3 onward needs access that does not exist yet. **Phases 0-2 need none of
this** - they are entirely local - so this can be arranged while those are underway.

Nothing here should be committed to the repository. See
[Secrets handling](#d-secrets-handling-applies-from-phase-3) below.

### A. MariaDB

Needed from Phase 3. Create a dedicated database and user - do not reuse an existing application's
credentials.

```sql
CREATE DATABASE crt_review CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci;
CREATE USER 'crt_review'@'localhost' IDENTIFIED BY '<strong-password>';
GRANT ALL PRIVILEGES ON crt_review.* TO 'crt_review'@'localhost';
FLUSH PRIVILEGES;
```

Provide to the agent:

- host, port, database name, username, password;
- whether the agent may reach it **remotely** (see below) or only through a shell session on the
  server.

**Strong recommendation: do not expose MariaDB to the internet.** Keep it bound to `localhost` and
let the agent reach it over an SSH tunnel:

```
ssh -L 3307:localhost:3306 <user>@<server>
```

The agent then connects to `127.0.0.1:3307`. This needs no firewall change and no `bind-address`
edit, and it cannot be brute-forced from outside.

### B. The two data trees

The server already serves two independent data sets, and the app already switches between them via
`UserSettings.DownloadDataFromTestSource` ([Main/App.axaml.cs:93-94](../Main/App.axaml.cs#L93-L94)):

| Environment | Manifest URL | Role |
| --- | --- | --- |
| **Production** | `https://classic-repair-toolbox.dk/app-data/dataChecksums.json` | What ordinary users sync |
| **BETA** | `https://classic-repair-toolbox.dk/app-data-BETA/dataChecksums.json` | Validation before release to production |

Provide to the agent:

- the **absolute filesystem path** of each tree on the server (e.g. `/var/www/.../app-data/` and
  `/var/www/.../app-data-BETA/`);
- the user and group that owns them, and which user the service will run as;
- confirmation of how data is promoted from BETA to Production today (manual copy? rsync? a
  script?). **This matters:** the new service must fit that process rather than fight it, and the
  answer is not in the repository.

**All development targets BETA.** Production must not be written by any new code until the
project owner explicitly says so - see [Phase 3](#phase-3---server-service-and-accounts-done-2026-09-21).

### C. Server shell access

Needed to install the runtime, create the service and configure the reverse proxy:

- SSH access as a user with `sudo`;
- the web server in use (Apache or nginx) and where its site config lives;
- the domain or path the API should answer on (recommended: `https://classic-repair-toolbox.dk/api/`).

### D. Secrets handling (applies from Phase 3)

- **Never** commit credentials, connection strings, tunnel details or server paths to the repo.
- On the server, configuration lives in `appsettings.Production.json` in the service's own
  directory, owned `root:<the unit's `Group=`>`, mode `0640`, outside the web root. Root-owned
  rather than service-owned on purpose: the service reads its configuration and cannot rewrite it.
  The group is the one the unit runs as - **not** the service user's own group, since systemd drops
  supplementary groups when `Group=` is set explicitly.
- For local development, use .NET user-secrets (`dotnet user-secrets`), which stores outside the
  repository tree.
- Add a `.gitignore` entry for any local config file that could hold a secret, and include a
  committed `appsettings.Example.json` with placeholder values so the shape is documented.
- If a credential is ever pasted into a session transcript, treat it as compromised and ask the
  project owner to rotate it.

---

## Phase 0 - Repository restructure [DONE 2026-09-20]

**What was built.** The layout matches the target exactly: `src/CRT.App/` (csproj renamed
`CRT.App.csproj`, `AssemblyName` left as `Classic-Repair-Toolbox` per the trap warning below) and
`tests/CRT.App.Tests/` (csproj renamed `CRT.App.Tests.csproj`). `Assets/` stayed at the root as
planned. Solution, both workflows, both vscode files, both hooks, and CLAUDE.md's ~40 relative
links were all updated. Verified clean from a from-scratch build (`rm -rf */bin */obj` first):
Release build 160 warnings / 0 errors (matches the pre-move baseline exactly), full suite
3111/3111 passing, `dotnet run --project src/CRT.App/CRT.App.csproj` launches correctly, and
`remind-wiki-mirror.sh`'s MAP was live-fire tested against real changed paths.

**Two things worth the next session knowing:**

1. **This filesystem is case-insensitive** (`core.ignorecase=true`). `Tests/` and `tests/` collide
   as the same physical directory. A `git mv Tests/X tests/Y` lands correctly in git's index, but
   any subsequent `rm -rf Tests` (or `tests`) deletes the same directory either casing names -
   which happened once during this phase and required restoring 139 files from `HEAD` via
   `git show`, verified byte-for-byte identical afterward except two deliberate edits. Anyone
   scripting further moves on Windows/case-insensitive filesystems should never assume `Tests` and
   `tests` are different targets.
2. **Avalonia resource `Include` paths need an explicit `Link` when the physical file moves outside
   the project folder but the app hardcodes the logical `avares://` path.** `AvaloniaResource
   Include="../../Assets/CRT_icon.png"` alone changed the logical resource path Avalonia derives,
   breaking all four `avares://Classic-Repair-Toolbox/Assets/...` references in `Splash.axaml`,
   `TabAbout.axaml`, `App.axaml` and `WorkbookPdfExporter.cs` - it surfaced as 223 failing headless
   UI tests with `FileNotFoundException: the resource avares://.../Assets/CRT_icon.png could not be
   found`. Fixed with `Link="Assets/CRT_icon.png"` (and the equivalent for the font glob) to pin the
   logical path independent of the physical one. This is now the pattern anything moved in Phase 1
   or later must follow for any `AvaloniaResource`/`Content` whose logical path is referenced by a
   hardcoded `avares://` URI anywhere in the codebase.

Not fixed, out of scope for this phase: `.claude/CLAUDE.md`'s link to
`Assets/UI mockups/worklog-mockup.html` was already broken before Phase 0 (the file does not exist
anywhere in git history's working tree at HEAD) and is unrelated to the restructure.

**Goal.** Move from a root-level project that claims the whole tree to a `src/` + `tests/` layout
where every project owns its own folder.

**Why this must come first.** [Classic-Repair-Toolbox.csproj](../Classic-Repair-Toolbox.csproj) sits
at the repository root, so its default compile glob is `**/*.cs` from the root down. That is why it
already needs explicit `Remove` blocks for `Tests/**` and `Assets/Webserver/**`
([lines 96-111](../Classic-Repair-Toolbox.csproj#L96-L111)). Add three more projects and every one
of them needs its own exclusion, forever, each a silent trap. Doing this move with two projects is
far cheaper than with five.

**This is a move, not a rewrite.** No code logic changes. The test suite passing at the end is the
proof.

### Target layout

```
Classic-Repair-Toolbox.slnx
src/
  CRT.App/                      <- today's app (csproj + Main/ Tabs/ Handlers/ Properties/)
tests/
  CRT.App.Tests/                <- today's Tests/Classic-Repair-Toolbox.Tests/
Assets/                         <- stays at repository root, unchanged
.github/workflows/
.claude/
```

`CRT.Data`, `CRT.Server` and `CRT.Maintainer` join `src/` in later phases (`CRT.Maintainer` left it
again on 2026-09-29, when it became CRT's Maintainer tab - Phase 8).

**Keep `Assets/` at the root.** It holds the data pack, the Wiki mirror and the webserver copy -
none of it belongs to a single project, the release workflow copies it by path, and moving it would
churn a great many links for no benefit.

### Tasks

1. **Move the app project.** `Classic-Repair-Toolbox.csproj` plus `Main/`, `Tabs/`, `Handlers/`,
   `Properties/` into `src/CRT.App/`. Use `git mv` so history follows.
2. **Rename the csproj to `CRT.App.csproj`.** Optional but recommended - the folder name and project
   name should agree. If renamed, `AssemblyName` must be set explicitly to keep the output DLL name
   stable, because Velopack packaging and `launch.json` both reference it.
3. **Move the test project** to `tests/CRT.App.Tests/`. Update its `ProjectReference` and
   `RootNamespace`.
4. **Delete the now-unnecessary `Remove` blocks** for `Tests/**` from the app csproj. The
   `Assets/Webserver/**` removes are still needed while `Assets/` stays at the root, since the app
   project no longer contains it - verify by building; if `Assets/` is outside the project folder
   the removes become dead and should go too.
5. **Update `InternalsVisibleTo`** ([line 117](../Classic-Repair-Toolbox.csproj#L117)) to the test
   assembly's new name.
6. **Update the solution** [Classic-Repair-Toolbox.slnx](../Classic-Repair-Toolbox.slnx) with the
   new paths.
7. **Update `.vscode/`**: `launch.json` line 13 (`program` path to the built DLL) and `tasks.json`
   lines 10, 22, 36 (solution path).
8. **Update the workflows.** Both
   [build-and-unittest.yml](../.github/workflows/build-and-unittest.yml) and
   [build-and-release.yml](../.github/workflows/build-and-release.yml) set `SOLUTION` and (release
   only) `PROJECT`. Update those two variables and check every other path reference - the release
   workflow copies the data pack and reads `InformationalVersion` by path.
9. **Update `.claude/hooks/`.**
   - `require-green-tests.sh` line 27 sets `SOLUTION`.
   - `remind-wiki-mirror.sh` lines 59-70 hold a `MAP` of **nine** path prefixes
     (`Tabs/Workbooks/`, `Handlers/Data/...`, `Handlers/MiniPro/`, `Handlers/Oscilloscope/`,
     `Tabs/Contribute/`, `Handlers/Online/...`, `Classic-Repair-Toolbox\.csproj`). Every one needs
     its new prefix. **This is the most easily forgotten step**, and its failure mode is silent -
     the hook simply stops matching anything and never warns again.
10. **Update documentation paths**: [.claude/CLAUDE.md](../.claude/CLAUDE.md) is dense with relative
    links (`../Handlers/...`, `../Tabs/...`); also [README.md](../README.md) and the Wiki pages
    `Compiling-yourself-from-source.md` and `Development-tools-used.md`.
    Note: `BUILDING.md` is **deliberately gone** - per-OS build instructions live in the Wiki page
    `Compiling-yourself-from-source.md` instead. CLAUDE.md's stale link to it has been corrected.
    **Do not recreate that file.** The CHANGELOG mentions it historically, which is correct and must
    not be edited.
11. **Check for hardcoded relative paths in tests.** Anything walking up to the repo root to find
    `Assets/` will have a different depth now. Per project rules, tests must never hardcode a
    Windows path - they run on `ubuntu-latest` in CI too.

### A data-loss bug found by USING it, not by testing it (2026-09-21)

**Reported live:** a new system was created, two board images imported, two components labelled
in the label editor - and the board then rendered EMPTY, and stayed empty across a board switch
and a restart.

**Every row was still in `draft.json`.** What had gone was `NewSystem`, the registration that says
a system exists only as a draft. Three writers rebuild `BoardDraft` field by field
(`LabelEditorDraftWriter`, `ComponentDraftWriter` x2, `KiCadCalibrationDraftWriter` - **four**
rebuild sites), and every one copied all eleven ROW SECTIONS while silently dropping `NewSystem`,
which is not a section and so was in nobody's checklist. `DraftBaseRevision` already carried it,
which is what proves this was an oversight rather than a decision.

**Why losing it is fatal rather than cosmetic.** A draft-only system has no `.xlsx` by
construction. `BoardDataReader.LoadAsync` tolerates that only when told to, and `DataManager`
decides with `draft?.NewSystem != null`. Drop it and the system stops being recognised as
draft-only, the missing file becomes a hard error, and the board loads as nothing - so a
contributor watches their own work disappear with no error shown anywhere in the UI.

**Why no test caught it.** Each writer had thorough tests, and they all constructed a plain
`new BoardDraft()` with no registration - so the field being dropped was invisible to every one of
them. The coverage was real; it was coverage of the wrong starting state.
`DraftWriterNewSystemPreservationTests` now guards the CLASS of bug (rebuild a `BoardDraft`
anywhere and the registration must survive), including the negative case that an ordinary draft
must never ACQUIRE a registration - which would make `DataManager` tolerate a genuinely missing
`.xlsx`, hiding real sync failures.

**The lesson for this document: a draft-only system is a distinct starting state, and every test
of draft-writing logic should have one.** The pure-logic tests all began from an empty draft, which
is the state a system over published data is in - never the state a brand-new system is in.

### Definition of done

- `dotnet build Classic-Repair-Toolbox.slnx -c Release` succeeds with no new warnings.
- `dotnet test Classic-Repair-Toolbox.slnx` is green with the same test count as before.
- F5 and the `watch` task both still run the app.
- A dry run of the release workflow (or a careful read) confirms publish and Velopack paths resolve.
- Both Wiki-related hooks still fire: make a trivial edit under a mapped path and confirm
  `remind-wiki-mirror.sh` still names the expected page.

### Traps

- **The `remind-wiki-mirror.sh` MAP fails silently.** Test it deliberately, per the last bullet
  above.
- **Renaming the assembly breaks Velopack.** The pack step uses `--packId` and an output DLL name;
  if the assembly name changes, the installed application's identity changes and existing
  installations may not update. Safest: keep `AssemblyName` exactly as it is today even if the
  csproj file is renamed.
- **Do not reorganise namespaces in this phase.** Folder moves plus namespace changes in one commit
  makes review impossible and hides real breakage. Namespaces can stay as they are indefinitely.

---

## Phase 1 - Extract CRT.Data [DONE 2026-09-20]

**What was built, and two deviations from the plan below - both discovered during execution, not
guessed at in advance.**

`src/CRT.Data/` now holds `BoardData.cs`, `BoardDataReader.cs`, `BoardDataWriter.cs`,
`BoardComponentHighlightStorage.cs`, `ComponentListBuilder.cs`, `ContactLinkFormatter.cs`,
`TextLinkFinder.cs`, plus two new files the plan's "Traps" section anticipated needing:
`ICrtLog.cs` (the `CrtLog`/`ICrtLog` logging seam) and `EpplusLicense.cs` (centralises the
`ExcelPackage.License.SetNonCommercialPersonal(...)` call - task 6, done as a follow-up in the same
session rather than left for later). `CRT.App` gained `Handlers/Data/AppLoggerAdapter.cs`, installed
as `CrtLog.Sink` right after `Logger.Initialize()` in `App.OnFrameworkInitializationCompleted`, so a
moved class logs to the exact same file it always did - confirmed by running the real app against
real board data and grepping its log for `Board Excel data loaded and cached for [...]` lines
originating from the moved `BoardDataReader.LoadAsync`.

**Deviation 1 - `DataValidator` was NOT moved**, despite being in the plan's confirmed list. It
directly calls `DataManager.HardwareBoards`, `DataManager.LoadBoardDataAsync` and
`DataManager.DataRoot` - a real dependency the plan's own file-content check had not caught. Moving
it would have meant either moving `DataManager` too (a 1600+ line app-orchestration class, clearly
out of scope for "the UI-free data layer") or changing `ValidateAllDataAsync`'s signature to accept
what it needs as parameters instead of reaching out for it - which the strategy doc's own coverage
table already flags as a owner decision ("a public API change... a decision for the project owner
rather than something to do in passing"), not something to do in passing during a phase whose whole
premise is "a move, not a rewrite." `DataValidator` stayed in `CRT.App`, byte-for-byte identical to
its pre-Phase-0 content. It is a leaf (nothing else in the moved set depends on it), so this cost
nothing beyond not getting the file where the plan originally guessed it should go.

**Deviation 2 - the three "also move" geometry files were skipped**, and this was flagged to the
project owner as a genuine plan conflict before proceeding, not decided unilaterally: `RectGeometry`,
`PolygonGeometry` and `HighlightRectBuilder` are all built on `Avalonia.Point`/`Rect`/`Matrix`,
which directly contradicts the plan's own "no Avalonia reference at all" instruction for
`CRT.Data` - a constraint that exists specifically because the future server must never depend on
a desktop UI framework. Investigation showed `RectGeometry`/`PolygonGeometry` are called from 21
files across the whole Schematics/Worklog rendering pipeline, so converting them to a portable
representation is a real refactor (touching every call site), not a file move, and the project owner
agreed to defer it to whichever later phase actually needs a portable geometry library - most
likely when the maintainer app is built and has a second, concrete consumer to design the abstraction
around, rather than guessing at its shape now with only one.

**Two things the next phase should know:**

1. **`InternalsVisibleTo` on `CRT.Data.csproj` must grant `Classic-Repair-Toolbox`, not `CRT.App`.**
   `CRT.App`'s project FILE is named `CRT.App.csproj` (Phase 0), but its `AssemblyName` is still
   `Classic-Repair-Toolbox` - kept unchanged deliberately so Velopack's update identity never moves.
   Granting `CRT.App` silently grants nothing, since no assembly by that name is ever produced; this
   surfaced as `CS0122` on `TabSchematics.LabelEditor.cs`'s use of `BoardDataWriter`'s internal
   `LabelEditorSaveRow`, fixed by granting the real assembly name instead.
2. **`CRT.Data.csproj` carries no `SonarAnalyzer.CSharp` reference**, so the five SonarAnalyzer
   warnings the moved files used to contribute to `CRT.App`'s build output (all already listed in
   `WarningsNotAsErrors`, so none were ever build-blocking) simply stopped appearing - the true
   warning count went from 160 to 155, confirmed by an isolated `dotnet build src/CRT.Data/...`
   showing 0 warnings. Not a regression in anything that gates a build, but a small loss of the
   analyzer visibility CLAUDE.md describes as the whole point of that package. Left as a follow-up
   rather than fixed in this phase, since adding the analyzer means tuning a fresh
   `WarningsNotAsErrors` list for a second project - its own small piece of work, not a natural part
   of "extract this data layer."

Verified clean from a from-scratch build (`rm -rf */bin */obj` first): Release build 155 warnings /
0 errors, full suite 3111/3111 passing (195 in `CRT.Data.Tests` + 2916 in `CRT.App.Tests` - the
exact same total as before the phase, split across two assemblies), a self-contained
`dotnet publish` correctly bundles `CRT.Data.dll` alongside the app executable, and the running app
was launched twice against real production board data with confirmed successful loads in its log.
Both hooks (`remind-wiki-mirror.sh`'s MAP, `require-green-tests.sh`) and the duplication-check
workflow step were updated for the new `src/CRT.Data/` path and live-fire tested.

**Goal.** Move the UI-free data layer into a shared `net10.0` class library that the app, the
server and the maintainer app all reference.

**Why this is the highest-value single change.** It retires the two-language duplication described
in [Why this is being done](#why-this-is-being-done) problem 4. Verified during planning: these
classes contain **no Avalonia references at all**, so the extraction is a move rather than a
rewrite.

Confirmed UI-free and ready to move:

- `Handlers/Data/BoardData.cs`
- `Handlers/Data/BoardDataReader.cs`
- `Handlers/Data/BoardDataWriter.cs`
- `Handlers/Data/BoardComponentHighlightStorage.cs`
- `Handlers/Data/DataValidator.cs`

Also move, as they are pure and will be needed server-side: `ComponentListBuilder`,
`ContactLinkFormatter`, `TextLinkFinder`, and the parts of `Handlers/Geometry/` the maintainer app will
need for previews (`PolygonGeometry`, `RectGeometry`, `HighlightRectBuilder`).

**Do not move** `WorklogManager`, `WorklogEntryScope` or `WorkbookPdfExporter` - all three reference
Avalonia, and none is needed by the server.

### Tasks

1. Create `src/CRT.Data/CRT.Data.csproj`, plain `net10.0`, **no Avalonia reference at all**. Add
   `EPPlus`. Set `TreatWarningsAsErrors` for Release, matching the other projects.
2. `git mv` the listed files in. Keep namespaces unchanged (`Handlers.DataHandling`) so no call site
   moves - a namespace tidy-up can happen later on its own if wanted.
3. Add a `ProjectReference` from `CRT.App` to `CRT.Data`.
4. Create `tests/CRT.Data.Tests/` and move the tests for the moved classes across from
   `CRT.App.Tests`. Add `InternalsVisibleTo` on `CRT.Data` for it.
   - **Watch the xUnit collections.** Per CLAUDE.md, `BoardDataReaderTests` and
     `BoardDataWriterTests` share the `"BoardData"` collection because of a shared static cache.
     They must stay in the same collection **and the same project** - splitting them across
     projects breaks the serialisation guarantee and produces intermittent failures.
     `Tests/.../xunit.runner.json` must be copied to the new test project too.
5. Add `CRT.Data` and its test project to the solution.
6. **Resolve the EPPlus licence call.** `ExcelPackage.License.SetNonCommercialPersonal(...)` appears
   in several files. Centralise it in one internal initialiser inside `CRT.Data` called from a
   static constructor or an explicit `Initialize`.
   **Flag to the project owner:** that licence mode is chosen for a desktop application used by an
   individual. Whether it also covers running EPPlus inside a server process is a licence question,
   not a technical one, and must be confirmed before Phase 3 ships anything.
7. Confirm `DataValidator` carries no UI assumptions once isolated (it reports through `Logger`;
   see the logging note in [Cross-cutting concerns](#cross-cutting-concerns)).

### Definition of done

- `CRT.Data.csproj` has no Avalonia dependency, direct or transitive.
- Whole solution builds Release with no new warnings; full suite green with the same count.
- The app runs and loads board data exactly as before.

### Traps

- **`Logger` is a static singleton in the app.** `CRT.Data` must not depend on the app's logger.
  Introduce a minimal `ICrtLog` inside `CRT.Data` with a no-op default; the app supplies an adapter
  at startup. CLAUDE.md's rule that no test may call `Logger.Initialize()` still stands.
- **Do not widen accessibility for convenience.** CLAUDE.md is explicit that `internal` is fine and
  `InternalsVisibleTo` costs no coverage. Keep types `internal` unless genuinely consumed outside.
- **This phase is worth doing even if the project stops here.** Say so in the session summary.

---

## Phase 2 - Local-first drafts in CRT [DONE 2026-09-20]

**All four sessions (2a, 2b, 2c, 2d) are complete, and every item in the Definition of done below
is verified rather than assumed.** A contributor can now author a complete new system with no server
involved, see every local edit marked on the board, toggle to the officially published view, import
board images and KiCad data with a report saying what will actually light up, and be warned - with a
per-row view - when the official data moves underneath a draft. Nothing in the app writes to `Data/`
or creates a `_UserContribution` sidecar any more; `Drafts/` is the only mechanism new authoring
writes through, and a sync provably cannot destroy a draft.

**Phase 3 is the next one, and it is the first to need the project owner access listed under "Access
the project owner must provide"** (MariaDB, server shell, secrets). Phases 0-2 needed none of it.

**Goal.** A contributor's edits are saved locally, survive sync, render on top of official data,
and are clearly marked as pending.

This is the largest app-side phase and will span several sessions. Split it as marked.

### The model

A second root, `Drafts/`, sits beside `Data/` in AppData and mirrors its structure exactly. Edits
are written **only** there. Sync continues to overwrite `Data/` freely and can never touch a draft.

At load time, `BoardDataReader` reads the official board and applies the draft over it:

- a row present in both: draft values win;
- a row only in the draft: added;
- a row the draft marks deleted: hidden.

A draft records the **base revision** of the system it was started from, for the drift warning
below and for the submission in Phase 4.

### Session 2a - Draft storage and the overlay read path [DONE 2026-09-20]

**What was built.** All four tasks below, exactly as scoped - no deviations from the plan this
time. New files: `src/CRT.Data/BoardDraft.cs` (the `BoardDraft`/`DraftRow`/`DraftRowState` model),
`src/CRT.Data/BoardDraftNaturalKeys.cs` (one key-builder per section, matching the keys this
document already named for Phase 4's "Retiring the UUIDs" - label+region+pin+name for a component
image, schematic+label for a highlight, etc.), `src/CRT.Data/BoardDraftApplier.cs` (the pure
`ApplyDraft(BoardData, BoardDraft?)` merge, generic over row type so all ten `BoardData` sections
share one `Merge<T>` rather than ten hand-rolled copies), `src/CRT.Data/DraftDataStore.cs`
(`draft.json` read/write, with its own typed on-disk DTO shape since `DraftRow.Entry` is `object`
at runtime and `System.Text.Json` cannot round-trip a bare `object` back to its real type), and
`src/CRT.App/Handlers/Data/DraftManager.cs` (root resolution - `--drafts-root=`, parsed identically
to `--data-root=`/`--workbooks-root=` - plus mapping an `ExcelDataFile` to its folder under
`Drafts/`). `AppConfig.DraftsFolderName` added beside `WorklogFolderName`.

**The read-path wiring** is `BoardDataReader.LoadAsync` gaining an optional `BoardDraft? draft`
parameter, applied via `BoardDraftApplier.ApplyDraft` on the way OUT - on both the cache-hit and
the fresh-read path - rather than folded into what gets cached. This was a deliberate design choice
beyond what the task list spelled out: `_cache` must only ever hold official data, or a draft edit
would need to invalidate the cache to take effect and the cache would need a second key per draft
state. `DataManager.LoadBoardDataAsync` resolves the draft via `DraftManager.LoadDraftFor` (re-read
from disk on every call, not cached - `draft.json` is small and this is not a hot path) and passes
it through. `DraftManager.Load` is wired into `App.axaml.cs` startup unconditionally, alongside
`WorklogManager.Load`, reading the same raw process args for the same reason (not gated on the
desktop-lifetime branch).

**"Sync never destroys a draft" is true by construction, not by a guard that could be forgotten**:
`Drafts/` and `Data/` are different trees, `OnlineServices`' sync code has no reference to
`DraftManager` at all, and the overlay is applied only at the point `BoardDataReader.LoadAsync`
returns - nothing sync-related touches `_draftsRoot` or `draft.json`. Proven directly by
`DataManagerDraftOverlayTests` and `BoardDataReaderTests`' new cases (one simulates a real sync: a
cache clear plus an official-file rewrite, with a draft still applying to the result afterward).

**Two things the next session should know:**

1. **`BoardDraftApplier.Merge<T>`'s `KeyOf<T>` switch is the one place a new `BoardData` section
   would need a matching arm.** If Phase 2 later adds a section `BoardData` does not have yet (none
   currently planned), both `BoardDraftNaturalKeys` and this switch need updating together, or the
   generic merge throws `NotSupportedException` for that row type.
2. **No UI exists yet for any of this** - no pending markers, no "view as published" toggle, no
   Drafts list, no way for a human to actually create a draft row through the app. A draft can
   currently only be constructed by hand (or by a test) via `DraftDataStore.Save`. Session 2b picks
   up the pending-markers UI; **Session 2c ("Authoring into drafts") is where this becomes
   reachable by a real contributor** - until then this is read-path plumbing only, verified true by
   its 60 new unit tests (`BoardDraftApplierTests`, `BoardDraftNaturalKeysTests`,
   `DraftDataStoreTests`, `DraftManagerTests`, `DataManagerDraftOverlayTests`, plus new cases in
   `BoardDataReaderTests`), not by anything visible in the running app yet.

Verified: `dotnet test Classic-Repair-Toolbox.slnx -c Release` green at 3171/3171 (3111 baseline +
60 new), 0 warnings beyond the pre-existing accepted set. `Commandline-parameters.md` updated in
the same change (a fourth parameter, `--drafts-root`); `remind-wiki-mirror.sh`'s MAP extended to
include `DraftManager.cs` under the same `Commandline-parameters` mapping `DataManager.cs` already
has, and live-fire tested (the hook did not re-flag `Commandline-parameters` since it was already
updated in this change, confirming the mapping and the doc edit are in sync).

1. Add `--drafts-root=<path>` alongside the existing `--data-root=`/`--workbooks-root=` switches,
   parsed identically (case-insensitive, quotes stripped, first match wins, unknown args ignored).
   Default: an AppData folder beside the data root. Document it in
   `Assets/Wiki/Commandline-parameters.md` **in the same commit**.
2. Define the draft file format. Recommended: a `draft.json` per system holding the changed rows
   plus base revision and a per-row state (`Modified` / `Added` / `Deleted`) - **not** a copy of the
   board Excel, which would be far harder to diff and to show.

   > **REVERSED ON 2026-09-23 - see [Phase 6a](#phase-6a---drafts-as-real-board-folders-done-2026-09-23).**
   > This is what was built and it worked, but it made a draft a COMPUTATION rather than a
   > file: there was nothing a contributor could open. The project owner asked for a draft to be
   > a real board folder, so the row deltas are gone and the workbook IS the draft. Do not
   > reinstate `draft.json` on the strength of this task.
3. Implement the overlay merge in `CRT.Data` as a pure function:
   `BoardData ApplyDraft(BoardData official, BoardDraft draft)`. Pure, unit tested, no UI.
   Cover: modify, add, delete, empty draft, draft for a row that no longer exists officially,
   conflicting keys, and every section.
4. Wire the read path so board loading applies any draft for that system.

### Session 2b - Pending markers in the UI [DONE 2026-09-20]

**What was built.** All three tasks, plus one piece of underlying plumbing the task list did not
spell out but turned out to be the real prerequisite for all three: `BoardDraftApplier.ApplyDraft`
deliberately discards which merged row came from the draft (the merged `BoardData` is meant to be
indistinguishable from a board that was always this way - see its own header comment), so nothing
downstream had a cheap way to answer "is this row drafted" without redoing the merge. Rather than
change the merge itself, `src/CRT.Data/BoardDraftSummary.cs` builds a second, small answer straight
from the `BoardDraft` (a `HashSet` of drafted Components board labels, another of drafted
ComponentHighlights keys) - `DataManager.LastLoadedDraftSummary` exposes it per board load, and
`Main._currentBoardDraftSummary` snapshots it alongside `_currentBoardData` the same way.

**Task 5, the chip.** `ComponentListItem.IsDrafted` (new, `src/CRT.Data/ComponentListBuilder.cs`)
and a `draftedBoardLabels` parameter on `BuildComponentItems`; `ComponentFilterListBox` in
`Main.axaml` got its first-ever `ItemTemplate` (it rendered through bare `ToString()` before,
which cannot mix with a conditional chip) - a small amber "Draft" pill via two new theme keys,
`Draft_Chip_Bg/Fg/Border`. All five call sites that rebuild that list (`Main.BoardSelection.cs`
x2, `Main.ComponentPopup.cs` x3) now pass `_currentBoardDraftSummary.DraftedComponentBoardLabels`.

**Task 5, the tint.** The harder half - flagged to the project owner mid-session as touching a hot
render path, and confirmed to proceed rather than defer. `HighlightSpatialIndex` (previously a
flat, positional `Rect[]` with no per-item metadata at all, and previously untested) gained a
parallel `bool[]? isDrafted` and `GetIsDrafted(int)`; `SchematicHighlightsOverlay.Render` now does
a second draw pass for drafted rects using a new `DraftedHighlightColor` (wired to the theme's
`Draft_Highlight_Tint`); `TabSchematics.UpdateHighlightsForComponents` builds the parallel array
from a new `draftedHighlightKeys` field, set by `Main.SetComponentHighlightRects` - the one place
`highlightRectsBySchematicAndLabel` is ever assigned - from `_currentBoardDraftSummary` again.

**Task 6, the toggle.** `UserSettings.ViewOfficialPublishedOnly` (opt-in, default false - unlike
`EnableWorklog`, a draft that exists should normally be VISIBLE), a Configuration tab checkbox,
and `DataManager.LoadBoardDataAsync` passing `null` instead of the resolved draft to
`BoardDataReader.LoadAsync` when it is on - the ONE place the overlay is applied, so no caller
needed its own check. `LastLoadedDraftSummary` still reports the true draft even while suppressed
(for a future "N rows hidden" banner); `Main._currentBoardDraftSummary` is forced to `Empty` in
that case instead, so the chip/tint cannot contradict what the toggle promises. Toggling it
re-runs the current board's load immediately via a reload helper shared with
`ApplyCatalogueVisibility` (renamed in spirit, not in name - see its own updated comment).

**Task 7, the Drafts tab.** A genuinely new conditional tab (same `IsVisible="False"` +
`Apply*Visibility()` pattern as Oscilloscope/Workbooks), hidden unless at least one system has a
non-empty draft. `DraftManager.EnumerateDraftedSystems(hardwareBoards)` (new - nothing enumerated
drafted systems before this) matches the known board list against `Drafts/` rather than trying to
walk the folder tree back into board identities, since `Drafts/` carries no manifest of its own by
design. `DraftManager.DiscardDraft` deletes the whole system folder (not just `draft.json`), same
"one folder is the whole record" model `WorklogManager.DeleteWorkbook` already uses. Discarding
goes through a new `DiscardDraftWindow` confirmation modal - a direct copy of `DeleteWorkbookWindow`'s
shape, including the Tunnel-route Enter/Escape-both-cancel fix that modal exists for.

**What this session deliberately did NOT touch:** nothing here lets a user create or edit a draft
- `TabDrafts` can only show and discard whatever `draft.json` already contains. Session 2c
("Authoring into drafts") is what makes a draft reachable by a real contributor; until then every
test in this session constructs a `BoardDraft` directly, the same as session 2a's.

**One real bug found and fixed while adding tests, unrelated to the feature itself:**
`UserSettings.LoadFrom` does not reset the static `_data` object at all when the target path does
not exist yet (only the deserialize-and-replace branch does) - a test redirecting to a fresh,
not-yet-written temp path left the PREVIOUS test's settings in memory, which surfaced as
`ViewOfficialPublishedOnly` leaking between tests in `DataManagerDraftOverlayTests`. Fixed by
writing `{}` before calling `LoadFrom`, forcing the reset branch - the same pattern
`UserSettingsTests.LoadSettings` already used for the same reason. Worth knowing for any future
test that redirects `UserSettings` to a path it has not written yet.

Verified: `dotnet test Classic-Repair-Toolbox.slnx -c Release` green at 3232/3232 (3171 baseline +
61 new: `BoardDraftSummaryTests`, `ComponentListBuilderTests` additions, `HighlightSpatialIndexTests`,
`DraftedHighlightMarkingTests`, `DraftPaletteTests`, `UserSettingsTests`/`DataManagerDraftOverlayTests`
additions, `DraftManagerTests` additions, `DiscardDraftWindowTests`, `TabDraftsTests`,
`TabConstructionTests` addition), 0 new warnings. `Configuration-tab.md` updated in the same change
(the new "Drafts" section); `Contribute-data-via-CRT.md` deliberately left untouched - the
`remind-wiki-mirror.sh` flag on it this session was a leftover from Phase 0/1's untracked move, not
a real signal, since nothing in this session touched the Contribute tab itself.

### Session 2c - Authoring into drafts [DONE 2026-09-20 - tasks 8, 9, 10 and 11 all complete]

**What was built so far.** The label editor half of task 8 - the highest-value, everyday-use piece
- redirected from mutating the board's live `.xlsx`/JSON sidecar to writing `draft.json` rows.
KiCad calibration's read/write redirect (the other genuinely "point an existing editor at the
draft root" piece of task 8) is also done. Component add/edit and file/link attachment - the
"genuinely new UI" piece - is now done too, by REUSING the existing `ComponentContributionWindow`
wholesale and redirecting only its save path; see the dedicated section below. Only task 9 ("Add a
new system") and session 2d remain - see "What remains" below.

**The real design decision this session made, flagged to and confirmed by the project owner before
writing any code:** `BoardDataWriter.SaveLabelEditorChangesAsync` mutated the officially-synced
`.xlsx` directly with EPPlus (inserting Components rows) and called
`BoardComponentHighlightStorage.SaveComponentHighlights` for the JSON sidecar - a completely
different persistence model from the row-based `draft.json` sessions 2a/2b already built. Pointing
it at an `.xlsx` under `Drafts/` instead (the literal reading of "point the editor at the draft
root") would have created a SECOND, inconsistent draft mechanism nothing else in the app
understands - `TabDrafts`, the chip, the tint and `BoardDraftApplier` all read the row-based model.
Confirmed: rewrite the writer to emit `DraftRow` entries instead, so there is still only ever one
draft mechanism.

**What was built:**
- `src/CRT.Data/LabelEditorDraftWriter.cs` (new) - pure `ApplyLabelEditorSave(official, currentDraft,
  schematicName, rows, region) -> BoardDraft`, replacing `BoardDataWriter` entirely (deleted, along
  with its now-obsolete `BoardDataWriterTests.cs` - 18 replacement tests in
  `LabelEditorDraftWriterTests.cs`). `LabelEditorSaveRow` (the DTO) moved into this new file.
  Reproduces the old writer's two real behaviours in the new shape: a genuinely new board label
  gets a drafted `Components` row (same blank-region-is-a-wildcard "already exists" rule); a
  schematic's highlight save is a **schematic-scoped replace**, not a per-row upsert - every
  OFFICIAL highlight on that schematic not in the new save gets a `Deleted` tombstone (a
  drafted-only row not in the new save is simply dropped instead - a tombstone only ever matters
  paired with an official row, see `BoardDraftApplier`'s own comment).
- **The one correctness rule the whole rewrite hinges on**: `BoardDraftApplier.Merge<T>` requires
  `Modified` for a row that already exists officially and `Added` for one that does not - an
  `Added` row whose key COLLIDES with an existing official one is silently DROPPED by the merge,
  never substituted. So `ApplyLabelEditorSave` takes the PURE OFFICIAL `BoardData` (not
  `Main.CurrentBoardData`, which already has the draft merged in and would collapse this
  distinction) to decide Modified vs Added per row correctly. `TabSchematics.LabelEditor.cs`
  resolves it via `BoardDataReader.LoadAsync(excelPath, cacheKey)` with no draft argument - a cache
  hit, since `DataManager.LoadBoardDataAsync` already populated that exact cache entry when the
  board loaded, so this costs nothing extra. Two tests
  (`The_drafted_highlight_actually_survives_BoardDraftApplier_ApplyDraft`,
  `Removing_a_rect_actually_hides_it_after_BoardDraftApplier_ApplyDraft`) prove the round trip
  through the REAL merge, not just that `ApplyLabelEditorSave` produced a row with the right state.
- `TabSchematics.LabelEditor.cs`'s `ApplyLabelEditorChanges` rewritten: resolves the current draft
  via `DraftManager.LoadDraftFor`/`GetSystemFolder`, calls the new writer, saves via
  `DraftDataStore.Save`, then reloads the board exactly as before (which re-applies the draft
  overlay through the ordinary `DataManager.LoadBoardDataAsync` path - no separate draft-aware
  reload needed).
- **A real, unrelated-looking bug fell out of the redesign and was fixed in the same change**:
  the save path used to call `Main.DisableLaunchDataSyncAfterLocalBoardEditAsync()`, which turns
  off launch-time sync and shows a scary warning dialog - because a save used to mutate the
  SYNCED file directly, so the next sync could silently overwrite it. A draft lives entirely
  outside the synced tree by construction (`Drafts/`, never `Data/`), so that risk no longer
  exists and the warning would now fire for something that cannot happen. The call is removed;
  the underlying `DisableLaunchDataSyncAfterLocalBoardEditAsync`/banner machinery in
  `Main.DataSyncStatus.cs` is left in place rather than swept out, since it now has zero callers
  and a wider dead-code removal is a separate concern from this session's actual scope - worth a
  follow-up look.
- `BoardDraft`/`DraftDataStore` gained an **eleventh section, `KiCadCalibrations`**
  (`KiCadCalibrationEntry` in `BoardData.cs` - schematic name, CAD name, offset/scale/mirror,
  mirroring `BoardComponentHighlightStorage`'s existing JSON field shape exactly), confirmed with
  the project owner as the right call since KiCad calibration had NO representation in `BoardDraft` at
  all before this. `BoardDraftTests.cs` (new) pins `IsEmpty` across all eleven sections via
  reflection so a twelfth section added later cannot be forgotten from that check silently.
- **KiCad calibration's read/write redirect, built later the same session**:
  `src/CRT.Data/KiCadCalibrationDraftWriter.cs` (new) - `BoardDraftNaturalKeys`/`BoardDraftApplier`
  were deliberately NOT extended for this, since `BoardData` itself does not carry
  `KiCadCalibrations` (calibration is read on-demand per schematic, never as part of an official
  `BoardData` load) - there is no `BoardData` field for the generic `Merge<T>` to fold into. Instead
  this is a small parallel path with the same "draft wins by natural key" rule everything else
  follows: `ResolveEffectiveCalibration(draft, schematicName, hasOfficial, ...official fields...)`
  takes the official JSON's own out-parameter shape straight from
  `BoardComponentHighlightStorage.TryLoadKiCadCalibration` (no adapting needed at the call site),
  looks for a drafted row keyed on `BoardDraftNaturalKeys.ForSchematic` (one calibration per
  schematic, so no schematic+label compound key is needed the way highlights need), and returns the
  drafted entry if one exists and is not a `Deleted` tombstone, else the official value, else
  `null`. `ApplyCalibrationSave(currentDraft, schematicName, ...)` upserts (via
  `LabelEditorDraftWriter.UpsertDraftRow`, made `internal` and reused rather than duplicated) one
  `KiCadCalibrationEntry` row and leaves every other section, and every other schematic's
  calibration row, untouched. Always writes `State = Added` - there is no official `BoardData`
  argument here to check "does this key already exist officially" against the way the label
  editor's writer does, and `BoardDraftApplier.Merge<T>` treats `Added`/`Modified` identically when
  a key COLLIDES with an official row (both substitute), so `Added` alone is correct.
  `KiCadCalibrationDraftWriterTests.cs` (new, 15 tests) covers both directions.
  `TabSchematics.KiCad.Render.cs` gained `ResolveEffectiveKiCadCalibration(schematicName)`, the one
  place all three read call sites (`GetKiCadViewCalibration`,
  `BeginKiCadTraceCalibrationMode`'s saved-mirror-flags check, `BuildKiCadCalibrationImageBounds`)
  now go through instead of calling `BoardComponentHighlightStorage.TryLoadKiCadCalibration`
  directly; the one write call site (`ApplyKiCadTraceCalibration`) now resolves the draft via
  `DraftManager.GetSystemFolder`/`LoadDraftFor` and saves through `DraftDataStore.Save`, the same
  shape `TabSchematics.LabelEditor.cs`'s save already uses.
  `BoardComponentHighlightStorage.SaveKiCadCalibration` is unchanged and still tested directly - it
  is still a legitimate, correct primitive for writing the official JSON, it is simply no longer
  called from the app's own editor write path (same relationship `SaveComponentHighlights` already
  has after the label editor's redirect).

**Component add/edit and file/link attachment, the last piece of task 8, built the same session.**
Confirmed with the project owner this was not going to be a network gap needing its own architecture
question (see below) before writing code: **reuse `ComponentContributionWindow` wholesale rather
than build a new editor.** It already had every field, every validation rule, the file picker, the
category-suggestion `AutoCompleteBox`, and the "delete this component" mode task 8 needed - the
only thing wrong with it was where a save WENT (a zip, POSTed to `AppConfig.ContributionUploadUrl`).
So this piece is a redirect exactly like the label editor's, not new construction, despite the
"genuinely new UI" framing task 8 originally carried (that framing predates this session's
research into what the window already had).

**A real architecture gap had to be closed first, and the project owner explicitly delegated the
design to the agent ("investigate and figure out what is the best-possible solution") rather than
picking between the two options offered:** `ComponentEntry.File`/`ComponentImageEntry.File`/
`BoardLocalFileEntry.File` are relative paths, and EVERY existing consumer resolves one as a
single `Path.Combine(DataManager.DataRoot, entry.File)` - confirmed via research across
`ComponentInfoWindow.axaml.cs` (x2), `TabOverview.axaml.cs` and `Main.BoardSelection.cs`, each
reimplementing the same one-root combine by hand, with zero existing "also check Drafts/"
fallback anywhere. A drafted attachment's bytes cannot live under `Data/` (sync-owned, freely
overwritten), so a second root had to exist and every real consumer had to be taught to check it.
**Resolved as `src/CRT.Data/DraftFileResolver.cs`** (new, pure): `Resolve(dataRoot,
draftSystemFolder, relativeFile)` tries the official path first, then
`<draftSystemFolder>/Files/<relativeFile>`; `BuildDraftFileDestination` is the matching write-side
helper. A drafted attachment's bytes are copied on save into that `Files/` folder under the SAME
relative path the entry's own `File` value names, so the stored entry itself needs no
draft-awareness - only how it is opened does.

**The three real app-side consumers were rewired to go through it in a follow-up turn**, after this
document briefly claimed the wiring was already done when only the write side
(`CopyNewlyAttachedFilesIntoDraft`) actually existed - caught by re-checking the claim against the
code rather than trusting the earlier summary:
`ComponentInfoWindow.axaml.cs`'s `SetComponent`/`LoadImagesAsync` (component images, component local
files - `SetComponent` gained a `draftSystemFolder` parameter, and `ComponentLocalFileItem` gained a
`SourceRoot` field recording which root a file actually resolved against, since a drafted file's
absolute path is never under `DataRoot` and `ExternalTargetLauncher`'s containment check needs the
right root passed as its override), `TabOverview.axaml.cs`'s `OnLinkClick` (component/board local
file links, via `Main.GetCurrentBoardEntry()`), and `Main.BoardSelection.cs`'s schematic-thumbnail
load (`DraftManager.GetSystemFolder(entry.ExcelDataFile)` computed once per board load). Covered by
three new tests in `ComponentInfoWindowTests.cs` (draft-only file resolves, official copy still wins
over a same-named draft one, neither-exists still falls back to the ordinary DataRoot path) -
`TabOverview.OnLinkClick` and `Main.BoardSelection.cs`'s thumbnail load stay untestable for the same
reason `OnLinkClick` already was (`Process.Start`/no headless `Main`), so their fix rides on
`DraftFileResolver`'s own existing coverage rather than a new UI test.

A fourth pre-existing gap was found but left alone as explicitly out of scope:
`TabOverview.OnLinkClick` opens a component local file via a raw `Process.Start` rather than through
`ExternalTargetLauncher`, unlike every other consumer - a genuine, unrelated security hardening gap,
noted for a future pass, not fixed here.

**`src/CRT.Data/ComponentDraftWriter.cs`** (new) is the third pure writer, alongside
`LabelEditorDraftWriter` and `KiCadCalibrationDraftWriter`. Its inputs deliberately mirror
`ComponentContributionWindow`'s own row DTOs field-for-field rather than defining a third parallel
shape, since the window already builds those exact same DTOs for its old payload. Two things make
it differ from the other two writers: `ApplyComponentSave` handles six sections
(`Components`/`ComponentImages`/`ComponentLocalFiles`/`ComponentLinks`/`BoardLocalFiles`/
`BoardLinks`) in one call, since the window edits and saves them together; and per-row upsert by
key rather than a schematic-scoped "replace everything" the way the label editor's highlights work
- a row missing from the new save is simply a row the "Remove" button removed from the
`ObservableCollection`, read directly as a `Deleted` tombstone. `ApplyComponentDelete` is a
SEPARATE entry point for "Delete this component" mode, not `ApplyComponentSave` called with empty
rows, because a whole-component delete also has to tombstone the `Components` row itself - which
an ordinary save that never touched that section would not do.
`ComponentDraftWriterTests.cs` (new, 37 tests) covers both directions, including two tests that
prove the tombstone round-trips through the real `BoardDraftApplier.ApplyDraft` merge, matching the
label editor's own "prove it through the real merge, not just the writer's output" pattern.

**`ComponentContributionWindow`'s save path was rewritten, its UI otherwise left alone.**
`OnSubmitClick`/`ProcessAndSendContributionAsync`/`BuildPayload`/`AssignZipEntriesToPayload`/
`AddTextEntryToZip`/`AddFileToZipSafe` (the zip-and-POST machinery) are gone, replaced by
`SaveComponentToDraftAsync`: resolves the pure official `BoardData` the same way the label editor's
save does (`BoardDataReader.LoadAsync` with no draft argument - a cache hit), copies any newly
picked files into the draft via `DraftFileResolver.BuildDraftFileDestination`
(`CopyNewlyAttachedFilesIntoDraft` - skips a row whose file was never touched, i.e. no
`OriginalFilePath`, since its bytes already live wherever they were loaded from), builds the
updated draft via `ComponentDraftWriter`, and saves via `DraftDataStore.Save`. **The email field and
the mandatory-comment requirement are gone with the network model** - a draft is not submitted to
anyone yet, so neither has anything to validate; the comment box stays as an optional "Note" the
contributor can use for their own future reference. The Save button no longer locks after a
successful send (`ApplySubmissionOutcome`) - saving a draft again is completely ordinary, unlike the
old one-shot network submission. Every user-facing string that described a "suggestion sent to the
developer" was reworded to describe a local draft that takes effect immediately and is only
submitted later (window title, notice text, delete-mode notice, success line).

**`ContributionPackaging`'s zip/feedback helpers are now dead code, left in place as a flagged
follow-up rather than swept out in this change** (same treatment `Main.DataSyncStatus.cs`'s
disable-sync machinery got after the label editor's redirect): `ContributionFileReference`,
`ContributionAttachment`, `ContributionZipPlan`, `AssignZipEntries` and `BuildFeedbackText` have
zero callers left in `src/CRT.App` outside their own definition file, but still have dedicated
tests (`ContributionPackagingTests.cs`) and are exactly the shape Phase 4 ("Submission pipeline")
will need again once drafts can actually be submitted - removing them now would mean rebuilding
the same thing later. `TryParseOutdatedVersionResponse`/`BuildDeleteComponentSummary`/
`ValidateNewComponent`/`ValidateComponentImageFile`/`ResolveExistingFilePath`/
`TryGetDataRootRelativeFolder` are all still called from the window and stay exactly as they were.

**Three headless UI test files needed a real rewrite, not just a search-and-replace**, since they
were built entirely around the network model: `ComponentContributionValidationTests.cs`,
`NewComponentContributionTests.cs`, `DeleteComponentContributionTests.cs`. Row/field validation,
delete-mode UI structure, header/notice text and the status-area rendering are all unaffected and
were kept; every test that drove `BuildPayload`, the email field, or the mandatory-comment
requirement was rewritten against the new row-building helpers
(`BuildComponentDraftRows`/`BuildComponentImageDraftRows`/etc, reached the same way `BuildPayload`
used to be - by reflection) or, for the deletion round-trip, against `ComponentDraftWriter`
directly. **What a real SAVE writes to disk is deliberately not exercised at the UI layer** - same
as `TabSchematics.LabelEditor.cs`'s own `ApplyLabelEditorChanges`, which has no dedicated headless
test either, because it needs a real board Excel file on disk that this suite does not construct.

Two Wiki pages rewritten for the same reason `Board-JSON.md` was in the earlier half of this
session: `Contribute-tab.md` and `Contribute-data-via-CRT.md` both said an edit here "sends a
suggestion to the developer" and "your own copy does not change" - both now describe the local
draft that takes effect immediately, name the "Drafts" tab and the "View boards as officially
published" toggle, and the walkthrough's step 5/6 renamed from "write the mandatory comment" /
"enter your email and send" to "write a note (optional)" / "save".

#### Task 9 - "Add a new system" [DONE 2026-09-20]

Built in full (both the registration/creation half and the import/report half) in one further
session, after an Opus planning pass whose investigation turned up materially more dependency
surface than the task list named. The plan itself recommended splitting it in two; the project owner
chose to build both halves together.

**The registration lives INSIDE `draft.json`, not in a registry file** - a new
`NewSystemRegistration` block on `BoardDraft` (hardware name, board name, notes, the ExcelDataFile
identity, a created-UTC stamp). Evaluated against a `Drafts/registry.json` alternative and rejected
on one decisive point: `DraftManager.DiscardDraft` deletes the whole system FOLDER ("one folder is
the whole draft"), so a registration stored inside it is retired atomically, while a registry file
would be left pointing at a folder that no longer exists unless a second, easily-forgotten write
kept it in step. It also keeps `DraftManager`'s own written-down invariant that "Drafts/ carries no
manifest of its own", and hands Phase 4 the identity inside the draft it is already submitting.
Discovery walks the tree instead (`DraftManager.EnumerateDraftOnlySystems`, bounded to the three
Manufacturer/Hardware/Board levels), which is self-validating: what is on disk IS the list.

**`BoardDraft.IsEmpty` now returns false when a registration is present**, before a single row
exists. A freshly created system has zero rows in all eleven sections, and without this it would be
excluded from `EnumerateDraftedSystems`, never listed on the Drafts tab, and the tab itself (hidden
when empty) would never appear - leaving no way to discard it. The `BoardDraftTests` theory
enumerates the `List<DraftRow>` sections by reflection and cannot see a non-list property, so
`NewSystem` carries its own explicit test.

**`ExcelDataFile` for a draft-only system is a real-looking path that is never created**:
`"<Manufacturer>/<Hardware>/<Board>/Data <Hardware> <Board>.xlsx"` (`NewSystemIdentity`). It has to
be folder-shaped because `DraftManager.GetSystemFolder` needs two segments and
`HardwareBoardEntry.ShortHardwareBoardLabel` needs three. No version suffix, deliberately - keeping
one in step with the resolved main workbook is exactly the burden this task removes.

**Four existing code paths silently broke for a draft-only system and were fixed as part of this**,
none of them named in the original task list:
- `BoardDataReader.LoadAsync` returned null on a missing file BEFORE looking at the draft. It now
  takes `allowMissingOfficialFile` and merges the draft onto `new BoardData()` instead. The flag is
  set from the DRAFT'S OWN registration, never from "the file happens to be missing" - for a system
  the main workbook lists, a missing file is a real sync failure that must keep failing loudly.
  Nothing is cached in that branch (`_cache` holds officially parsed data only).
- `ComponentContribution.SaveComponentToDraftAsync` and `TabSchematics.LabelEditor`'s
  `ApplyLabelEditorChanges` both resolved the pure official board first and gave up when it was
  null - so the first component added to, and the first label drawn on, a brand-new system would
  both have failed. Both now resolve the draft first and pass the flag through.
- `Main.GetCurrentBoardKiCadRawPaths` derived its folder from the (nonexistent) official path, so
  imported KiCad data would never have been read. It now searches both roots, official first, the
  same rule `DraftFileResolver` applies to every other board file.

**Sync and validation bookkeeping learned `HardwareBoardEntry.IsDraftOnly`** (a stored flag, not an
inferred `File.Exists`, so four sites cannot disagree): the two `boardExcelFiles` sets that drive
sync skip these, and `DataValidator` skips them entirely rather than reporting every file of every
drafted system as missing on each launch. The orphan-cleanup path was checked and needs no guard -
its callers read entries straight from the workbooks rather than from `HardwareBoards`, and its
`thisTryAddMappedFilesFromBoardExcel` returns early on a missing file - so the plan's concern about
it failing closed for everyone turned out not to apply. Verified rather than assumed.

**`DataManager.RefreshDraftOnlySystems`** rebuilds only the draft-only half of `HardwareBoards`, so
creating or discarding a system takes effect without a restart and without re-reading the main
workbook (which would also reset `Oscilloscopes` and the protected-file set). Called from the create
path and from `TabDrafts`' discard - the latter is easy to miss and would otherwise leave a
discarded system in the drop-downs pointing at a deleted folder.

**No `.xlsx` is written anywhere, and this is the design, not an omission.** Task 9's wording
("create the board workbook from a schema the app already owns") is satisfied by `draft.json`, whose
sections ARE `BoardData`'s schema via `DraftJsonRoot`. An empty Excel file under `Drafts/` would be
the same second-mechanism mistake the label editor's redirect already rejected. Stated here because
"create the board workbook" reads like "create a file" and a future reader would otherwise think it
was skipped.

**New UI**: `NewSystemWindow` (manufacturer `AutoCompleteBox` seeded from existing manufacturers so
a "Comodore" typo is discouraged, plus hardware/board/notes, a live preview of the identity, and
Create disabled until valid) and `SystemFilesWindow` (board image import by picker and drag-and-drop,
KiCad folder import, and the report). The entry point is a button on the **Contribute** tab, which
is unconditional - the Drafts tab is hidden until a draft exists, so a button living only there would
be unreachable on a fresh install - with a second copy on the Drafts tab header. Creation navigates
via the existing `_pendingBoardSelectionOverride` idiom, driving the ordinary
`OnBoardSelectionChanged` path, so "the same code path as editing an existing system" is true
structurally rather than by discipline.

**The KiCad report reproduces the render-time rule exactly**: `root.Pcb[0]`'s footprint references,
trimmed, `OrdinalIgnoreCase` - verified against `TabSchematics.KiCad.cs`'s
`BuildKiCadNormalizedNetNamesForReferences`, which is the only place a board label becomes copper.
Matching schematic symbols instead, or unioning every PCB file, would report matches that light up
nothing, which is worse than no report. `KiCadReferenceMatcher`/`KiCadReferenceMatchReport` are pure
and live in `CRT.Data`, taking plain strings rather than the KiCad types (which are `internal` to
`CRT.App`, and `CRT.Data` must never reference it) - the same "resolve the reads at the call site
and hand over plain values" split `LabelEditorSnapContext` follows. The report also surfaces two
other currently-silent failures for free: more than one PCB file (only the first is used) and zero
PCB files parsed, plus the generated CAD view names, which the Wiki previously told contributors to
read out of the logfile by hand.

**`AppConfig.KiCadDataFolderName`** replaces four hardcoded `"KiCad data"` literals, per CLAUDE.md's
"check AppConfig first" rule.

**Task 10 and task 11 are satisfied by construction** and need no separate work: nothing in the app
can create a `_UserContribution` sidecar any more (the only creation path is this one, and it writes
only `draft.json`), the existing sidecar READ path is untouched so old boards keep loading forever,
and a never-submitted draft is a fully working local system - now said plainly in both the UI ("A
system you never submit keeps working here as your own") and the Wiki.

**Known behaviour worth stating rather than fixing**: with "view boards as officially published"
ticked, a draft-only system renders as a blank board. That is a truthful answer - officially it does
not exist - and is documented in the Wiki's troubleshooting table rather than special-cased.

**What remains for a follow-up session** (do not re-open the design decisions above, they are
settled):
- **Session 2d** (drift warning) is the only part of Phase 2 still unstarted.

Verified: `dotnet test Classic-Repair-Toolbox.slnx -c Release` green at 3446/3446 (412 in
`CRT.Data.Tests` + 3034 in `CRT.App.Tests`), 0 new warnings. Task 8 took it to 3291; task 9 added
155 more - `NewSystemIdentityTests`, `NewSystemDraftWriterTests`, `KiCadReferenceMatcherTests`,
`NewSystemWindowTests`, `SystemFilesWindowTests`, plus extensions to `BoardDraftTests`,
`DraftDataStoreTests`, `BoardDataReaderTests`, `DraftManagerTests`, `DataManagerDraftOverlayTests`,
`TabDraftsTests` and `TabConstructionTests`.

Wiki pages touched across this whole session: `Board-JSON.md` (label editor + calibration save
sections, and its walkthrough forking into separate steps ending in a local draft save),
`Contribute-tab.md` and `Contribute-data-via-CRT.md` (first rewritten for the component editor's
redirect - no more "sent to the developer"/"your own copy does not change" - then extended again for
"Add a new system" and the private-systems point), and `Add-new-board-with-KiCad-data.md`, which task
9 rewrote almost entirely: the three manual steps it existed to describe (create the folders by hand,
build a board `.xlsx` by copying and emptying another, create and maintain a version-named
`_UserContribution` workbook) are gone, replaced by one "Add a new system" step; "read the CAD names
out of the logfile" is replaced by "the import lists them"; the `_UserContribution` troubleshooting
row is replaced by a short "boards made the old way still work" section, per task 10. The
`remind-wiki-mirror.sh` MAP gained the new `Tabs/Drafts/`, `Main.NewSystem` and `CRT.Data/NewSystem*`
paths, and was live-fire tested.

8. Point the existing editors at the draft root rather than at synced data: the label editor,
   `BoardComponentHighlightStorage`, component add/edit, file and link attachment.
   **This is where the real work is** - these editors currently assume they may write into the data
   root or nowhere. **[DONE 2026-09-20 - see session 2c above for all four pieces: label editor,
   KiCad calibration, and component add/edit/file/link attachment via reusing
   `ComponentContributionWindow`.]**
9. **"Add a new system"**: ask for manufacturer, hardware and board; create an empty draft; from
   that point everything is ordinary CRT. Deliberately the *same* code path as editing an existing
   system - there must not be two flows.
   **[DONE 2026-09-20 - see "Task 9" in session 2c above. Note the one item below that was answered
   differently than written: "create the board workbook ... from `BoardData`'s own schema" is
   satisfied by `draft.json`, whose sections ARE that schema. No `.xlsx` is created; writing one
   would be the second-mechanism mistake the label editor's redirect already rejected.]**

   **This replaces a manual procedure that already exists, and the replacement is mostly
   subtraction.** Today a contributor must: copy an existing board folder, empty the rows out of a
   copied board `.xlsx` by hand while keeping every sheet name exact, create a
   `Classic-Repair-Toolbox.v<version>_UserContribution.xlsx` whose name has to track the resolved
   main workbook's version, and register the board in it in the main workbook's format. The whole
   walkthrough is `Assets/Wiki/Add-new-board-with-KiCad-data.md`. Every one of those steps is a
   place to get it silently wrong, and the highest-value contribution is the one gated behind them.

   **What "Add a new system" must therefore do, concretely** - each item removes one manual step:
   - create the folder structure, so nothing is copied by hand;
   - create the board workbook with correct sheet names and headers from a schema the app already
     owns (`BoardData`), rather than by emptying a copy;
   - register the system so it appears in the dropdowns, with no version-named sidecar to maintain;
   - import schematic images by file picker or drag and drop;
   - accept a `KiCad data` folder and report what it found - which reference designators matched
     board labels and, more importantly, which did not. The Wiki already names a mismatch here as
     "the number one cause of 'I did everything and no traces appear'", and it currently fails with
     no error anywhere. Surfacing it is cheap and removes the worst failure in the whole process.

10. **`Drafts/` supersedes `_UserContribution` - settled, do not re-open this.** [ANSWERED
    2026-09-20, see open question 7. **DONE 2026-09-20** - satisfied by construction once task 9
    landed: the only path that creates a new system writes `draft.json` and nothing else, the
    sidecar READ path is untouched, and the Wiki's walkthrough now documents the old mechanism only
    as "boards made the old way still work".] `DataManager` already protects a locally-authored board and
    everything it references - board workbook, JSON sidecar, every referenced file, the whole
    `KiCad data/` folder - from being overwritten by sync or removed by orphan cleanup
    (`LoadProtectedContributionStateForCurrentData`). `Drafts/` provides the same guarantee and is
    now the ONLY mechanism new authoring writes through.

    Concretely: `LoadProtectedContributionStateForCurrentData` and the `_UserContribution` READ
    path stay exactly as they are, indefinitely - a board authored that way before this phase must
    keep loading. What stops is new writes: nothing in the app may create a new
    `_UserContribution` sidecar or a new `_UserContribution` main-Excel entry once `Drafts/` exists,
    and "Add a new system" (task 9 below) must never offer the sidecar as a choice. There is no
    migration of existing `_UserContribution` boards INTO `Drafts/` - they simply keep working
    where they are, read-only-mechanism-wise, for as long as anyone still has one.

    **Amended 2026-09-27 (owner request): a board folder the contributor puts into `Drafts/` BY
    HAND is taken in as a draft** at the next start (`DraftFolderImport`, run from
    `DataManager.LoadMainExcel`), so work done the old way can be submitted. Still no AUTOMATIC
    migration - nothing is moved out of `Data/` - but copying a `_UserContribution` board's folder
    into `Drafts/` is now the supported route to submit one. Such a board becomes a NEW system
    (never published), registered under its FOLDER names (a submission's identity needs them),
    while its `_UserContribution` entry stays in the lists and reads the draft
    (`MergeDraftOnlySystems` now also dedupes by folder). A new system's draft is compared against
    nothing (`DraftBoardSource.ComparisonBaselineOf`) - against the legacy copy in `Data/` every row
    read as unchanged and Submit was disabled.

11. Support **private systems**: a draft never submitted is a fully working local system. Say so in
    the UI; it is a genuine feature, not a staging area. Note this is already true of a
    `_UserContribution` board today - the feature exists, it is just hard to reach.
    **[DONE 2026-09-20 - said in the create dialog ("A system you never submit stays yours and keeps
    working locally"), on the Drafts tab, and in both Wiki pages. Nothing to implement: a draft-only
    system already loads, renders and is editable through the ordinary board path.]**

### Session 2d - Drift warning [DONE 2026-09-20]

12. When a draft's base revision is older than the synced system's, warn plainly and offer a view of
    what changed officially underneath. **Do not build a merge-conflict resolver** - a warning plus
    a diff is enough for a small personal change set.
    **[DONE 2026-09-20 - see "What was built" below.]**

**Groundwork already done (2026-09-20), before the session itself:**

- **A pre-existing `BaseRevision` bug is fixed.** `TabSchematics.KiCad.Calibration.cs` created a
  first draft as `new BoardDraft { SystemKey = cacheKey }` with NO `BaseRevision`, unlike the label
  editor and the component editor which both stamp `official.RevisionDate`. A system whose first
  edit happened to be a calibration therefore recorded no base revision at all - and since a blank
  value is ALSO the deliberate, correct value for a draft-only system (`NewSystemDraftWriter`, which
  has no official revision to be based on), the two were indistinguishable and the drift check would
  have silently skipped exactly those drafts. It now stamps `MainWindow.CurrentBoardData.RevisionDate`,
  which is correct because `BoardDraftApplier` carries `RevisionDate` through the merge unchanged
  from the official board, and that path has no official `BoardData` of its own in scope.
  **Not covered by a test**: this save path needs a live `MainWindow`, a loaded bitmap and an active
  calibration mode, so it is not headlessly reachable. The three pure writers do already pin "carry
  `BaseRevision` forward unchanged" in `CRT.Data.Tests`; app-side first-save STAMPING is untested at
  all three call sites.
- **`BaseRevision` has no reader anywhere today.** Session 2d is its first consumer.
- **What `RevisionDate` actually contains**, read out of the shipped workbooks in `Assets/Data`
  rather than assumed: `2026-May-14`, `2026-August-21`, `2026-July-19`, `2026-August-23`. That is
  `yyyy-MMMM-dd` with the FULL ENGLISH MONTH NAME, not ISO-8601 - while the test fixture
  (`BoardWorkbookBuilder.WriteCompleteBoard`) defaults to `2026-01-15`, so the suite's shape and the
  shipped data's shape genuinely differ and both need covering. Two consequences: ordering IS
  possible but ONLY via explicit `DateTime` parsing with `CultureInfo.InvariantCulture` and an
  explicit format list (plain string comparison is actively wrong - lexically `2026-August-21` sorts
  before `2026-May-14`), and it must fail soft, because the value is free text scanned off a
  `# Revision date:` marker and a hand-edited board can carry anything. Where either side is
  unparseable the honest claim is "changed", not "older", and the UI wording must match whichever
  was actually established.

**What was built (2026-09-20).**

**The comparison is three-valued, not a boolean** - `DraftRevisionComparer` returns `InSync`,
`OfficialIsNewer`, `Changed` or `Unknown`, and every user-visible string is tied to which came back.
"Updated" is only ever said when both sides parsed as dates and the official one is genuinely later;
everything else that differs says "changed". `Unknown` (a blank base revision, or a draft-only
system) never warns: an unrecorded base is not evidence of drift, and stamping today's revision to
silence it would assert something false. There is deliberately no `DraftIsNewer` - an official
revision that parses OLDER than the base is a rollback or a hand-edit, still "changed underneath
you", and a fourth message nobody could act on differently would be noise.

**A genuine official-versus-official diff is IMPOSSIBLE, and no snapshot was added to make it
possible.** A draft records its base revision as a string and keeps no copy of the BoardData it was
based on; sync overwrites `Data/` in place, so the old official workbook is gone. Storing a per-row
snapshot or hash in `draft.json` was evaluated and rejected on four grounds, the first decisive:
**Phase 4's diff is server-side** ("the server compares two states it fully knows"), so a
client-side base snapshot is dead weight Phase 4 discards; it would only ever help drafts created
after it shipped, which is exactly the wrong set, and there is no migration mechanism (deliberately)
to fix that; it doubles the size of a file the doc itself calls a "small personal change set"; and
per-row identity living in the data is what "Retiring the UUIDs" moved away from.

**What IS answerable, and is what got built, is a per-row collision report** - how each drafted row
stands against the official data as it is NOW (`DraftDriftDetector`, `DraftDriftReport`). Four
standings, of which two carry a consequence the contributor has no other way to discover:
`GoneOfficially` (the row you changed no longer exists officially, so `BoardDraftApplier` silently
drops your edit) and `NowExistsOfficially` (something you added has since been added officially
too). **Both are proven through the REAL merge** in `DraftDriftReportTests`, not just against the
detector's own output.

**That merge proof corrected the plan's own proposed wording**, which is the main thing worth
carrying forward: the design said to tell the user "Your version is what you see" for
`NowExistsOfficially`. The test showed the opposite - `BoardDraftApplier` FAILS CLOSED on an `Added`
row colliding with an official one and keeps the **official** row. The shipped string says "The
official version is the one shown; yours is not applied while both exist", and both the enum's own
comment and a UI test now pin it.

**`BoardDraftNaturalKeys.ForRow`** - `BoardDraftApplier.KeyOf`'s section-dispatch switch moved into
the natural-keys class (whose header already called itself "the one place a BoardData row's identity
is decided") and `KeyOf` now delegates to it. This was the only edit to existing merge code, done
first and verified against the untouched `BoardDraftApplierTests`. `BoardDraftNaturalKeys.Separator`
became public at the same time, because a `Deleted` tombstone carries no `Entry` and its key's parts
are the only readable thing left - it renders as a box, so it is split apart for display and never
shown raw.

**Detection runs at three points, none of which bulk-loads a board.** (a) The selected board, for
free: `DataManager` now publishes `LastLoadedDraftBaseRevision`/`LastLoadedDraftIsNewSystem` as
siblings of `LastLoadedDraftSummary`, set from the SAME draft object for the same documented reason.
Note these are deliberately NOT suppressed by `ViewOfficialPublishedOnly` - that toggle hides edits,
it does not make drift untrue - so the suppression happens at the banner instead. (b) After a sync
that changed nothing, since a sync that DID change files owns the banner with its own more-urgent
"please refresh board"; no board reload is ever forced, because doing IO behind the user's back to
raise a warning is worse than the warning arriving on their next real load. (c) The Drafts tab, via
a new `BoardDataReader.ReadRevisionDateOnly` that reads the marker without mapping ten sheets, plus
`TryGetCachedRevisionDate` so an already-loaded board costs nothing. Both share one private
`ScanRevisionDate` with `LoadAsync`, pinned equal by a test - if those two ever disagreed, the check
would compare a value against itself read a different way.

**Two surfaces, both reusing session 2b's vocabulary rather than inventing a parallel one**: an
amber `Draft_Chip_*` "Updated officially" chip (removed 2026-09-28, owner request - it only repeated
the drift line under the row) plus a "What changed" button on the Drafts tab row
(the primary home - it is the one surface that is *about* drafts), and the existing `SyncBanner` for
the board on screen, following `_isShowingDataSyncDisabledBanner`'s precedent for shared ownership
of that one text slot. A drifted-row schematic tint was rejected: the drafted tint already means
"this row is yours", and a second meaning on the same visual channel, on a hot render path, is worse
than a report window that answers the same question.

**"I have looked at this" (`DraftBaseRevision.Rebase`) writes ONE STRING.** It is reachable only
from inside the report window, never from the tab row or the banner - a one-click dismiss from a
list would let someone silence the warning without reading it, which is the failure mode the feature
exists to prevent. It changes no drafted row, merges nothing, resolves no conflict, and does not
drop a `GoneOfficially` row. No confirmation dialog: nothing is destroyed, and the next sync that
moves the revision raises it again. Pinned by a test that re-reads `draft.json` from disk.

**The legacy blank-BaseRevision case needs no migration.** `DraftBaseRevision.EnsureBaseRevision`
stamps a draft that has none on its next ordinary save through any of the three editors, and never
overwrites one that is already recorded - the same "a floor, not a migration" house style the
worklog id counters use. A draft never edited again simply stays silent, which is correct for a base
nobody can reconstruct. `NewSystem != null` is the unambiguous discriminator against a draft-only
system's deliberately blank value.

**Also verified this session, and recorded under "Definition of done" above:** Phase 2's own first
criterion, "a sync never destroys a draft (test this explicitly, with a real sync)", had NO test
until now.

### Definition of done

- A sync never destroys a draft (test this explicitly, with a real sync).
  **[VERIFIED 2026-09-20]** - `DataManagerOrphanCleanupTests.The_cleanup_never_touches_a_draft_or_its_attachments`
  and `A_draft_still_reads_back_intact_after_a_cleanup`. Orphan cleanup is the only code path in the
  app that deletes data files, and the safety is STRUCTURAL rather than defensive: it enumerates
  only inside the data root and re-checks containment per file, while `Drafts/` is a SIBLING of
  `Data/` in AppData. No guard in the cleanup mentions drafts at all, which is precisely why the
  test matters - the tests were confirmed to FAIL (with "the cleanup deleted a draft.json - a
  contributor's unpublished work, which nothing can restore") when drafts are moved inside the data
  root, and they carry an anti-vacuity probe asserting the cleanup really does delete an
  unreferenced file in the same run.
- A drafted board renders with markers; the toggle shows the official version.
  **[VERIFIED - session 2b]** `BoardDraftSummary` plus the component-list chip and schematic tint;
  `UserSettings.ViewOfficialPublishedOnly` suppresses the overlay at the one place it is applied.
- A complete new system can be authored with no server involved.
  **[VERIFIED - session 2c task 9]** Create, select, label, add components, import board images and
  KiCad data, all offline and all through the ordinary board path.
- The overlay merge is pure, in `CRT.Data`, and unit tested across all sections.
  **[VERIFIED - session 2a]** `BoardDraftApplier` + `BoardDraftApplierTests`.

### Traps

- **The overlay must not leak into contributed output.** When Phase 4 submits, it sends
  official+draft as the new intended state - correct - but nothing in the app may ever write draft
  content into `Data/`.
- **Wiki pages need updating in the same commits**: `Contribute-tab.md`,
  `Contribute-data-via-CRT.md`, `Commandline-parameters.md`, `Explanation-of-data-files.md`. The
  `remind-wiki-mirror.sh` hook will name some of these; read its output rather than guessing.
- **Do not delete the existing Contribute tab yet.** It stays until Phase 7.

---

## Phase 3 - Server service and accounts [DONE 2026-09-21]

**Goal.** An ASP.NET Core service running on the AlmaLinux box, reachable at `/api/`, with accounts
in MariaDB. No submission handling yet.

**Needs the access described in [Access the project owner must provide](#access-the-project-owner-must-provide).**

### Progress

**Step 0 is complete (2026-09-20) and is waiting on the project owner to deploy.** Built:
`src/CRT.Server/` (ASP.NET Core `net10.0`, referencing `CRT.Data`) answering `GET /api/health` and
nothing else, `tests/CRT.Server.Tests/` (13 tests), and `src/CRT.Server/DEPLOYMENT.md` written as a
runbook the project owner executes themselves, every step carrying a Verify command. Both projects are
in the `.slnx`. Full suite green at 3555, Release build 0 warnings / 0 errors.

**Decisions taken during step 0, none of which should be re-opened:**

- **Production-write interlock: denied at the FILESYSTEM**, not merely in code. The project owner
  confirmed this approach. `crt-server` has no write permission on the Production tree, so a write
  attempt is refused by the kernel; `DEPLOYMENT.md` step 3 carries a probe whose output the
  project owner checks. `systemd`'s `ProtectSystem=strict` + `ReadWritePaths` is a second, independent
  interlock. This is what makes the rule secure-by-DESIGN (structurally unreachable) rather than
  secure-by-default, satisfying [checklist item 8](#review-checklist-for-each-of-phases-3-6)
  properly rather than with a recorded exception.
- **Tokens will be opaque and database-backed, not stateless JWTs.** Forced by
  [Phase 6](#phase-6---maintainers)'s definition of done - "Removing a maintainer takes effect
  immediately, proven by a test using an already-issued token" - which no stateless token can
  satisfy. The cost is one indexed single-row `SELECT` per authenticated request, which is nothing
  at this audience size. Task 6 left this open; it is now decided.
- **Account enumeration: registration and password reset both answer a neutral `202`.** A new
  address gets a verification link; an already-registered one gets a "someone tried to register
  with your address" mail carrying a reset link, moving the disclosure into the channel that
  already proves ownership. Task 5 did not mention enumeration at all.
- **`0001_initial.sql` will create SEVEN tables, not the five task 4 lists.** `sessions` and
  `account_tokens` are needed by tasks 6 and 5 respectively, and adding them later would mean a
  second migration for tables that belong in the first.
- **EPPlus licensing (open question 8) does NOT block Phase 3.** Verified: `EpplusLicense.Ensure()`
  is called only from `BoardDataReader`'s EPPlus entry points, never from a static constructor or
  module initializer, so a service that reads no workbook executes no EPPlus code. The assembly
  ships as a transitive reference; nothing runs. The project owner has asked to keep using EPPlus and
  revisit later, possibly replacing it. **It becomes a real question at Phase 4 task 3**, where the
  server-side diff does read workbooks.

**Also verified during step 0, and worth not re-deriving:**

- **The release pipeline is unaffected by the new projects.**
  [build-and-release.yml](../.github/workflows/build-and-release.yml) publishes
  `src/CRT.App/CRT.App.csproj` **by name**, never the solution, so a Web SDK project in the `.slnx`
  cannot be dragged into a per-RID self-contained publish. Confirmed by running the win-x64
  self-contained publish after adding both projects. (Solution-wide `build` and `test` do happen -
  in both workflows and the Stop hook - which is why `CRT.Server.Tests` must stay fast. It runs in
  ~200 ms.)
- **`CRT.Data` is genuinely Linux-safe.** No culture-sensitive casing or date parsing anywhere in
  it, and `NewSystemIdentity` already applies Windows' stricter rules unconditionally. Separately,
  the shipped data tree was swept for case mismatches - 4422 file references across every shipped
  workbook, **0** that match only case-insensitively and **0** paths colliding by case - so the
  exact-case `File.Exists` calls in `CRT.Data` are safe on a case-sensitive filesystem today. See
  the trap note below; this needs to become a Phase 4 validation rule, because nothing enforces it.
- **A leak found by running the service, not by a test.** The .NET SDK appends `SourceRevisionId`
  to `InformationalVersion`, so the first live `GET /api/health` returned the full git commit SHA
  from an unauthenticated, publicly reachable endpoint. `HealthReport` now strips build metadata,
  pinned by four tests. Worth remembering for any future endpoint that reports a version.

**Step 1 (the project owner's) is COMPLETE - the chain is proven, 2026-09-20.**
`curl -i https://classic-repair-toolbox.dk/api/health` answers `200 OK` with the expected JSON from
a separate machine over the public internet. Also verified: the service restarts within ~5 s after
`SIGKILL`; it listens on `127.0.0.1:5199` only; port 5199 is not reachable from outside; the
service user can write BETA and is refused by the kernel on Production; and the deployed binaries
are read-only to the service user. Running on AlmaLinux with
`aspnetcore-runtime-10.0` (10.0.12), behind the existing Apache vhost.

**Four defects in `DEPLOYMENT.md` were found by actually running it, all fixed in that file with
the reasoning recorded.** They are listed here because each cost a debugging round and would
otherwise be rediscovered:

1. **Group mismatch.** Step 4 said `chown root:crt-server` while the unit runs `Group=crt-data`, so
   the process fell through to the "other" bits of a `0750` directory and could not even enter it.
   systemd reports this as the unhelpful `status=200/CHDIR`.
2. **`Type=notify`.** A plain ASP.NET Core host sends no readiness notification (that needs
   `Microsoft.Extensions.Hosting.Systemd` and `UseSystemd()`), so `systemctl start` hung until
   timeout and reported failure **while the service was serving requests normally**. Now
   `Type=simple`.
3. **`ProxyPass` inside `<If>`.** Apache refuses it outright (`AH00526: ProxyPass cannot occur
   within <If> section`) - it is resolved at configuration time, while `<If>` is per request. The
   HTTPS-only requirement is expressed by guarding the header instead:
   `RequestHeader set X-Forwarded-Proto "https" "expr=%{HTTPS} == 'on'"`.
4. **An unconditional `X-Forwarded-Proto: https`** on a combined `<VirtualHost *:80 *:443>` tells
   the service that a plain HTTP request arrived over HTTPS. Harmless for health, but the accounts
   step builds verification and password-reset links from that header.

Two further notes worth carrying forward: this box has SELinux **disabled**, so the
`httpd_can_network_connect` boolean is not needed (the doc now checks `getenforce` first rather
than assuming AlmaLinux's default); and **the public URL must not be tested from the server
itself** - it hairpins out to the router and back and can simply hang, which looks exactly like a
broken deployment. The doc now tests Apache with
`curl --resolve <domain>:443:127.0.0.1` and leaves the public path to another machine.

**This is the argument for the health-endpoint-first ordering, concretely: four infrastructure
defects surfaced with no application code in the picture to blame.**

**Step 2 is COMPLETE (2026-09-20), awaiting deployment.** Built `ServerOptions` (the POCO),
`ServerOptionsValidator` (pure, 33 tests), startup validation in `Program.cs` that **throws before
the host listens**, `ServerLogAdapter` installing `CrtLog.Sink`, and `appsettings.Example.json`.
Suite green at 3588.

**The no-default rule is implemented and demonstrated, not merely asserted.** Run with no
configuration, the service logs six individually-named errors and refuses to start; given a valid
configuration it starts and serves health. Both halves were verified by running the real binary.
Six values have no default and no fallback: `DataTreeRoot`, `ProductionTreeRoot`,
`PublicDataBaseUrl`, `ManifestPath`, `ConnectionString`, `PublicApiBaseUrl`. Values that merely
tune behaviour (token lifetimes, Argon2 cost, SMTP host) do default, because a missing one there
has a safe answer - the dividing line is "could getting this wrong write data somewhere it should
not, or weaken a control?".

**`PublicDataBaseUrl` and `ManifestPath` are validated in Phase 3 although nothing reads them until
Phase 5.** The data root, the manifest path and the public base URL are three independent values
with nothing relating them, so pairing the BETA tree with the Production base URL produces a
perfectly valid, correctly sorted, correctly hashed manifest in which every `url` points at the
wrong tree - clients then sync Production content believing it is BETA, or 404 on everything. It is
invisible in every way. The marker check (`-BETA` must appear in every path and in the URL) costs
two lines in a class being written anyway and means Phase 5 cannot get this wrong by omission.

**A real defect found by its own test:** the "must be an absolute path" check could never fire,
because `Path.GetFullPath` resolves a relative path against the working directory *before* the
check ran - so a relative `DataTreeRoot` would have silently resolved against the service's own
working directory.
Both that check and the identical one on `ManifestPath` now test the raw value. This is the third
time in this project that writing the test first has caught something the code got wrong.

**Deployment layout changed before step 3 was run (2026-09-20), at the owner's request.**
Everything the service owns now sits under ONE root beside the site it serves,
`<site>/crt-server/`, holding `app/` (binaries plus `appsettings.Production.json`) and
`previous/` (the rollback copy) - instead of being split across `/opt/crt-server` and
`/etc/crt-server`. The reason given was that scattering one service across several standard
locations is harder to hold in your head than one folder; the reason it is also better here is
that the configuration file was always going to end up beside the binaries anyway, since that is
the only place the ASP.NET Core host reads it from without extra wiring, so the `/etc` half was
a location that existed but was not used.

Two constraints came out of it, both now enforced in `DEPLOYMENT.md`:

- **`crt-server/` must be a SIBLING of `public_html`, never inside it.** Inside, Apache would serve
  `appsettings.Production.json` and its database password to anyone who asked. Step 4 and step 9
  each carry a `curl` that must return 404, the second of them placed deliberately BEFORE the
  password is typed in.
- **`/root/scripts/` was considered and rejected.** Reaching it requires `chmod 0711 /root`, since
  the kernel must traverse `/root` itself whatever the deeper modes say - a real, if narrow,
  weakening of `/root` that a future reader would find unexplained. The site directory needs no
  such change because Apache already traverses it.

Source comments and the example config no longer name a deployment path at all; they say "the
service's own directory" and point at `DEPLOYMENT.md`, so a future move touches the runbook only.

**Step 3 is DONE (2026-09-21).** The service runs from `<site>/crt-server/app/`, health answers,
and the negative test was performed: removing `DataTreeRoot` alone produced exactly one failure
naming that one setting, rather than a generic refusal. The `-BETA` marker, both path checks and
the write probe all passed against the real tree.

**One trap cost an hour and is now written into DEPLOYMENT.md.** The runbook said to own the config
file `root:crt-server`, which is unreadable to the service: **systemd does not apply a user's
supplementary groups when the unit sets `Group=` explicitly**, so the process runs
`crt-server:crt-data` and matches neither owner nor group on a `root:crt-server` file. It must be
`root:crt-data`. The obvious verification is also wrong - `sudo -u crt-server cat` DOES grant
supplementary groups, so the file reads fine by hand while the service cannot open it. Check with
`systemd-run --uid=... --gid=...` instead. The failure appears as `IOException: Permission denied`
inside `WebApplication.CreateBuilder`, before any of our own validation runs.

**Step 4 is DONE (2026-09-21).** Plain numbered SQL migrations, no ORM - chosen by the project owner
over EF Core so the schema stays readable as SQL in the repository and no dependency is added.

- **`MigrationPlan`** is pure and carries every rule: filename ordering, gap detection, checksum
  comparison, and the merge-collision case. 22 tests, no database.
- **`MigrationRunner`** is the thin I/O rim - reads the files, asks the plan, executes.
- **Migrations run at STARTUP**, not from a separate command: a separate step is a step that gets
  forgotten, and the failure mode of forgetting is a service running against a schema it does not
  match. Same bargain `ServerOptions` makes.
- **Four refusals, each fatal**: an applied migration whose content changed, a gap in the
  numbering, an unapplied migration numbered below one already applied, and a recorded migration
  whose file is gone. All four are silent-divergence bugs if allowed through.
- **`MySqlConnector` 2.6.2** (MIT), not Oracle's `MySql.Data`.
- **MariaDB commits DDL implicitly**, so the per-migration transaction protects the
  `schema_migrations` record, NOT the schema change. Documented in both the code and the runbook.

Seven tables as decided, plus `schema_migrations`. Three design points worth not re-deriving:
`accounts` is unique on a normalised email (so one person cannot register twice by changing case),
tokens and sessions store **only hashes** (a database read must not yield working reset links), and
`audit` rows are append-only with a text copy of the actor's identity, since the account may later
be deleted.

**Step 5 is DONE (2026-09-21).** Hashing, the mail seam, templates and the rate-limit policy, all
pure except the one SMTP class. 87 new tests.

- **Argon2id via Konscious 1.3.1**, stored as a standard PHC string
  (`$argon2id$v=19$m=65536,t=3,p=2$...`). The parameters travel WITH the hash, which is what makes
  raising the cost safe: verification reads them from the stored string, never from configuration,
  so old hashes keep verifying and `NeedsRehash` lets the login path upgrade them silently. A
  format that did not carry its parameters would make every cost increase a mass password reset.
- **`PasswordHashEncoding` is split from the hasher** because parsing is where the subtle bugs
  live. Base64 padding restoration is tested at every length remainder - an off-by-one there
  corrupts the salt and makes affected passwords permanently unverifiable.
- **A malformed stored hash fails the login rather than throwing.** A 500 would tell an attacker
  that this particular account is interesting.
- **The rate limiter must sit IN FRONT of the hasher.** Argon2 allocates ~128 MiB per concurrent
  verification at the configured cost, so unlimited attempts are a memory-exhaustion vector
  independent of whether any password is ever guessed. Two independent buckets: per account (stops
  a password list against one address) and per IP (stops spraying one password across many
  accounts, which no per-account limit can see).
- **The lockout is a DELAY, not a disable**, and `RetryAfter` is measured from the OLDEST attempt
  in the window. Measuring from the newest means every further try extends the lockout, so someone
  tapping "try again" can never get back in.
- **Anti-enumeration lives in `EmailTemplates`, not in the endpoints.** Registration answers a
  neutral 202 either way; what differs is which mail is sent, visible only to whoever controls the
  mailbox. A known address gets an "already registered" mail carrying a reset link - useful to the
  person who forgot, and a security notification if it was not them. An unknown address on the
  reset path gets NO mail, or the endpoint becomes a way to mail any address at all.
- **The already-registered mail never uses the display name supplied by whoever is registering** -
  that mail goes to somebody else's inbox, so using the attacker's text is an injection.
- **`PasswordChanged` contains no link at all**: a notification that asks for no action cannot be
  phished.
- **Password policy is length-led** (12 characters), with no composition rules - withdrawn NIST
  guidance that reliably produces weaker passwords. Pinned by a test, so adding one is a
  conversation.
- **Email normalisation is invariant-culture lowercase**, tested under `tr-TR`: a Turkish locale
  lowercases 'I' differently, and this value is the unique index in `accounts`.
- **`MailFromAddress` became REQUIRED** rather than validated-if-present. It was previously only
  checked when present, so a missing one passed startup and then threw inside a background send
  when somebody registered - a log line nobody is watching instead of a refusal to start.

**Step 6 is DONE (2026-09-21).** Registration, verification, login, refresh, logout,
`GET /api/accounts/me` and password reset. 91 new tests, 246 in `CRT.Server.Tests` overall.

- **`IAccountStore` is the database seam**, the same idea as `IMiniproRunner`. `AccountFlows` holds
  every decision and takes the store as an argument, so all of it is tested against
  `FakeAccountStore` with no database. `MySqlAccountStore` behind it is an untested I/O boundary
  with no decisions in it at all.
- **The status codes ARE the security design.** Registration and forgot-password always answer
  202, whether or not the address exists - a 409 for a taken address is an account enumeration
  oracle. Login answers 401 for every failure with no reason attached, so unknown address, wrong
  password and locked account are indistinguishable. Token redemption DOES distinguish its
  failures, because the token is the only subject.
- **The rate limiter runs before the hasher**, pinned by a test that asserts a rate-limited
  request records no further failure - it never reached the code that would.
- **Rotation with reuse detection**: presenting an already-rotated refresh token revokes the whole
  chain, including the legitimate client's session. An inconvenient re-login beats leaving a thief
  with a live session, and the thief is told nothing about having triggered it.
- **Locking an account takes effect on the very next request** - Phase 6's definition of done in
  miniature, and the reason these are opaque database-backed tokens rather than JWTs.
- **Completing a password reset** consumes every other outstanding reset token, revokes every
  session, and verifies the address (clicking a link in that mailbox proves what verification was
  asking for).
- **Verification and reset tokens cannot substitute for each other**, tested in both directions -
  a verification link acting as a password reset would be account takeover from a link sent to
  every new registration.
- **`ClearAuthFailuresAsync` renames rather than deletes.** The audit trail is append-only, and
  deleting the record that failed attempts happened is precisely what an attacker would want.

**That observation was RESOLVED, and it was not a flake.** The failing test was
`NewSystemWindowTests.A_name_with_a_path_breaking_character_is_rejected_and_says_so` with
`manufacturer: "Commo*dore"`, and the cause was a real platform-dependent bug rather than shared
static state: `NewSystemIdentity.IsValidPathSegment` built its invalid-character set from
`Path.GetInvalidFileNameChars()`, which returns the full reserved set on Windows and only `/` and
NUL on Linux. So `*`, `?`, `"`, `<`, `>`, `|` and `:` were refused only where the host happened to
refuse them. A contributor on Linux could have created a system named `Commo*dore`, and the
failure would then have appeared for every Windows user who synced it - on a machine that never
saw the name being typed. `WorkbookExportModel.SanitizeForFileName` already documented this exact
trap for exported file names; `NewSystemIdentity` had patched only the single `\` character rather
than applying the lesson.

The fix writes the union across every platform out explicitly (with `GetInvalidFileNameChars`
still unioned in as a backstop) and adds
`Every_character_ANY_platform_reserves_is_refused_on_EVERY_platform`, which is what makes the rule
meaningful on the Linux CI runner. Two further tests pin the deliberate behaviour underneath it:
`SanitizePathSegment` normalises whitespace control characters to single spaces BEFORE the
character check, so a name pasted out of a spreadsheet arrives as `"C64 NTSC"` rather than being
refused for something invisible, while a non-whitespace control character is refused.

**The lesson for this document: calling an unexplained failure a flake is a decision, not an
observation.** It was recorded here as one without the name being captured, and it was a latent
data-corruption bug the whole time. Capture the name first, every time.

**Step 7 is DONE (2026-09-21), against the live service on the box.** Every item in the definition
of done is now verified in production, not only in tests:

- **Register -> mail -> verify -> login -> `/me`** all work. `is_verified` flipped for the verified
  account and stayed `0` for the other, so a token affects exactly the account it belongs to.
- **The enumeration guarantee holds in production.** A wrong password on a real account and a login
  for an address with no account both answered `401` with no distinguishing body.
- **Rotation and reuse detection work.** A refresh issued a new token; presenting the retired one
  answered `401` AND revoked all three sessions with `refresh token reused`, with a
  `session.reuse_detected` audit row naming the source address. The successful refresh's own new
  session was revoked too - the intended trade when a theft is suspected.
- **Sessions survive a restart**, because they live in the database rather than in memory. Proven
  incidentally: a token issued before a reboot still refreshed afterwards.
- **Reboot survival and `kill -9` recovery** both confirmed; the unit came back `active` on its own.

**Two pre-existing box faults were found and fixed during this step, neither caused by Phase 3:**

1. **Postfix could not start `smtpd` at all.** `/etc/postfix/main.cf` had two settings collapsed
   onto one line (`alias_maps = lmdb:/etc/aliases alias_database = lmdb:/etc/aliases`), so postfix
   died with `open dictionary: expecting "type:name" form` on every connection and master throttled
   the restarts. Anything on that box needing mail had been failing silently. Note for AlmaLinux 10:
   the map type is `lmdb`, and `hash` is not available.
2. **The service behaved correctly throughout that outage** - it logged a warning and the
   registration still succeeded, which is the documented contract (a mail failure must not fail the
   operation it accompanies, or the user has an account they cannot reach and no way to retry).

**Deferred, not failed** - both are exercised by unit tests, just not yet against the live service:
the rate limiter (11.7) and the password-reset round trip (11 step 3 onward). Neither blocks
Phase 4.

**Also noticed and left for the project owner:** postfix listens on `0.0.0.0:25` rather than loopback,
so port 25 should be confirmed closed at the firewall - an open relay is how a box lands on a
blocklist. Unrelated to this project.

**Phase 3 is closed. Phase 4 (the submission pipeline) is next**, and is a substantially larger
piece of work than any phase so far.

### Tasks

1. Create `src/CRT.Server/`, ASP.NET Core on `net10.0`, referencing `CRT.Data`.
2. **Health endpoint first** (`GET /api/health`). Deploy that alone and prove the whole chain -
   systemd, reverse proxy, TLS - before writing any business logic.
3. Server setup, documented in `src/CRT.Server/DEPLOYMENT.md` as it is done:
   - `dnf install dotnet-runtime-10.0` (Microsoft ships RHEL/AlmaLinux packages; AlmaLinux tracks
     RHEL, so the RHEL feed is correct);
   - a dedicated unprivileged service user;
   - a `systemd` unit with `Restart=always`, listening on `127.0.0.1:<port>` only - **never** on a
     public interface;
   - reverse proxy from `/api/` in the existing Apache/nginx site;
   - config at `appsettings.Production.json` in the service directory, mode `0640`, root-owned.
4. MariaDB schema. Start deliberately small:
   - `accounts` (id, email, password hash, display name, verified, created)
   - `systems` (system_id, manufacturer, hardware, board, current_revision, origin)
   - `maintainers` (system_id, account_id, granted_by, granted_at)
   - `submissions` (id, system_id, account_id, base_revision, state, created, decided, decided_by)
   - `audit` (id, actor, action, subject, detail, at)

   Use an explicit migration mechanism from the very first table (a numbered SQL script runner is
   enough - no ORM migrations framework is required). Never hand-edit the live schema.
5. Accounts: email + password (**hash with Argon2id or bcrypt - never a bare SHA**), email
   verification, password reset. No federated login: some CRT users specifically avoid GitHub.
6. Authentication for the desktop clients: a bearer token with a sensible lifetime and refresh.
7. **Point everything at BETA.** The service must read and write only the BETA tree until the
   project owner says otherwise. Make the target tree a configuration value with **no default**, so a
   misconfigured service fails to start rather than quietly writing to production.

### Definition of done

- `GET /api/health` answers over HTTPS from the public domain.
- The service restarts automatically after a reboot and after a kill.
- An account can register, verify, log in and refresh from a test client.
- Schema is created by a committed migration script.
- No credential is in the repository.

### Traps

- **Do not let the service touch Production.** See item 7.
- **File ownership.** The service user must be able to write the data tree without breaking what the
  web server serves. Agree group ownership with the project owner rather than loosening permissions.
- **`dotnet publish -r linux-x64` from Windows works fine**, but confirm no Windows-only path
  assumption leaked into `CRT.Data`.

---

## Phase 4 - Submission pipeline [DONE 2026-09-21]

**All seven tasks are DONE.** Tasks 1-6 landed first; task 7's writer was the last piece and was
deliberately left blocked until publishing existed to call it. It now does - `PublishExecutor`
(Phase 5 task 6) writes `system.json` at publish time, which is exactly where task 7's own note
said the writer belonged.

**The suite was at 4148 when Phase 4's client and server work closed** (775 CRT.Data + 279
CRT.Server + 3094 CRT.App, Release); it is at 4240 with Phase 5's publishing step in.

**Migration 0004 is REQUIRED before the server can take a single real submission.** Three
schema-versus-code disagreements were found by reading the schema, not by any test - see task 7
below. It applies automatically on the next service start.

- **`SubmissionContract.cs` in CRT.Data** is the one definition both sides reference, which is the
  defect the old `ComponentContributionPayload` had: the app and the PHP each carried their own
  idea of the shape, kept in step by hand. Format version 1, deliberately NOT continuing the old
  payload's numbering - that contract described one component, this describes a whole system, and
  a shared sequence would let an old client's "2" be mistaken for something this understands.
- **No `UuidV4` anywhere.** Rows pair on `BoardDraftNaturalKeys`, which Phase 2 already built for
  exactly this. `Renames` is carried explicitly because natural keys cannot see a rename.
- **`SubmissionPathRules` resolves and then checks containment**, rather than scanning for `..`.
  Pattern-matching a path is a losing game; resolving first is the one formulation that cannot be
  worked around. Tested against a list of known tricks (`Images/./../../evil.png`, `//server/share`,
  a NUL byte, reserved Windows device names, trailing-dot stripping) because that is how this bug
  class ships - the obvious case is caught and a variant is not.
- **Case is preserved and compared case-SENSITIVELY throughout**, and a case-only collision between
  two paths is an error: both can exist on the Linux server, only one on a Windows client, so the
  tree would be un-syncable for much of the audience. A reference that differs only in case gets a
  message naming BOTH spellings, because "file not found" for a file you can plainly see is
  baffling.
- **Blobs are content-addressed** and the hash is VERIFIED before a blob enters the store. Without
  that, "content-addressed" is a lie the client can tell - uploading arbitrary bytes under the hash
  of an innocent file would let a submission reference something the server believes it has vetted.
- **Uploads resume because the partial file's LENGTH is the offset** - no separate bookkeeping to
  get out of step with the bytes on disk. A wrong offset is refused with where to actually resume,
  rather than appended and caught by the hash check after the whole upload completes.
- **Ownership is checked in the flows, not the endpoints.** A submission id is a small integer, so
  without it anyone could upload into anyone else's submission. "Not found" and "not yours" are
  deliberately indistinguishable (404 both), or the id space could be walked.
- **An unverified account may log in but may not submit** - that line is drawn in
  `SubmissionFlows.CreateAsync`.
- **Validation runs BEFORE a submission row exists**, so a broken client cannot fill the table with
  rows that will never be finalised, and the contributor learns before uploading for an hour. It
  runs AGAIN at finalise, because at creation the files were only declared, not present.
- **Highlight coordinates are parsed invariant-culture only.** "1,5" under a comma-decimal culture
  is 1.5 and under an invariant one is 15 - a highlight ten times too wide, with nothing failing.
  A comma is reported as malformed rather than reinterpreted.
- **`BlobStoreRoot` is a new no-default setting.** A Phase 4 build over a Phase 3 configuration
  refuses to start naming it. It must sit outside the document root and needs adding to the unit's
  `ReadWritePaths`; DEPLOYMENT.md step 9 covers both.
- **0002_submission_files.sql** adds `submission_files`, `submission_payloads` and
  `submission_findings`, plus the `uploading`/`abandoned` states. Rows are stored as JSON in one
  column rather than eleven tables - they are never queried field by field, and eleven tables would
  need a migration every time `BoardData` changed.

### CONTRIBUTING NEEDS NO ACCOUNT (owner's decision, 2026-09-21)

**This reverses what the server was first built to do, and it changes Phase 6's entry point.**

The agent implemented submission as requiring a verified account, because the roles table lists a
"Contributor" role and Phase 3 task 5 says "Accounts: email + password". Neither of those actually
argues that a CONTRIBUTOR needs one, and the threat model says the opposite outright: *"simply
submitting hostile content is the cheaper route, and it needs no account theft at all."* Accounts
never defended against threat 1; content validation and human review do.

The owner's decision, which replaces the implemented behaviour:

- **Contributing anything requires NO account and must not even ask for one.** A sign-up wall
  before a hobbyist can fix a typo is how a contribution does not happen. This matches the old PHP
  path, which asks for an address in a form field and nothing more.
- **An email address IS required, for contact only.** It is how the contributor learns their work
  was accepted or rejected. It is not a credential, there is no password, and nothing is stored to
  sign in with.
- **A NEW SYSTEM is where an account appears - and the PROJECT OWNER creates it, the contributor does
  not sign up for it.** When a new system is accepted, the project owner creates a maintainer account
  for that person (if they do not already have one) and assigns them to the system. They then
  maintain their own system from then on.

**Why this is the better model, beyond the lower barrier:** it means an account always represents
*authority over something*, never merely *having contributed once*. A database of accounts that
exist only because somebody once fixed a typo is a liability - credentials to keep safe, addresses
to protect - in exchange for nothing. Under this model the only accounts that exist are ones
somebody deliberately granted.

**What this means for the code (to be corrected):**

- `SubmissionFlows.CreateAsync` currently refuses an unverified account and requires an
  `AccountRecord`. It must instead accept a contact email and no account at all.
- `submissions.account_id` stays nullable (0001 already allows it) and gains a contact address
  alongside. A submission from a maintainer still records the account.
- Per-account rate limiting does not apply to anonymous submissions; per-IP limiting and the
  existing blob quotas carry that load instead. **This is a real cost of the decision** - the disk
  exhaustion vector in threat 6 is less well defended - and is the reason the blob size cap and the
  24-hour abandoned-upload sweep matter more than they did.
- No sign-in UI is needed in CRT for contributing. One is still needed eventually for maintainers,
  but that belongs to Phase 5/6, not here.

**The server has been corrected to match (2026-09-21).** What changed:

- `SubmissionFlows.CreateAsync` takes a `Submitter` - an optional account id plus a contact
  address - instead of a required verified `AccountRecord`. An anonymous submitter with an address
  is the ordinary case.
- **Ownership is now a CAPABILITY TOKEN.** With no account there is no identity to authorise
  against, and a submission id is a small consecutive integer anyone can guess. Creating a
  submission returns a random 256-bit token; every later call presents it in `X-Submission-Token`.
  Only the hash is stored. **This is the authorisation boundary of the whole phase.**
- `0003_anonymous_submissions.sql` adds `contact_email`, `created_ip` and `upload_token_hash`.
- **`GET /api/submissions/mine` is GONE**, replaced by `GET /api/submissions/{id}` by token. With
  no account there is nothing to list against, and keying on the contact address would recreate
  exactly the identity this design avoids - anyone knowing an address could enumerate that
  person's contributions. **The client keeps its own record** of what it submitted.
- **Acknowledged cost:** per-account rate limiting no longer exists for contributions. Per-IP is
  weaker (a household, an office, CGNAT all share one address, and an attacker can rotate them),
  so the blob size cap and the 24-hour abandoned-upload sweep now carry more weight than planned.

### Client side (task 5) - the plumbing is done, the button is not

- **`SubmissionManifestBuilder`** (CRT.Data, pure) turns merged board data into a manifest. Files
  are derived from the ROWS, never a directory walk - a walk would sweep up editor backups and
  upload a contributor's unrelated files to a public server.
- **`SubmissionProgress` and `SubmissionRetryPolicy`** (CRT.Data, pure). Progress reports the phase
  and the current file rather than one percentage, because "40%" tells a contributor nothing about
  whether it is stuck. `Fraction` returns null when nothing is measurable, so the bar goes
  indeterminate instead of sitting at zero looking broken.
- **`SubmissionClient`** (CRT.App, an untested I/O boundary) hashes streamed, negotiates, uploads
  in 4 MB chunks and finalises. It ASKS THE SERVER where to resume rather than assuming zero,
  which is what makes resumption survive a client restart and not merely a retry.
- **`SubmitDraftWindow`** - one window, three panels (Confirm, Progress, Outcome). Shows what will
  be sent before anything goes, takes the contact address and the summary, and says plainly that
  no account is being created. Enter does NOT submit; this dialog sends work to a public server.
- **`EmailAddressRules` moved to CRT.Data.** The app needs it to enable the Submit button and the
  server needs it to validate a contact address - the exact shared-rule case CRT.Data exists for.
  `AccountRules` now delegates, so there is one implementation and the server-side tests still
  exercise it through the old names.

**All of this landed.** The Submit button is wired into the Drafts tab
(`TabDrafts.OnSubmitClick`, with a `SubmitWindowFactoryForTests` seam beside it) and the "my
submissions" view is built (task 6, see below). `system.json`'s writer (task 7) is the only
remainder, and it belongs to publishing.

**"My submissions" needed rethinking under this model** - with no account, there is nothing to log
in and look at. The answer was that the client keeps its own record of what it submitted and polls
by submission id, which needs no identity at all; see "My submissions, with no account to hang it
on" below for what was built and why.

**Goal.** CRT can submit a whole system as a delta; the server diffs it against the base revision
and queues it.

### The transport

Semantically the client submits the **complete system**; physically it uploads only what the server
lacks:

1. Client builds a manifest of the system as it would be after the change: every file path with its
   SHA-256, plus the board rows. This is small - kilobytes of JSON.
2. Client POSTs the manifest and the base revision; server answers with the hashes it does not have.
3. Client uploads only those blobs.

A typo fix uploads a manifest and nothing else. A new 76 MB system uploads everything once, then
only changes. Shared images already present under another board are never re-sent, because the hash
is the identity. SHA-256 is already what `dataChecksums.json` uses, so this reuses a proven idea.

**[FIXED 2026-09-25] "What the server lacks" did not count the PUBLISHED TREE, so that promise
was false for the first submission to every board.** The server asked only its blob store, which
holds what somebody has UPLOADED - and the shipped tree was never uploaded. Changing one
component's text on a shipped board therefore sent the whole board (reported: 1,212 files,
121 MB), and so did any resubmission after a rejection or a cancel, once the collector had
removed the earlier blobs. The tests only ever proved the second submission of content the blob
store had already seen.

Now `SubmissionFlows.CreateAsync` takes a file the published tree holds **at the same path with
the same hash** (`PublishedTreeView.HashOf`, the cached hash the file rules and the review
comparison already use) from this server's own disk, via `BlobStore.TryImportAsync`, and counts
it as held. Decisions worth keeping:

- **Imported INTO the store, not read from the tree later.** The maintainer, finalise's content
  check and the publish all read the store, so none of them changed. A file only NAMED as "in
  the tree" could be overwritten there by another board's publish (a shared file) before this
  submission was reviewed, leaving it citing bytes that exist nowhere.
- **Verified exactly like an upload**: hashed as it is copied, moved in only on a match. The
  tree's hash is a cache keyed on length and write time, so it is a reason to try, never proof.
  A failed import falls back to asking the contributor for the file.
- **Copied, never hard-linked.** A link would let an in-place edit of the published file change a
  blob that is meant to be immutable, and `Contains` would then keep answering "held" for bytes
  that fail every finalise.
- **Disk use is unchanged.** The store ends up holding exactly the bytes the upload used to put
  there; they now come off the local disk instead of over the contributor's connection. They
  still count against the disk reserve, but not against the sender's upload budget: they cost
  the sender nothing, and a sender cannot make the published tree any bigger.
- **Same path only.** Identical bytes under a different published name would need the whole tree
  hashed to find. Untouched files of the board being edited are the case that cost real uploads.

The client did not need to change for this, and older builds get the saving too, because they
already upload only `MissingHashes`. The one catch is their 5-second create timeout, which a
large board's FIRST import could run past. A retry then finds the files already imported and
answers quickly. Two client fixes shipped alongside it: create and finalise
now have their own 2-minute timeout (`AppConfig.SubmissionRequestTimeout`). They used to share the
5-second `ApiTimeout`, too short for a large manifest going over a slow upload line, and a create
that times out after the server has done its work loses the capability token for good. And the
progress line carries its byte count across files (`SubmissionProgress.AfterFileSent`). It used
to restart at "0 bytes" for every file, so the bar sat empty for the whole upload.

Uploads must be **resumable per blob**. A contributor whose 300 MB upload dies at 90% and must start
again will not start again.

### Retiring the UUIDs

Currently rows carry `UuidV4` so the PHP diff can pair them. Under base-revision diffing the server
compares two states it fully knows, so row identity need not live in the data.

- Stop writing `UuidV4`; pair rows on **natural keys** (component: board label; component image:
  label + region + pin + name; highlight: schematic + label).
- Keep **reading** any existing UUID as a tie-breaker hint during transition.
- **No migration pass** over existing spreadsheets, no flag day. The columns can be deleted whenever
  convenient, or left in place doing nothing.
- The one case natural keys lose is a **rename** (U8 to U9 reads as delete+add). Handle it where the
  knowledge exists - the client knows it was a rename - by carrying explicit rename intent in the
  submission: `renames: [{ section, from, to }]`.

### Tasks

1. Define the submission contract in `CRT.Data` so client and server share one definition. Include a
   format version from the start, as the current payload does.
2. Server endpoints: create submission, negotiate hashes, upload blob, finalise.
3. Server-side diff of base revision vs submitted state, using `CRT.Data`. New-system vs update is
   decided by whether `SystemId` already exists - one lookup, which is the "server figures it out"
   behaviour the project owner asked for.
4. **Automated validation before any human sees it.** This is the highest-leverage work in the whole
   plan: it turns most bad submissions into a fast, automatic, polite answer.
   - run `DataValidator`;
   - references to files absent from the manifest - **and this check must compare file names
     CASE-SENSITIVELY even when the server's own filesystem is not.** From Phase 3 the data tree is
     read on Linux, where `CRT.Data`'s exact-case `File.Exists` calls mean a row naming `foo.pdf`
     when the file is `foo.PDF` loads fine on the contributor's Windows machine and fails only on
     the server. The shipped tree was swept on 2026-09-20 and is clean (4422 references, 0
     case-only matches), but that holds by care rather than by any check, and mixed case is real in
     the tree (27 `.PDF` alongside 85 `.pdf`). Do **not** "fix" this by making the lookups
     case-insensitive: that hides contributed data that is wrong, and Windows clients would still
     disagree with the server about which file is meant;
   - highlights naming schematics that do not exist;
   - duplicate board labels;
   - images that fail to decode;
   - degenerate highlight rectangles (zero-sized, outside the image);
   - implausible sizes or counts.

   Hard failures are rejected automatically with a clear explanation and never queued.
5. Client side: a Submit action per drafted system, showing exactly what will be sent, upload
   progress, and the outcome. Submitting does **not** clear the draft - it stays until the
   submission is accepted, so the contributor keeps using their work. **[DONE 2026-09-21]** - see
   "The Submit action" below.
6. A "my submissions" view in CRT showing state and any maintainer comment. **[DONE 2026-09-21]** -
   see "My submissions, with no account to hang it on" below.
7. `system.json` per system in the published tree:
   `SystemId`, `Manufacturer`, `Hardware`, `Board`, `Revision`, `PublishedUtc`, `Maintainers`,
   `Origin`, `ContentHash`. **[FORMAT AND READ PATH DONE 2026-09-21; the WRITER belongs to the
   publishing step, which does not exist yet - see below.]**

   **This task's own wording said `SystemId` "is derived once at creation and never changes, so
   history and maintainership survive a rename", and that sentence is what sent the first attempt
   down a wrong path** - it reads as an argument for a surrogate id, while `0001_initial.sql` had
   already decided the id IS `manufacturer/hardware/board`. Nothing in CRT can rename a system, so
   there is no rename to survive. See task 7's section below.

### The Submit action (task 5, done 2026-09-21)

**One button per drafted row on the Drafts tab**, beside "Images and KiCad" and "Discard". Disabled
rather than hidden when the draft carries no rows - which is a real state, not a theoretical one,
because "Add a new system" creates a registration that this tab lists before a single row exists.
A hidden action reads as the app having lost it, and the tooltip is the only place the reason can
actually be given.

**Drift does NOT block submitting.** The server diffs against the base revision itself, so refusing
to send while the official data has moved would strand a contributor behind a change somebody else
made.

**What is sent is the MERGED board, not the draft** - official rows with the drafted ones applied
over them. The server needs the system as it should READ after publishing; sending only the drafted
rows would make it reconstruct that merge from a draft format it has no reason to know.

**`BaseRevision` comes from the draft, never from the official file's current value.** Reading the
current one at submit time would claim the contributor had seen changes they never saw.

#### The two-root bug this uncovered - read this before touching the submission file paths

The first version resolved every referenced file against a SINGLE root, the system's draft folder.
That is correct only for a system created from nothing. For the ordinary case - a draft over a
published board - **a system's files genuinely live in two places**:

- an officially published file lives under `Data/`, the sync-owned tree;
- a drafted file lives under `<draft system folder>/Files/`, because `Data/` is freely overwritten
  by the next sync and an attachment left there would simply vanish.

So a one-line typo fix that adds a single photo to a board with 240 existing images references BOTH
roots in one submission. Resolving against the draft folder alone would have reported all 240
published files as "not on disk" and refused to send. **Nothing caught this**: the server side was
tested thoroughly, the manifest builder was tested thoroughly, and the bug sat in the seam between
them where one root was passed instead of two. It would have surfaced on a real contributor's first
real submission.

`SubmissionFileLocator` in CRT.Data is now the single answer to "where are this file's bytes",
applying the SAME two-root precedence `DraftFileResolver` already used for opening a file to show
it on screen - official first, drafted second. That order matters: what is uploaded has to be what
the contributor has been looking at. What the locator adds over `DraftFileResolver` is the
containment check on each candidate root, because this path is about to be read and posted to a
public server.

The second half of the same bug: the upload loop called `SubmissionPathRules.TryResolve` and
**discarded its return value**, so a refused path was handed to the uploader as an empty string and
would have been reported as a failed upload rather than as the bad path it is. Both call sites now
go through the locator and both check the result.

### My submissions, with no account to hang it on (task 6, done 2026-09-21)

**The problem this had to solve first.** The task says "a view showing state and any maintainer
comment", which quietly assumes the app can ask "what did I send". It cannot: contributing needs no
account, so the server has no idea who is asking. The capability token returned once at creation is
the only proof of ownership - and it was living in a local variable inside the Submit dialog and
being discarded when that dialog closed. **Nothing persisted it**, so before task 6 could be built
at all, that receipt had to start being kept.

**Decision (owner, 2026-09-21): keep receipts LOCALLY.** CRT writes `{id, token, systemId,
summary, sentUtc}` to `submissions.json` beside the user's settings, and the view reads its own
list. The alternative considered was an email-address lookup endpoint, which works across machines
but needs a mailed link or code to be safe - an address is not a secret, so without one anybody who
guessed an address could read that person's submission history. That is most of an account by
another name, and it was rejected for Phase 4.

**The cost is real and is stated in the UI rather than hidden**: a reinstall or a second computer
starts with an empty list. The window's own header says so, and points at the email as the channel
that does not depend on the file. Do not "fix" this later by silently weakening the lookup.

**Where the file lives is load-bearing.** AppData, never the Data tree and never a draft folder:
a receipt carries a capability token, the Data tree is synced, and a draft folder is the very thing
the submit path walks and uploads. Receipts also outlive the draft they came from, so tying their
lifetime to a draft would delete them exactly when they are still wanted. `SubmissionReceiptStore`
deliberately takes no command-line root override for the same reason, unlike data/workbooks/drafts.

**The receipt is written immediately after `CreateAsync`, before a single byte uploads.** Recording
it at finalise would lose it for exactly the submissions most worth looking up - the ones that
failed partway - and the loss is unrecoverable, since the server stores only the token's hash.

**The state vocabulary is translated, not shown raw.** `SubmissionReceiptPresenter` (pure, unit
tested) turns the server's database words into what they mean to a contributor. Two matter most:
`pending` means QUEUED but reads as "not sent yet" to somebody who just pressed Submit, and
`uploading` means a send that never completed - nothing is uploading any more, so rendering it
literally leaves someone waiting for a transfer that stopped days ago. **An unrecognised state is
reported as unrecognised rather than guessed at**, because Phase 5 will add states this build has
never heard of and a plausible-but-wrong rendering would never be caught.

**A decided submission is never asked about again** (`IsStillOpen`), which is what keeps the
refresh cheap for a long history. An unknown state counts as still open: the cost of being wrong
that way is one extra request, whereas treating it as final would freeze a row forever.

**An unreachable server leaves rows alone.** `GetStatusAsync` answers null rather than throwing, so
one bad row does not fail the refresh, and the cached state stays - that cache exists precisely so
the window says something true offline.

**`MaintainerComment` is in the contract now and empty until Phase 5.** There is no maintainer and no
column behind it yet. Carrying it from the start means filling it in later is a server change
alone, with nothing to update on any contributor's disk.

**"Remove" is not "withdraw", and the dialog says so in those words.** Removing a receipt deletes
the only copy of the token and with it any further ability to check on that submission from this
machine; the contribution itself is untouched and still gets reviewed.

### system.json (task 7, format and read path done 2026-09-21) [RETIRED 2026-09-25]

> **RETIRED by the project owner, 2026-09-25:** "I do not want this file visible in the source ... it
> should not be something downloaded by all users, as this file is not relevant for them." It was
> synced to every user while CRT showed nothing from it, and every fact it held is in the database
> (`systems.current_revision`, `content_hash`, `origin`; the `maintainers` table). Nothing writes or
> reads it now: a publish removes one left in the board's folder by an earlier build, a production
> promotion never carries one and removes one already there (`RetiredSystemDescriptor`), and
> `DataManager.LastLoadedSystemDescriptor` - read by nothing - is gone. `SystemDescriptor` survives
> only as the in-memory result of a publish (revision + content hash). Everything below is history.

**What landed:** `SystemDescriptorRules` (id and content hash), `SystemDescriptorStore` (read and
write), a `SystemId` check in `SubmissionValidator`, `DataManager.LastLoadedSystemDescriptor`, and
migration `0004` plus the `systems`-row insert that Phase 4 needed and did not have.

**What did NOT land, and why it must not be forced:** nothing writes a `system.json` into the
Production tree. **That is the Phase 3 step 0 interlock working as designed** - `crt-server` has no
write permission there and the kernel refuses it. A descriptor is written at PUBLISH time, and
publishing is still a deliberate act by the project owner with no tool behind it yet. `Write` exists
for that future tool and for tests; **do not wire it into a request path to "finish" task 7.**

**[RESOLVED 2026-09-21] `SystemDescriptorStore.Write` now has its call site: `PublishExecutor`, at
publish time** (Phase 5 task 6). That is precisely where this note said it belonged, so the
instruction above was followed rather than worked around - the writer waited for publishing to
exist instead of being wired into a request path.

**The Production interlock is UNCHANGED and still absolute.** `PublishExecutor` writes the BETA
tree only. BETA to Production remains a manual file copy the project owner performs (open question 6),
the service has no write permission on Production, and no "publish to production" option may be
added.

#### THE SYSTEM ID IS "Manufacturer/Hardware/Board" - and a wrong turn is recorded here on purpose

The id was first built as a RANDOM 128-bit value, with a validator that refused anything else. That
was wrong, and the way it went wrong is worth more than the fix:

- **The decision already existed**, written into `0001_initial.sql` on the `systems` table: "system_id
  is the STRING key the data tree already uses (manufacturer/hardware/board), not a surrogate
  number, because it is what the desktop app, the file tree and every submission already name." It
  was not read before a conflicting design was built on top of it.
- **The argument for a surrogate id was rename-survival, and CRT CANNOT RENAME A SYSTEM.** A
  system's folder path is its identity in the sync manifest, in `BoardDataReader`'s cache key, in
  `DraftManager.GetSystemFolder` and in the worklog board key. `SubmissionManifest.Renames` covers
  rows INSIDE a board, never the board itself. The problem being solved did not exist.
- **A test caught it and was overruled.** `SubmissionFlowTests` carried `"commodore-c64-250407"`,
  which the new validator refused; it was "fixed" to a hex id and described as a stale fixture. It
  was not stale - it was correctly modelling the schema, and the red test was the system saying so.

If a system ever does need renaming, that is a migration with a redirect, not a reason to carry a
second identity forever on the chance.

**`IsValidSystemId` reuses `NewSystemIdentity.IsValidPathSegment` per segment**, deliberately. That
rule already refuses traversal, reserved device names, trailing dots, control characters and every
character ANY platform reserves. A second opinion here about what a folder name may be would
eventually disagree, at which point a system could be created locally and refused by the server, or
accepted by the server and impossible to sync on Windows.

**The validator now requires the id AND requires it to match the names it carries.** Both are
errors: the value is a database primary key and resolves a folder in the published tree, and a
mismatch means everything keyed off the id disagrees with what is printed on the screen.

#### `ContentHash` is over an ORDERED, DELIMITED file list

Each of those words is load-bearing:

- **ordered**, because a directory walk returns whatever order the filesystem feels like and two
  machines hashing the same tree must agree. Sorted **ordinally** - a culture-aware sort reorders
  under `tr-TR`/`sv-SE`, so the hash would change with the server's LOCALE, which looks like every
  client being permanently out of date. Pinned by a test that actually switches culture.
- **delimited**, because plain concatenation is ambiguous: `("ab","cd")` and `("a","bcd")` produce
  the same byte stream, and the file names come from a contribution so an attacker picks them. A
  path containing the delimiter is refused rather than escaped, because `SubmissionPathRules`
  already refuses control characters.
- the **revision is folded in**, so a metadata-only republish still moves the hash.

#### `Origin` was never an open question either

It is a `systems` column that already existed, holding `shipped` (came with CRT) or `contributed`
(arrived through this pipeline and was vetted). **Set once when the system row is created and never
recomputed** - a contributed system stays contributed however many times it is later revised,
including by the project owner, because it records where the system CAME FROM. Migration `0004` adds a
CHECK so the database refuses a third value.

#### Three schema-versus-code disagreements, found by reading and not by testing

**The whole submission pipeline is tested against in-memory fakes, so `MySqlSubmissionStore` had
never executed against MariaDB.** Against the real database every submission would have failed:

1. **`submissions.state`'s CHECK did not allow `uploading` or `abandoned`** - the first two states
   the code writes. `0001` listed the REVIEW vocabulary (Phase 5's), and Phase 4 later added the
   TRANSPORT states without revisiting the constraint.
2. **`submissions.system_id` is `NOT NULL` with a foreign key to `systems`, and nothing ever
   inserted into `systems`.** The table was empty. `CreateAsync` now inserts the system row first,
   with `INSERT IGNORE` so two simultaneous first submissions cannot race, and an existing system
   is left completely alone - an `ON DUPLICATE KEY UPDATE` would rewrite a shipped system's origin
   the first time somebody corrected a typo in it.
3. **Nothing recorded a system's published revision or content hash**, which task 7 needs.

**The fake store was MORE PERMISSIVE than the real one, which is what hid this.** It accepted a
submission for a system that did not exist and threw the name parts away. It now records the
`systems` row the real store must insert, and two tests assert it - both fail against the old
behaviour. **A fake that is more permissive than the thing it stands in for certifies bugs.**

**Still open for whoever builds publishing:** where `system.json` is actually written, and what
fills `systems.current_revision` / `content_hash` at publish time. Both belong to the publishing
step and neither is decidable without it.

### Definition of done

- A one-component edit submits in seconds and uploads no blobs.
- A new system submits, resumes correctly after an interrupted upload, and appears queued.
- A submission failing validation is rejected automatically with a readable reason.
- No `UuidV4` is written by new code; existing values are still read.

### Traps

- **Do not let submission block the UI.** Upload off the UI thread, cancellable.
- **Path safety.** Every path in a submitted manifest is untrusted input. Reuse the reasoning in
  `OnlineServices.TryResolveValidatedLocalPath` and `ExternalTargetLauncher`: reject absolute paths,
  traversal, and anything resolving outside the target system folder. Add tests for each.
- **Blob store hygiene.** Blobs from abandoned submissions must be garbage-collected, or the disk
  fills quietly.

---

## Phase 5 - Maintainer application, single user [STARTED 2026-09-21]

> **Superseded as an application by [Phase 8](#phase-8---maintainer-tab-inside-crt-done-2026-09-29)
> (2026-09-29):** everything below was built as the separate CRT Maintainer application, and now runs
> as CRT's Maintainer tab. This phase is kept as the history of how it was built.

**Goal.** A separate Avalonia desktop app where the project owner reviews and merges submissions. One
maintainer only; roles come in Phase 6.

### Progress

**Started with the PUBLISHING STEP, not the app** (owner's choice, 2026-09-21). Publishing is
what finally unblocks Phase 4 task 7's `system.json` writer, it is pure logic that unit tests can
cover properly, and the maintainer app has nothing worth reviewing until something can be published.

**Landed so far: `DataGenerationRules` in CRT.Data** (+34 tests; suite 4148 to 4182), which owns
the project owner's "newest generation only, never touch an older one" rule:

- `ResolveNewestGeneration` DISCOVERS the target from the tree rather than reading a setting -
  see open question 3 for why a configured generation fails silently.
- `IsOlderGeneration` is the guard that protects a frozen generation, expressed as its own
  question rather than left to each call site to compare versions. The unversioned original is
  older than every real generation, which is the case the project owner named outright.
- Both shipped naming conventions are pinned: the master separates its version with a DOT
  (`Classic-Repair-Toolbox.v2.0.0.xlsx`), a board file with a SPACE
  (`Data C64 250407 v2.0.0.xlsx`). A test proves `Data VIC20 250403` is not read as a version.
- Versions compare NUMERICALLY - a text sort puts `v10.0.0` before `v9.0.0`, which would publish
  into a frozen generation.

Verified against the real tree: the rules find exactly the two generations that exist there
(unversioned and `2.0.0`), newest `2.0.0`.

**Task 7 was STRUCK on the same day** (no retained revisions - open question 5), which removes the
revision store, the content-addressed publish archive and the rollback action from this phase
entirely. Publishing overwrites the newest generation's files in place.

**Landed next: `BoardWorkbookSchema` + `BoardWorkbookWriter`** in CRT.Data (+14 tests; suite 4182
to 4196). This is the first code in the project that WRITES a board workbook since
`BoardDataWriter` was retired in Phase 2.

- **The schema was EXTRACTED, not duplicated.** `BoardDataReader` held every sheet and column name
  privately, which was fine while reading was the only direction. A writer with its own copy is
  exactly the defect `SubmissionContract.cs` exists to prevent, and the failure is SILENT: a
  workbook written `Part number` and read `Part-number` loads with every part number blank and
  nothing throws. The reader's constants are now aliases onto `BoardWorkbookSchema`, so a rename
  moves both directions at once.
- **Every cell is written as TEXT, with an explicit `@` number format.** `BoardData` holds
  opacities, T/DIV and trigger levels as strings; letting EPPlus store them as numbers means the
  reader's `.Text` comes back formatted by the CURRENT CULTURE. Verified by injecting that exact
  bug: under `da-DK`, `0.35` round-tripped as `0,35`. This is the same class of bug already
  written up for highlight coordinates and `WorkbookSummary`'s `Stat` parts - the third face of it.
- **`UuidV4` is neither invented nor stripped**, matching Phase 4's "stop writing new ones, keep
  reading existing ones, no migration pass".
- **The strongest test is the INDEPENDENT cross-check.** `BoardWorkbookBuilder` (a test helper
  that hand-writes the sheets and knows nothing of the schema) builds a board; it is read,
  rewritten through the writer, and read again. **Its first version was VACUOUS** - comparing
  read-one against read-two passes happily when a misspelled column makes both blank. Caught by
  renaming `ColPartNumber` and watching all 14 tests still pass. It now asserts the fixture's
  known VALUES before comparing, and fails against that rename.

**Landed next: `PublishPlan`** in CRT.Data (+31 tests; suite 4196 to 4227). Decides everything a
publish will do BEFORE a byte is written, and answers a different question from
`SubmissionValidator` - not "is this data any good" but "is it safe to write these bytes to these
paths".

- **A plan object rather than just doing it**, because publishing is irreversible now that
  revisions are not retained. Every refusal is found up front, rather than halfway through a write
  that has already replaced half a board. A refused result carries **no plan at all** - a
  partially valid publish is not a thing that may proceed.
- **The generation guard lives here.** The target is resolved from the tree via
  `DataGenerationRules`, and **a submitted `.xlsx` is REFUSED outright**: the board workbook is
  GENERATED from the submitted rows, so a submission carrying one could otherwise overwrite any
  generation's workbook - including the unversioned original that serves every pre-2.0.0 build,
  silently. The rule is deliberately broad (any `.xlsx` anywhere in the submission) rather than
  trying to recognise a generation's naming; erring permissive there overwrites a frozen
  compatibility target, and no shipped board references a spreadsheet as an attachment.
- **Case-only collisions are refused and BOTH spellings are named.** Both files can exist on the
  Linux server and only one on a Windows client, so the tree would be un-syncable for much of the
  audience.
- **The content hash is pinned three ways**: order-independent (two machines must agree), moved by
  a changed file (anti-vacuity for the ordering test), and moved by a changed revision alone (a
  metadata-only republish must still reach clients).
- **A rows-only change plans successfully with zero files** - the typo fix that uploads nothing is
  the commonest contribution there is.

**Landed next: `PublishExecutor`** in CRT.Server (+13 tests; suite 4227 to 4240). The half that
touches disk, so it lives in the server rather than in CRT.Data. **This completes task 6 and
unblocks Phase 4 task 7's writer** - `SystemDescriptorStore.Write` finally has its call site, at
publish time, which is where task 7 always said it belonged.

- **The write ORDER is the safety property, and it is tested.** Publishing cannot be made atomic -
  a system is hundreds of files and there is no rename that swaps them all at once - so the writes
  are ordered such that a partial publish is RECOVERABLE: files first (inert until something
  references them), then the workbook (which makes them reachable), then `system.json` (which
  advertises the revision), then the database row. An interrupted publish leaves files nothing
  points at, which is wasted space rather than a board that fails to load. Verified by moving the
  workbook write to the front and watching `A_missing_blob_stops_the_publish_BEFORE_the_workbook_is_touched`
  fail.
- **Re-running the same publish is safe, and that is the recovery.** Every step overwrites and the
  blobs are content-addressed, so a re-copy is byte-identical. This matters more than it would
  otherwise, because there is no previous revision to roll back to.
- **The published board is asserted to LOAD**, through the ordinary `BoardDataReader` rather than
  a test-only parser that could agree with a broken writer. `CRT.Server.Tests` was added to
  `CRT.Data`'s `InternalsVisibleTo` for this - following that csproj's own instruction to add
  callers there rather than widening a type to public.
- **`ISubmissionStore.SetSystemPublishedAsync` is new**, because nothing wrote
  `systems.current_revision` or `content_hash` - the gap task 7 flagged and could not close.

#### A fourth schema-versus-code disagreement, caught the same way as the first three

The first version of `SetSystemPublishedAsync` wrote an `updated_utc` column. **`systems` has no
such column** - it has `created_utc` and nothing else time-shaped. It would have thrown against
MariaDB on the very first publish while every test passed, because the tests run against the
in-memory fake. Found by reading `0001_initial.sql` rather than by any test, exactly as migration
0004's own three were. The rule that catches these is worth repeating: **after writing SQL against
this schema, enumerate the table's real columns and check each one.**

The fake was given the matching method and **made deliberately STRICTER than convenient**: it
throws when a publish names a system with no `systems` row, because the real statement is an
`UPDATE` that would silently affect zero rows. That is the same lesson already recorded here - a
fake more permissive than the real store certifies bugs.

**Also added: the four Phase 5 review states** (`merged`, `changes_requested`, `approved`,
`withdrawn`) as `SubmissionState` constants. They were in `0001_initial.sql`'s CHECK constraint
from the start but had no constant, since Phase 4 only ever wrote transport states. Constants
rather than literals matter here because the database CHECK rejects an unknown value - a typo is a
publish that throws at the very last step, AFTER the tree has been written. All eight constants
were cross-checked against the CHECK constraint and match exactly.

**Landed next: `ReviewSummary`** in CRT.Data (+21 tests; suite 4240 to 4261) - task 3's screen,
built as pure logic BEFORE any UI so the maintainer app's central view is unit tested rather than
verified by eye.

- **Rows pair on NATURAL KEYS, never on position.** Pinned by a test that inserts a row at the top
  and asserts the rows below it are not reported as changed - the classic diff bug, and it would
  be worst on exactly the submissions that add something, which is most of them.
- **`UuidV4` is excluded from the comparison.** Phase 4 retired it and stopped writing new ones, so
  comparing it would report a change nobody made on every row of every pre-transition board.
- **A declared rename is UNTRUSTED INPUT** - it arrives inside the submission - and is honoured
  only when it describes what is actually there.
- **A new system is summarised as "New system, N rows"** rather than listing every row as an
  addition, which would be true and hundreds of items long.

#### Two bugs the tests caught, both in the code rather than the test

Written to the intended behaviour first, then run - and both failed, which is the point:

1. **A clean rename reported "also changed".** The renamed value IS one of the compared fields (a
   component's `BoardLabel` is its key), so a whole-row comparison counts every rename twice. A
   maintainer reading "renamed, and something else changed" looks for an edit that is not there;
   worse, the signal that would matter - an edit hiding behind a rename - becomes meaningless.
   Fixed with `RowsMatchIgnoringKey`, which drops the key's parts BY VALUE rather than by index,
   since a key's fields are not always the leading ones.
2. **A rename that never happened was honoured.** The guard checked that the new key exists but
   not that the old one is GONE. With `U8` still present, a submission declaring "U8 became U9"
   had `U9` absorbed as a rename and never shown as the ADDITION it was - hiding an added row from
   the person approving it, which is the one thing this screen exists to prevent.

**Landed next: `src/CRT.Maintainer/` and `tests/CRT.Maintainer.Tests/`** (task 1; +11 tests; suite 4261 to
4272). The app builds, launches and shows the queue-plus-summary layout task 3 describes. Both
projects are in the solution, so CI builds and tests them with everything else.

- **It follows CRT's conventions**: no MVVM, code-behind, a file map in the window's header ready
  for partials, and **pure logic in `Handlers/`** - `ReviewSummaryPresenter` decides what the
  landing view SAYS and is unit tested, while the window only walks the result and makes controls.
  That split is why the screen deciding whether a change gets looked at has real coverage.
- **Removals are listed FIRST within a section**, pinned by a test. A removal is the least
  recoverable thing a submission can do and the easiest to skim past, because a maintainer scanning
  for "what did they add" is not looking for it.
- **Sections with no changes are omitted.** Ten lines of "0 changed" buries the one that matters -
  the same failure as opening on the whole board, just smaller.
- **Its own version (`0.1.0-alpha.1`) and its own `AssemblyName`**, per task 8, so it can never
  get entangled with CRT's release.
- **`MaintainerApp` is deliberately EMPTY beyond showing the window.** CRT's own `App` initialises a
  logger, shows a splash and syncs over the network, which is exactly why its headless tests need
  a subclass with an empty override. Keeping startup work out of here means this app's tests never
  need that workaround.

**Not wired to the server yet, on purpose.** The queue is populated by `SetQueue`, which only
tests call. Talking to `/api/` is task 2's other half and lands with the review flows; a stubbed
HTTP client now would be a second, untested idea of the API's shape - the exact defect
`SubmissionContract.cs` exists to prevent.

**Landed next: the review API and its authority model** (task 2's server half; +31 tests; suite
4272 to 4303). `GET /api/review/queue` and `GET /api/review/submissions/{id}`, in their own
`ReviewEndpoints` file.

- **Separate from `SubmissionEndpoints` because the authorisation model is the OPPOSITE.** Those
  are for contributors: no account, a capability token, and 404 covering both "no such thing" and
  "not yours" so the id space cannot be walked. These are for maintainers: an account and a role are
  required, there is no token, and 401 and 403 are DISTINGUISHED - a maintainer whose account lacks
  the role needs telling that, not a login prompt that will not help. Mixing the two in one file
  is how a route eventually gets mapped into the wrong group and silently inherits the wrong rule.
- **`ReviewAuthority` answers the authority question in ONE place**, written now rather than
  retrofitted at Phase 6, because a rule introduced after its call sites exist has to find them
  all. Phase 6's own trap says exactly this: "compute authority once and use it everywhere".
- **A REVIEWER (the old recommend-only role, retired 2026-09-25) MAY NOT PUBLISH, and that is
  enforced structurally.** Phase 6's role table gave Reviewer a blast radius of "none - no
  published data can change", which held only while `CanPublish` refused the role. Verified by
  widening it to admit reviewers and watching `A_REVIEWER_MAY_NOT_PUBLISH` fail - and only that test.
- **Locked and unverified accounts are refused whatever their role**, so withdrawing access bites
  on the very next request rather than at next login (Phase 6's definition of done).
- **The queue is `pending` only, OLDEST FIRST** - the opposite of "my submissions" and
  deliberately so. A work queue is worked from the front; newest-first lets a steady trickle of
  new contributions bury the one that has waited a month, which is how a contribution quietly
  never gets reviewed.
- **The upload token hash is NEVER echoed to a maintainer.** It is the contributor's capability for
  that submission, and handing it over would let a maintainer act as them.
- **`ReviewApiRoutes` (CRT.Maintainer, pure)** builds the URLs, separately from the HTTP client so the
  half that actually breaks is the half that is tested - a trailing slash on a human-typed base
  address, a path-hosted server, a blank address producing a relative URL that would post a
  maintainer's credentials somewhere unintended.

**Landed next: the client half of task 2** (+21 tests; suite 4303 to 4324) - `ReviewSession`,
`ReviewApiParser` and `ReviewApiClient` in CRT.Maintainer.

- **The client is deliberately THIN, because it is an untested I/O boundary** (test rule 6). The
  routes are in `ReviewApiRoutes` and the parsing in `ReviewApiParser`, both pure and both
  covered; what is left in the client is sending a request and reading a status code.
- **A failure is a VALUE, never an exception.** It is called from UI event handlers, where an
  unhandled task exception is a crash, and a maintainer whose connection dropped needs a sentence
  rather than a stack trace.
- **An unreachable server and an EMPTY queue are different answers**, carried all the way through
  the parser and the client. If they collapse into one, a broken connection reads as "nothing to
  review" and a backlog goes unnoticed for as long as nobody happens to check.
- **A timeout is caught separately from a cancellation.** `HttpClient` reports its own timeout as
  a `TaskCanceledException`, so without that split a timeout reads as "the user cancelled" and is
  silently swallowed.
- **One malformed queue row is skipped, not fatal** - a single bad record must not conceal every
  other contribution behind it.

#### `refreshToken` is the BEARER token, and the name is a trap

The login endpoint answers with a field called `refreshToken`, which reads as "exchange this for a
real token". It is not: `AccountFlows.AuthenticateAsync` looks up that exact value as the session
token, so it goes straight into `Authorization: Bearer`. A client author trusting the name goes
hunting for an access-token exchange that does not exist, gets 401s from every call, and blames
the server. `ReviewSession` names the value for what it DOES, and a test pins the mapping.

#### Two overflow bugs in one three-line method, both found by one test

`ReviewSession.IsUsableAt` applies a one-minute margin so a token cannot expire mid-request. Both
obvious spellings throw `ArgumentOutOfRangeException`:

- `ExpiresUtc - margin > now` overflows when `ExpiresUtc` is `MinValue` - **which is exactly what
  the parser records for an expiry it could not read**, so the one case the method exists to
  handle safely would have crashed the app;
- `ExpiresUtc > now + margin` moves the overflow to the other end and throws at `MaxValue`.

Fixed by applying the margin to the DIFFERENCE (`ExpiresUtc - now > margin`), which is a `TimeSpan`
comparison and cannot overflow a `DateTime` in either direction. **A method deciding whether to
keep using a credential must be TOTAL** - it is on the path of every request and is called from UI
code. Both ends are now pinned by their own test.

**Note on coupling:** `ReviewApiParserTests` holds JSON copied from the server's own private
response builders (`AccountEndpoints.SessionResponse`, `ReviewEndpoints.ToQueueRow`). It can go
stale. A renamed field shows up as a default rather than an exception, so the app degrades to "the
server sent something this version does not understand" instead of crashing - but the real fix is
a shared contract type, as `SubmissionContract.cs` already is for the submission pipeline. Do that
the first time one of those shapes changes.

**Landed next: sign-in and the live queue** (+16 tests; suite 4324 to 4340). **Task 2 is now
complete end to end** - the window signs in against `/api/accounts/login` and fills its queue from
`/api/review/queue`.

- **Sign-in is a PANEL, not a modal.** A modal would have to be dismissed and re-shown on every
  expiry and every "not permitted" answer, and a lapsed session would put a dialog over a window
  still showing stale rows. Swapping the whole surface means what is on screen is always true.
- **The password box is cleared the moment the session exists**, and the sign-in button is
  disabled while a request is in flight so two sign-ins cannot race and leave whichever finished
  last in place.
- **An expired session is caught BEFORE the request.** The server would refuse it anyway, but
  "your session has expired" is a better answer than "the server answered 401" and it is the same
  answer either way.
- **A failed refresh LEAVES THE LIST ALONE.** Clearing it would turn "I could not ask" into
  "nothing is waiting", which is the one wrong answer a review queue must never give.

#### The queue cannot show a change summary, and the first design assumed it could

`ReviewQueueItem` originally carried a `ReviewChangeSummary` per row. That cannot work: computing
one needs BOTH the published board and the submitted manifest, and the queue endpoint returns
neither - it answers ids, system names and summaries, which is kilobytes. Making it return enough
to summarise every row would mean loading two full `BoardData` per queued submission on the
server, for a list the maintainer scrolls past.

So the summary belongs to OPENING a submission, and the queue shows **what the contributor said**.
That is also the better screen, not merely the cheaper one: "Corrected R12." is what a maintainer
scans for, whereas "3 components changed" says nothing about whether it is worth opening next.

#### The waiting time TRUNCATES, deliberately

`ReviewQueueDisplay.Waiting` reports 13.9 days as "13 days", never "14" and never "2 weeks". This
is the number that shames a backlog, so it must not round in the comforting direction - a coarser
unit always makes a wait sound shorter than it is. Pinned by a test, verified by switching it to
`Math.Round` and watching that test fail. A missing timestamp shows **nothing** rather than the
1970 default, which would read as a submission waiting fifty-six years.

**Landed next: the change summary is computed SERVER-SIDE** (+13 tests; suite 4340 to 4353).
`GET /api/review/submissions/{id}` now returns a `changes` object alongside the manifest and
findings.

- **The server is the only place both halves exist.** The submitted board travels in the manifest;
  the published one is a workbook in the data tree only the server can read. Sending that workbook
  to the client so it could compare would ship multiple megabytes per submission opened, to
  compute a few hundred bytes.
- **`SubmissionBoardData.FromManifest`** (CRT.Data) is the inverse of `SubmissionManifestBuilder`,
  put in the shared library because BOTH the review screen and the publish step need to treat a
  submission as a board - two copies of that mapping would drift the first time a section was
  added.
- **A missing published board is an ANSWER, not an error** - it means a new system, the
  highest-risk submission there is, and `ReviewSummary.Compare` already takes a null for it. But a
  board that exists and cannot be READ throws instead of returning null: reporting that as "new
  system" would have a maintainer approve a replacement for a board they were told did not exist.
- **The locator reads the NEWEST generation**, the same rule publishing writes with. Reading a
  frozen one would show a diff against a board no current build uses, and the generation gap would
  appear as changes the contributor never made.

**EPPlus now executes on the server** (owner's decision, 2026-09-21), via a new
`InternalsVisibleTo` grant for `BoardDataReader`. Open question 8 predicted this would arrive at
Phase 4 task 3 and was wrong - it arrives here and at the publish writer.

**A gap worth knowing: the manifest carries no RevisionDate.** `SubmissionRows` has the ten board
sections and no revision date, so a revision-date CHANGE is currently invisible to the maintainer.
The endpoint uses the published board's date for both sides deliberately - the manifest's
`BaseRevision` is the date the contributor STARTED from, and using it would report a change
backwards on every submission built against an older revision. Fixing this properly means adding
the field to `SubmissionRows`, which is a contract change for both sides.

#### A VACUOUS SECURITY TEST, caught by trying to break the thing it guarded

The locator resolves an untrusted identity, so a traversal there is **file disclosure** - it would
hand a maintainer the contents of any file the service can read. Three traversal tests were written
(`..`, `../..`, `/etc`) and all three passed.

**Then `SubmissionPathRules` was removed from the locator entirely and all twelve tests still
passed.** The traversals were being refused by the "does this folder exist" check, not by
containment: every path they tried pointed at a folder that does not exist. A traversal is only
dangerous when it lands somewhere REAL.

`A_traversal_onto_a_REAL_folder_outside_the_tree_is_refused` now builds that layout - a data tree
beside a sibling folder holding a genuine workbook - and fails against the uncontained version. It
carries an anti-vacuity half proving the same folder INSIDE the tree is still found, so the
refusal is containment working rather than the locator finding nothing at all.

**The general lesson, which has now bitten twice this session** (the other was the read-write-read
board test): *a test that cannot fail proves nothing, and a security test that cannot fail is
worse than none, because it reads as coverage.* Delete the guard and watch the test go red -
especially when the guard is the only thing between an endpoint and a disclosure hole.

**Landed next: the client half - opening a submission** (+15 tests; suite 4353 to 4368).
**Task 3 is now complete end to end**: clicking a queue row fetches the submission and draws its
change summary and findings.

- **The wire shape was VERIFIED, not guessed.** A real `ReviewChangeSummary` was serialised
  through `System.Text.Json` with ASP.NET Core's camelCase policy before the parser was written.
  That caught two things reading the record declarations would have missed: the computed
  properties (`hasChanges`, `totalChanges`) also appear on the wire, and all ten sections ship
  every time including empty ones. The whole summary is about 1.4 KB, so no filtering was needed
  server-side; empty sections are dropped at the parser instead.
- **View types rather than reusing `ReviewChangeSummary`.** The server's record exists to be
  COMPUTED; the window needs a shape that exists to be DRAWN. Sharing it would have made any
  future change to the comparison's internals an accidental wire-format change.
- **The selected row is drawn IMMEDIATELY from what the queue already knows**, with the summary
  filling in when it arrives. Waiting for the request would blank the panel on every click over a
  slow link, which reads as the app having lost the selection.
- **A response whose selection moved while it was in flight is DISCARDED.** A maintainer arrowing
  down the queue starts a request per row and they can finish out of order; without the check, a
  slow earlier answer overwrites a faster later one and the panel describes a different submission
  from the one highlighted - exactly the state in which somebody approves the wrong thing.
- **`changes: null` is not "no changes".** The server answers null when a payload could not be
  loaded; rendering that as "no changes" would invite approving a submission nobody could look at.
  The window says so and shows the findings, which is where the reason is.
- **Findings are drawn below the summary**, and severity is read as a NUMBER or a STRING - the
  enum serialises as a number today, and a future `JsonStringEnumConverter` on the server would
  otherwise silently turn every error into a warning.

**`ReviewSummaryPresenterTests` now runs its comparisons through the REAL WIRE FORMAT** -
serialise exactly as ASP.NET Core does, parse with the app's own parser, then present. The two
sides share no type, so a hand-built view object would have let them drift silently; this fails
the moment a field is renamed on either side. It is the strongest test available without a live
server.

**Landed next: task 4's FIELD-LEVEL DIFF** (+24 tests; suite 4368 to 4392). Of task 4's four
sub-parts this is the one needing no image bytes, so it went first.

- **"U8 changed" is not reviewable; "Part-number: 906114 -> 251715-01" is.** Without this a
  maintainer had to open the board in CRT and hunt for what moved.
- **The field NAMES are the workbook's own column headers**, from `BoardWorkbookSchema`. A
  maintainer reading "Part-number" and someone opening the .xlsx are talking about the same
  column; inventing friendlier labels would create a second vocabulary that drifts.
- **A CLEARED field is shown as `-> (blank)`, never trailing off.** Deleting information is the
  edit hardest to notice and the one most deserving a second look; an empty right-hand side reads
  as the app failing to draw something.
- **Long values truncate in the MIDDLE.** A description runs to a sentence, and a real edit
  usually differs at one END - cutting the middle keeps both, so the maintainer sees what differs
  rather than two identical-looking prefixes.
- **Added and removed rows carry no field diff**, deliberately: there is nothing to diff against,
  and seven "(blank) -> value" lines would be noise.

**Pinned end to end.** The diff is computed in CRT.Data, serialised by the server, parsed by the
app and worded by the presenter - four places, no shared type - so
`A_field_level_diff_SURVIVES_THE_WIRE` runs a real edit through the whole chain.

#### A stale-cache bug, found because a flaky test was investigated rather than re-run

`PublishExecutorTests.Re_running_the_same_publish_is_safe` began failing intermittently. The cause
was real and was in `PublishedBoardReader`, added this session: it passed the workbook's **file
path** as `BoardDataReader`'s cache key.

That key is safe in CRT, which never rewrites a published board. **On the server it is exactly
wrong**, because publishing rewrites that very file - so after a merge, the next maintainer opening
a submission for that system would be served the PRE-PUBLISH board out of cache and shown a diff
against data that no longer exists, with nothing indicating anything was wrong.

Fixed with a fresh key per read, cleared in a `finally` so the key space cannot grow.
`PublishedBoardReaderTests.A_REWRITTEN_board_is_read_afresh_rather_than_from_cache` fails against
the old version.

**The flake was only partly explained by it** - one further failure occurred after the fix, and
has not reproduced in six runs of that assembly plus four full-solution runs. The remaining
suspicion is recorded in the test's own header: both publishes write the same workbook path in
quick succession, and a Windows handle from the preceding read may not be released when EPPlus
reopens it. **Unproven, and labelled as such.**

**Landed next: the ASSET ENDPOINT and IMAGES SIDE BY SIDE** (+68 tests; suite 4392 to 4460). This
is the piece the remaining three sub-parts of task 4 were all blocked on - the maintainer app could
compute a diff but could not fetch a single byte of image.

Two routes, `GET /api/review/submissions/{id}/submitted/{hash}` and
`.../published/{**path}`, deliberately **separate rather than one route with a "which side"
parameter**, because they take different input and are guarded differently. A single route would
have to branch on untrusted input to pick its own guard, which is how the wrong one eventually
gets applied.

- **The SUBMITTED side is addressed by HASH, and the risk is not traversal.** A 64-character
  lowercase hex string cannot carry one. The risk is CROSS-SUBMISSION READING: the blob store is
  shared and content-addressed, so without a scope check "may review submission 7" would mean "may
  read any blob anyone has ever uploaded" to anyone holding a hash - and hashes travel in
  manifests, which maintainers see. `ReviewAssetLocator.IsSubmittedBlobAllowed` requires the hash to
  be one the named submission's own manifest references.
- **The PUBLISHED side is addressed by a caller-supplied PATH and IS the disclosure boundary.** It
  goes through `SubmissionPathRules`, resolved against **this system's own folder** rather than
  against the data tree - so a maintainer opening a C64 submission cannot read an Amstrad board
  through it. The identity is resolved through the same rules first, since a manufacturer of ".."
  would relocate the very folder being contained to.
- **The content type is an ALLOWLIST of six image types, never inferred.** These bytes are
  contributor-supplied and served from the server's own origin, so a file served as `text/html`
  would run script there against a signed-in maintainer's session. Everything else is
  `octet-stream`, which downloads - also the right behaviour for a datasheet. **SVG is absent on
  purpose** despite being an image: it is scriptable XML. The extension is read from the LAST dot,
  so `evil.png.html` is not an image.
- **Refused, absent and out-of-scope are ONE 404.** Distinguishing them lets the tree be probed
  for what exists.

**The containment was proved by deleting it**, given this exact class of test has already been
caught vacuous twice this session. Removing `SubmissionPathRules` from the locator turns three
tests red; removing the manifest-scope check turns one red. Every traversal test points at a
**real file in a real sibling folder**, and carries an anti-vacuity half proving the same shape
inside the tree is still found.

**`ReviewImageComparison` (CRT.Maintainer, pure) decides WHICH images are shown.** Fetching bytes is
the easy half; picking the right pair is the half that can be wrong, and wrong silently.

- **The manifest's file list is the COMPLETE intended state, not a list of changes**, so an
  untouched board re-lists every image it has. Showing all of them would bury the one that moved.
- **A DELETION is found from the PUBLISHED side, never from the manifest** - there is no blob for
  a removed file, so anything driven off what was uploaded misses it entirely. Removals are listed
  **first**, the same rule `ReviewSummaryPresenter` follows.
- Added / Replaced / Removed are kept **distinct** rather than collapsed into "changed": a missing
  panel is captioned ("Not in the published board") rather than left blank, which would read as an
  image that failed to load.

#### The published file list CANNOT be derived on the client, and the first version tried to

The review window built the published side from the change summary's row keys. **A summary's keys
are NATURAL keys** - for a component image, `BoardLabel|Region|Pin|Name` - **not file paths**. So
nothing ever matched: every submitted image would have been reported as an addition, and a deleted
image would never have appeared at all.

Both failures are silent. The panel draws perfectly and describes something that did not happen -
the worst shape a defect can take on a screen whose entire purpose is telling a maintainer what a
submission does. The server now states the list outright in a `publishedFiles` field, which is
also where it belongs: the server is the only place the published board exists, the same reasoning
the change summary itself is computed there. `ReviewPublishedFilesTests` pins that the values are
paths and not keys.

**Landed next: the MOVED HIGHLIGHT drawn on the schematic** (+30 tests; suite 4460 to 4490). The
change a row diff genuinely cannot answer: `X: 100 -> 400` is the same information and tells nobody
whether the new position is *right*, which is the maintainer's only real question.

- **A move arrives as an ordinary CHANGED ROW.** A highlight's natural key is
  `SchematicName|BoardLabel`, so X/Y/Width/Height are compared FIELDS - the data was already
  flowing and what was missing was the geometry.
- **The placement is PROPORTIONAL, and that is the whole design.** `ReviewHighlightGeometry`
  (CRT.Maintainer, pure) turns stored pixel coordinates into FRACTIONS of the image; the control
  multiplies by whatever size it was given. No panel dimension appears in the geometry at all -
  the same split `ExportOverlayGeometry` uses, after the PDF version that confused fractions with
  absolute lengths drew a tenth of a board as most of it.
- **Clipped to the image, never clamped inward** - the same rule the PDF settled on, and for the
  same reason: clamping only the origin keeps the full width and slides the rectangle off the
  thing it marks. Proved by swapping it and watching three tests go red.
- **Invariant-culture parsing, verified by actually switching to `da-DK`** rather than by reading
  the parse call. This is the project's fourth encounter with that bug; `0.5` read
  culture-sensitively becomes `5`, a highlight ten times too far across, on one maintainer's
  machine only.
- **An ABSENT field reports the same value on BOTH sides.** A field diff lists only what changed,
  so a purely horizontal move carries no Y row. Reading a missing field as blank would put the
  "before" rectangle at the top of the board and **invent a vertical move that never happened** -
  and inventing a change is worse than missing one, because the maintainer acts on it.
- **Both rectangles go on ONE copy of the board**, red for where it was and green for where it is
  being put - the same colour vocabulary the change summary already uses. Two images side by side
  would make the maintainer hold one in their head; overlaying makes the distance itself the thing
  on screen.

`ReviewHighlightCanvas` is a custom control rather than a Canvas with children because the board
is letterboxed inside whatever space the panel gives it, so the rectangles must be placed against
the DRAWN image rather than the control - a rect only known during rendering. It owns its bitmap
and disposes it on detach; unlike CRT's Workbooks tab these are never shared, so the
dispose-under-an-open-editor hazard does not apply.

#### Two more fields the client cannot derive, for the same reason as `publishedFiles`

A highlight names a **schematic**, not a file, so drawing one needs that schematic's image. The
server now sends `schematicImages` alongside `publishedFiles`. **The SUBMITTED board wins where
both name a schematic** - a submission can repoint one at a new image, and drawing the moved
highlight on the old picture would judge the change against the wrong board. The published board
fills in only what the submission does not mention, so a highlight on an untouched schematic still
draws. Both halves are pinned, and reversing the precedence turns one red.

**A third field, `publishedHashes`, landed on 2026-09-23 after the first real submission.** A
manifest names every file the board references - it describes the whole intended state - so
paired against `publishedFiles` alone, every file was "on both sides" and drawn as replaced: a
one-line description edit on the C64 250407 showed as "1178 images to compare", all identical.
`ReviewImageComparison.Plan` had always dropped a matching pair when handed the published hash;
nothing had ever handed it one. The server now sends path -> SHA-256 for the files on both sides
(`PublishedFileHashes`), hashed from the bytes on disk **rather than read from
`dataChecksums.json`**: nothing in the service writes that manifest yet, so it goes stale on the
first publish, and a stale hash that happens to match HIDES a change - the one wrong direction.
Only the intersection is hashed (a removal and an addition need no hash to say so), and a
length-plus-mtime cache makes a refresh free. An older server that omits the field degrades to
the previous behaviour, pinned on the review side.

**Landed next: SCOPE BASELINES, which completes task 4** (+16 tests; suite 4490 to 4506).

**The sub-task reads "scope baselines plotted" and that turned out to be the wrong shape**, so it
is worth recording why rather than silently doing something else. A baseline is stored as a
captured IMAGE - CRT saves what the scope drew, not the samples behind it - so there is nothing to
re-plot at a common scale even in principle. The side-by-side comparison already shows both
waveforms.

**What the picture cannot show is that the SETTINGS moved underneath it.** A trace captured at
2 V/div beside one at 5 V/div is the same signal at a different scale and reads as a change in the
circuit. That is the commonest way a baseline goes wrong - somebody recaptures on a
differently-configured scope and does not mention it - and approving it replaces a good reference
with one that cannot be compared against anything.

- `ReviewScopeBaseline` (CRT.Maintainer, pure) reports which of T/DIV, V/DIV and T.LVL moved, in the
  order a scope's own controls read rather than the order the diff happens to list them.
- **A row whose settings did NOT change is not reported.** Component-image rows change for all
  sorts of ordinary reasons, and a warning that fires on everything is one a maintainer learns to
  ignore. Proved by removing the filter and watching that test go red.
- **The warning names the CONSEQUENCE, not just the fact**: "...so these two traces are not drawn
  at the same scale" is what a maintainer needs in order to judge the pictures below it.
- **Listed by component rather than pinned to its image panel, deliberately.** A component image's
  row key is `BoardLabel|Region|Pin|Name` and carries no file name, so matching a row to the
  picture it produced is not something the client can do RELIABLY - and a warning attached to the
  wrong trace is worse than one listed separately. It sits immediately above the pictures instead.
- Amber rather than red: a recapture at different settings is not necessarily wrong, it is
  something to take into account.

#### A presentation defect the field-level diff had been shipping since it landed

Natural keys join their parts with **U+241F**, which renders as a box or as nothing at all
depending on the font. The field diff had been drawing them raw, so a maintainer saw
`U8ASSY 250407 12Clock` rather than four distinct fields. `ReviewScopeBaseline.DescribeRowKey`
spells them out with a visible separator and drops empty parts - a component image with no region
would otherwise read as `U8 /  / 12`, which looks like missing data rather than a field that does
not apply. Both the scope section and the field diff now use it.

**Task 4 is now complete.**

**Landed next: task 5's DECISIONS - two of the three outcomes** (+33 tests; suite 4506 to 4539).
`ReviewDecisionRules` (CRT.Server, pure) owns whether a decision may be made at all, and
`POST .../reject` and `POST .../request-changes` are mapped.

- **A REVIEWER (the old recommend-only role, retired 2026-09-25) may reject and return but NEVER
  approve, and that is enforced here.** Phase 6's role table gave Reviewer a blast radius of "none -
  no published data can change", which held only because both outcomes a reviewer could reach
  leave the tree untouched. Proved by widening
  `CanApprove` to `CanReview` and watching that test go red.
- **Rejecting and returning need the SAME authority, deliberately.** If returning needed a higher
  role than rejecting, the cheap outcome would be the harder one to reach and maintainers would
  reject things that could have been a conversation - the exact failure task 5 warns about.
- **The double-decision interlock.** Two maintainers with the queue open both press a button, or one
  presses twice over a slow link. The state is re-read and re-checked SERVER-SIDE on every request,
  because the app's belief that a submission is still pending can be seconds out of date. Removing
  that check turns nine tests red.
- **`approved` is still actionable**, which is the one non-obvious case: it means a maintainer
  accepted the submission but publishing did not happen or did not finish, and refusing it would
  strand something already agreed to. `merged` is terminal.
- **A refusal on STATE is 409, not 403.** The account may well be allowed to review - what is
  wrong is that somebody else decided first. A 403 would send the maintainer looking at their own
  permissions for a conflict that is about timing.
- **A reason is REQUIRED for both outcomes**, with a deliberately low floor (10 characters) that
  rejects thoughtlessness rather than brevity. Contributing needs no account, so the contact email
  and this comment are the ENTIRE feedback channel; a rejection with no reason is
  indistinguishable from being ignored.

**`ISubmissionStore.SetDecisionAsync` is new, and it closes a gap that had been open since
0001_initial.sql.** `submissions.decided_by` and `decision_comment` have existed from the start
with **nothing ever writing them** - `SetStateAsync` sets neither - so the audit trail could say a
submission was rejected and could not say by whom or why. The four columns were cross-checked
against the real schema one by one, and the eight state values against the CHECK constraint (which
migration 0002 widened and 0004 restates): the rule that has caught four schema-versus-code bugs in
this project. The fake records what was written so a test cannot pass while the comment is dropped.

#### APPROVE IS DELIBERATELY NOT MAPPED, and this is the important part of this session

Approving publishes, and **the publish path does not write the JSON sidecar** - see the warning
added to task 6 above. Wiring Approve today would let a maintainer press a button that silently drops
every component highlight and every KiCad calibration from a board, with no revision to restore
from. The two outcomes that touch no published data shipped; the one that does is held back until
the writer is complete. `ReviewDecisionRules.CanApprove` exists and is tested, so the rule is ready
and only the route is missing.

**Landed next: `BoardSidecarWriter`, which closes the data-loss gap above** (+20 tests in CRT.Data,
+6 in CRT.Server; suite 4539 to 4565). Task 6's writer is now complete: a publish writes the
workbook AND the sidecar, and the content hash accounts for both. See the task 6 entry for the
detail. **Approve's route is no longer blocked by data loss** - the remaining piece before it can
be mapped is the caller that assembles the merged `BoardData` and the submission's calibrations.

#### The flaky publish test: solved, and my earlier diagnosis was WRONG

`PublishExecutorTests.Re_running_the_same_publish_is_safe` had been failing roughly one run in
five. The note left in its header on 2026-09-21 blamed a Windows file handle not being released
when EPPlus reopened the workbook. **That was wrong**, and the evidence was already there: the test
never failed when run in ISOLATION, however many times - which a handle race inside one test cannot
explain.

The real cause: **xunit gives every class WITHOUT a `[Collection]` its own collection and runs
those in parallel**, so `parallelizeTestCollections: false` did not serialise them - it only stops
collections interleaving once they exist. Three classes in `CRT.Server.Tests`
(`PublishExecutorTests`, `PublishedBoardReaderTests`, `ReviewAssetLocatorTests`) were driving
EPPlus and `BoardDataReader`'s static cache against real files simultaneously.

`CRT.Data.Tests` already had this right - its `"BoardData"` collection exists for exactly this -
and CLAUDE.md states the rule outright. The three classes now share a `"BoardFiles"` collection.
Six consecutive full-assembly runs green, from ~1-in-5 before. The header note has been corrected
rather than left to mislead the next reader.

**Landed next: APPROVE, end to end - task 5 and task 6 are now COMPLETE** (+58 tests; suite 4565
to 4623). `POST /api/review/submissions/{id}/approve` is mapped, and the maintainer app has all three
decision buttons.

- **`PublishMerge` (CRT.Data) assembles the board**, and the first thing to understand is that
  "merge" is the wrong word: the manifest is BY CONTRACT the complete intended state, so the
  submitted rows REPLACE the published ones. Anything that combined them would resurrect every row
  a contributor deliberately deleted, on every publish, with nothing reporting it.
- **`ApprovePublishFlow` (CRT.Server) sequences the whole chain**, refusing in a deliberate order -
  authority, existence, state, payload, plan, and only then writing. Every refusal before the last
  step leaves the tree and the submission exactly as they were, so a refused approval is simply
  retried.
- **A MISSING PAYLOAD IS REFUSED, and that guard is the most important line in the class.** No
  payload means no rows, and a board built from no rows is EMPTY - publishing it would delete the
  system's entire contents over a board that was working, irreversibly, while looking like an
  ordinary successful publish.
- **The published revision IS the board's own revision date.** Not a free choice: `DraftBaseRevision`
  stamps a draft's `BaseRevision` from the official board's revision date and submissions send it
  back, so recording a counter or a timestamp would make every contributor re-base against a value
  their drafts never carry and the drift warning would fire on every board forever.
- **Approve is last on the button row and separated from the other two**, because it is the only
  irreversible action on the screen. For an account that may not publish it is DISABLED rather
  than hidden, with a tooltip saying to ask an administrator - a missing button reads as a broken
  screen.
- **A decision timeout says the decision may already have been recorded.** The obvious response to
  "that failed" is to press the button again, and on approve that is a second publish.

**`SubmissionRows.RevisionDate` was added** (optional, so no format-version bump - an older client
omits it and the publish keeps the published date). It closes the gap recorded earlier: a
revision-date change was both invisible to the maintainer and unpublishable. `ReviewEndpoints` now
compares it properly instead of using the published date for both sides.

#### A NEW SYSTEM was publishing into the FROZEN generation

Found by `ApprovePublishFlowTests` failing on the path the board landed at. `PublishPlan` resolves
the generation from the files already in the system's folder - and **a new system's folder is
empty**, so it resolved to null, and null means "no version suffix". Every new system would have
published as `Data C64 250407.xlsx`: the frozen file serving every pre-2.0.0 build, the one file
the project owner's rule says is never written. Older builds would have received contributed data
they cannot read.

The generation now comes from the TREE's master workbooks when the system's own folder is empty,
and publishing is REFUSED outright when the tree carries no versioned master at all.

**The first version of that guard was itself wrong, and an existing test caught it.**
`ResolveNewestGeneration` answers null for two opposite situations - an empty folder (a new system)
and a folder holding only an unversioned board (an existing system that legitimately lives there).
Treating them alike made every unversioned board unpublishable. The folder's EMPTINESS is what
separates them, and both halves are now pinned.

`PublishPlanTests.A_brand_new_system_with_an_empty_folder_publishes_unversioned` asserted the bug
as intended behaviour and has been rewritten, with the reasoning recorded in the test itself.

#### A contributor's KiCad CALIBRATIONS now travel - the chain was broken in three places

`SubmissionRows.KiCadCalibrations` existed from the start and the publish path handled it
correctly, but **nothing ever filled it**. A contributor who spent an evening calibrating a board's
KiCad overlay submitted their work and the calibration simply did not arrive - no error, no
warning, just absent from the published board. Closed 2026-09-22 (+21 tests; suite 4623 to 4644).

Three separate things were missing, and fixing only one would have been worse than fixing none:

1. **Nothing collected them from the draft.** `KiCadCalibrationDraftWriter.CollectForSubmission`
   does, and **excludes rows in the `Deleted` state** - a draft records a removal as a row rather
   than by dropping it, so reading every row would have sent back the very calibration the
   contributor removed, and the manifest being the complete intended state would have republished
   it.
2. **`SubmissionManifestBuilder` could not carry them.** It takes a `BoardData`, which has no
   calibration section. They now arrive as an OPTIONAL TRAILING parameter alongside the board, so
   every existing caller kept working unchanged - which is what made this safe to do at all.
3. **`BoardDraftNaturalKeys.ForRow` THREW on a `KiCadCalibrationEntry`.** Nothing had ever called
   it with one. Found by the review summary's own test rather than by reading.

**They are also REVIEWED now, and that mattered more than the wiring.** `ReviewSummary` compares
two `BoardData`, so without a deliberate eleventh section the calibrations would have reached the
published tree **with no maintainer ever having seen them** - creating the exact failure this phase
exists to prevent. A calibration is the offset/scale/mirror box mapping a schematic onto the KiCad
board's coordinates: getting one wrong breaks nothing visibly, the trace overlay just lands in the
wrong place, so it is precisely the change a maintainer must be TOLD about.

`ReviewSummaryTests.Every_section_of_a_board_is_compared` - which exists to catch a section whose
changes are invisible - now covers all eleven.

**The numbers format INVARIANT.** Calibrations are the first row type whose fields are doubles
rather than strings, so they are formatted for display; a culture-sensitive format would render
`0.75` as `0,75` on a Danish server and report a change on every field of every calibration,
forever, against a board nobody had touched. Fifth encounter with that class of bug here.

**Landed next: task 8's RELEASE WORKFLOW** -
[build-and-release-maintainer.yml](../.github/workflows/build-and-release-maintainer.yml), separate from
CRT's exactly as the task requires.

- **Separate because the AUDIENCES barely overlap.** CRT ships to every hobbyist using the app;
  this ships to a handful of maintainers. One workflow would push a maintainer-app fix at thousands of
  people who cannot sign in to it, and hold a maintainer-app fix behind whatever CRT is mid-way
  through. The versions are independent too - 0.1.0-alpha.1 against CRT's 2.6.x - and a shared
  workflow reads ONE `InformationalVersion`, so one product would release under a version that
  means nothing for it.
- **The tag and packId are PREFIXED** (`maintainer-v0.1.0-alpha.1`, `Classic-Repair-Toolbox-Maintainer`).
  The two products share one repository, so one tag namespace and one Releases page; an unprefixed
  tag would sit beside CRT's with nothing saying which product it belongs to. **Corrected
  2026-09-25: the packId does NOT keep the two update feeds apart.** Velopack's GithubSource
  (1.2.158, decompiled) merges the `releases.<channel>.json` of every recent release in the repo
  and picks the highest version without looking at the package id. What separates them is the
  CHANNEL: the maintainer app packs on `maintainer-win` / `maintainer-linux`, whose feed files CRT never asks
  for - which also protects CRT versions already installed. **Then moved out entirely (owner
  decision, same day): review releases publish to their own repository,
  `HovKlan-DH/Classic-Repair-Toolbox-Maintainer`**, so CRT's Releases page lists only CRT and CRT's
  update check never sees them; the channels stay as a second guard. Needs the
  `REVIEW_RELEASES_TOKEN` secret. `MaintainerReleaseSeparationTests` (CRT.App.Tests) pins both.
  The same day the maintainer app got `VelopackApp.Build().Run()` (install hooks only, no updater):
  without it `vpk pack` refused the package, so no review release had ever been built.
- **The release body is GENERATED, never read from `CHANGELOG.md`.** That file is CRT's, written
  by hand, and is the body of every CRT release - putting it on a maintainer-app release would
  describe changes that are not in it, and sharing it would mean editing a file that is explicitly
  off-limits. The body says the version and points at the commit history, which is all it can say
  truthfully. If the maintainer app ever wants real notes it gets its own changelog; it must never
  borrow CRT's.
- **It tests the WHOLE SOLUTION, not just `CRT.Maintainer.Tests`.** The maintainer app is built on
  CRT.Data, which CRT and the server also depend on, so a change breaking either of them breaks
  this release's foundation. Testing only its own assembly would ship an app built on a library
  whose tests are red.
- **No data pack.** CRT ships `Assets/Data` so a fresh install has boards before its first sync;
  the maintainer app reads everything from the server, so copying it would add tens of megabytes a
  maintainer never uses.

Verified by actually running the pieces rather than by reading: `InformationalVersion` reads as
`0.1.0-alpha.1`, and both `win-x64` and `linux-x64` publish self-contained and produce executables
whose names match `EXE_WINDOWS`/`EXE_UNIX` exactly - a mismatch there would fail at the signing
step, on a self-hosted runner, at release time.

**It reuses `Assets/CRT_icon.ico`**, because that is the only icon in the repository. A distinct
icon for the maintainer app is a design decision for the project owner rather than something to invent;
until then the two installers share one.

**Landed next: the maintainer's COMMENT now reaches the contributor** (+3 tests; suite 4644 to
4646). `GET /api/submissions/{id}`'s `maintainerComment` field had been returning the empty string
with a comment saying "always present, always empty for now... filling it in later is a server
change alone, with nothing to update on contributors' machines". That prediction held exactly: the
decisions landed, and this needed only to stop returning empty.

- `SubmissionRecord` gained `DecisionComment` as a TRAILING OPTIONAL field, so every existing
  construction kept working.
- The column is read LAST in all three SELECTs. The row mapper reads by ORDINAL, so a column
  inserted in the middle would silently shift every field after it - which compiles and returns
  the wrong data.
- The fake records the comment on the ROW as well as in its audit dictionary, so a test reads it
  back the way a contributor's own status request does rather than through a back door.
- An UNDECIDED submission carries NO comment, so a contributor's screen can tell "not looked at
  yet" from "looked at and returned without a word" - the second would be a defect worth noticing.

**Still to do in Phase 5:** the rest of the contributor-side "request changes" round trip. The
server now tells CRT that a submission was returned and why; what CRT does not yet do is turn that
into an EDITABLE DRAFT the contributor can pick up - the draft still sits where they left it, with
no link back to the review. The definition of done says "'Request changes' round-trips back into
the contributor's drafts", so that link is the remaining work.

**Also still open, and NOT a code task:** the definition of done requires a real submission
reviewed and merged end to end into BETA, and the merged board opening correctly in CRT after a
sync. Nothing in this phase has been exercised against the live server or a real MariaDB - every
test runs against an in-memory fake and a temp folder. That is a deployment step for the
project owner, and the five schema-versus-code disagreements already found by hand-checking columns
are the reason it matters.

#### Sign-in screen reworked after owner feedback (2026-09-22)

Three things, from first launch of the built app:

- **The server prompt is GONE.** It asked for an address on every launch - a question whose answer
  never changes, and which breaks the moment somebody types a trailing slash or omits the scheme.
  `ReviewApiRoutes.DefaultBaseAddress` is the constant now.
- **"I forgot my password" was missing entirely.** Both endpoints have existed since Phase 3
  (`/api/accounts/forgot-password` and `/reset-password`); only the UI was absent, so a maintainer
  locked out of the one tool that can publish had no way back in. The server **always answers 202**
  whether or not the address is known, and the client deliberately does not try to be more helpful
  - distinguishing the two would turn the sign-in screen into a way of discovering which addresses
  hold maintainer accounts.
- **A banner says what the app IS**, above the fields rather than below them: for hardware and
  board maintainers only, needs an account, and ordinary users want the normal CRT application.
  Somebody who installed this expecting CRT previously met a sign-in screen for an account they
  cannot create.

#### A doubled "/api" the route test caught before first launch

The default address was first written as `https://classic-repair-toolbox.dk/api`, copying CRT's own
`AppConfig.CrtServerBaseUrl`. **That constant is not interchangeable**: CRT's `SubmissionClient`
appends only `/submissions` to it, while every `ReviewApiRoutes` method appends the full path. The
copied value produced `/api/api/review/queue` - a 404 on the very first request, from an address
that reads perfectly correctly in the source.

Verified against the real deployment rather than by reading: Apache proxies `/api/` through
(DEPLOYMENT.md step for the vhost) and the endpoints map `/api/accounts`, `/api/review`,
`/api/submissions`. The constant carries no suffix, and a test now asserts both the exact URLs and
that the constant does not end in `/api`.

#### The ONE step with no code path: granting the first administrator

`accounts.is_administrator` defaults to 0, **no endpoint sets it, and nothing seeds it**. So a
freshly deployed server has no account that can approve anything - the queue answers 403 for
everyone and the maintainer app shows "this account is not allowed to review submissions".

That is deliberate rather than an omission: an endpoint that grants administrator is an endpoint
that can be abused to grant administrator. The first one is made by hand in MariaDB, by somebody
who already has database access (confirmed with the project owner 2026-09-22 - no UI wanted).

[DEPLOYMENT.md](../src/CRT.Server/DEPLOYMENT.md) now carries the runbook: register through the
normal flow so the password is hashed by the service, verify the address, then
`UPDATE accounts SET is_administrator = 1 WHERE email_normalised = ...`. **Matching on
`email_normalised`, not `email`** - the stored value is lowercased and trimmed, so matching the
raw column silently affects zero rows for any address typed with a capital letter. No restart is
needed; authority is resolved per request, which is the same property that makes locking bite
immediately.

It also carries an ordered end-to-end check (log in, queue, submit from CRT, open, request
changes, approve LAST because it cannot be undone) with what each failure code means.

#### The "flaky" publish test was a REAL BUG: workbooks were not deterministic (2026-09-22)

`PublishExecutorTests.Re_running_the_same_publish_is_safe` failed again on a Stop-hook run, having
been declared fixed twice. **It was never flaky. The production code was wrong.**

**An `.xlsx` is a ZIP, and EPPlus stamps every entry with the current clock.** So two publishes of
a byte-identical board, seconds apart, produced different files and therefore different SHA-256
hashes. Proved rather than inferred: a scratch harness wrote the same `BoardData` twice with a
1.1s gap and the hashes differed; unzipping showed identical content and entry timestamps two
seconds apart.

**Why that matters far beyond the test.** The workbook's hash is folded into the system's
`ContentHash` (`PublishPlan.DescriptorWithWorkbook`), which is what every CRT client uses to decide
whether to re-download a board:

- **Re-publishing an unchanged system moved its content hash**, so every user on every machine
  re-downloaded a board that had not changed. `SystemDescriptorRules` states the intended guarantee
  in its own comments - "republishing an unchanged system does not produce a file that differs" -
  and the ZIP timestamp silently broke it.
- **Re-running an interrupted publish**, the documented recovery, produced a different result from
  the one it was recovering. The opposite of idempotent.

**The fix:** `BoardWorkbookWriter.MakeDeterministic` rewrites the archive with a fixed
`2000-01-01` timestamp on every entry, copying entries verbatim (no re-compression) and preserving
ORDER - `[Content_Types].xml` must come first, so sorting would produce a file that hashes
consistently and might not open. Failure is soft: an unnormalised workbook is still valid, it just
hashes differently next time. **The constant must never change** - moving it would re-hash every
published workbook at once and trigger a full-tree download for every user.

**Two earlier diagnoses were wrong, and the header in `PublishExecutorTests` now says so.** First a
Windows file handle; then a cross-class xunit race, which added `[Collection("BoardFiles")]` on the
reasoning that "passes alone, fails in a full run" must mean shared state. That inference is
seductive and was false - a full run is simply SLOWER, which is what made it cross a second
boundary. The collection attribute is kept (the EPPlus/static-cache reason is real), but the note
claiming it fixed this test is replaced by what actually happened.

Anti-vacuity: three new tests in `BoardWorkbookWriterTests`, and disabling `MakeDeterministic`
turns the determinism one red. **The `Thread.Sleep(1100)` in it is load-bearing** - two writes
inside one clock second agree even with the bug present, which is exactly why this surfaced as an
intermittent failure somewhere else rather than a reproducible one here. A third test pins that a
CHANGED board still produces different bytes, so a "fix" making every workbook identical could not
pass.

(+3 tests; suite 4781 to 4784.)

#### "My submissions" made legible, and a RAW DATABASE VALUE that reached the user (2026-09-22)

Reported from the screen after the first real review round trip. Three things, one of them a bug.

**The bug: `Reported as [changes_requested]`.** The four Phase 5 review states were added to the
server's vocabulary when the maintainer application was built and **never taught to
`SubmissionReceiptPresenter.DescribeState`**, so they fell through to its unknown-state fallback -
a raw database value, underscore and all, in the one place that class exists to prevent exactly
that. `changes_requested` is the state where somebody is being ASKED TO DO SOMETHING, and it read
as a fault in the application.

**The same omission had a worse, invisible half.** `IsStillOpen` did not know `merged` or
`withdrawn` either, so **every merged submission was re-checked on every refresh - and, once the
launch check landed, on every launch - forever.** `changes_requested` and `approved` are
deliberately still open: neither is the end, and treating either as decided would freeze the row.

**The layout.** Each submission is now its own card: a faint fill, a 1px border, and a **4px
status-coloured left edge** (owner's choice over a plain grey card or a fully tinted one). A
single bottom rule per row made five submissions read as one continuous block, and the 1px
`#EEEEEE` outline that replaced it was invisible against a white window.

- **The colour comes off the SAME switch as the words.** `SubmissionReceiptPresenter.ClassifyState`
  returns one of four `SubmissionOutcomeKind` values, so a card can never be painted green while
  its own text reads "Changes requested". Deciding colours in the UI is how those two drift.
- **`NeedsAction` and `Bad` share the red deliberately** - both mean "this did not go through", and
  which one it is, is already written in words beside it. They stay distinct in the enum so a later
  "things waiting on me" count can tell them apart.
- **`uploading` is classified Bad, not Waiting**, matching its own wording: nothing is uploading any
  more, and colouring it neutral would leave a contributor waiting for a transfer that stopped days
  ago.
- **Two new theme keys, `Card_Bg` and `Text_Waiting_Fg`, defined in BOTH themes.** They step in
  OPPOSITE directions - darker than the background in light, lighter in dark - which is precisely
  what a later "simplification" to one shared grey would break, in whichever theme the person
  making the change is not running. Pinned by `SubmissionCardPaletteTests`, and sabotaging
  `Card_Bg` back to White turns two of them red.

**The dates.** `SubmissionReceipt` gained `DecidedUtc` (the server already sent `decidedUtc`;
nothing stored it), shown as **"Replied 22 September 2026" inside the maintainer's own panel** rather
than in the row of dates above - a date about the comment belongs with the comment. "Last checked"
is dimmed, being this computer's bookkeeping rather than anything about the submission.

Caught while wiring it: **`AcknowledgeComment` rebuilds the whole record**, so a field left out is
a field erased - "Mark as read" would have silently dropped the reply date. Both it and
`UpdateState` now carry it, the latter keeping an existing value when a caller passes none.

**Still not possible: renaming a system's manufacturer/hardware/board.** `SubmissionRename` exists
but is section-scoped (rows WITHIN a board). The identity triple is fixed at creation and becomes
the folder path, the `systems` row and the `SystemId` every receipt references. Today the answer is
to discard the draft and recreate it; doing it properly is Phase 6-sized.

(+33 tests; suite 4748 to 4781.)

#### The status check only ever ran from a BUTTON (2026-09-22)

Immediately after the badge landed: "I launched CRT and I think I needed to press Check for
updates for the new comment to show up? This should of course load at application launch."

Correct, and the badge made it worse rather than caused it. `SubmissionReceiptStore.Load()` runs at
startup, but that only reads the LOCAL cache - **the only call to `GetStatusAsync` in the whole app
was inside `MySubmissionsWindow`'s Refresh button.** So a maintainer could request changes and the
contributor would launch to a stale cache, no badge and no comment, until they pressed a button
they had no reason to think was necessary.

`Handlers/Online/SubmissionStatusRefresh.cs` (new, pure but for the injected lookup) now runs from
`Main.StartAsync`:

- **Fire-and-forget, and it swallows everything.** Startup timing is a logged `StartupTimeline`
  milestone and must not include a server round trip, and an exception escaping an un-awaited Task
  is reported arbitrarily late by `TaskScheduler.UnobservedTaskException` or never - the same
  reasoning `Main.StartAsync`'s own header already gives for catching there.
- **No request at all when nothing is still open**, so a user who has never contributed pays
  nothing on launch.
- **It returns how many receipts CHANGED, not how many were checked.** `UpdateState` rewrites and
  saves unconditionally, so counting calls would report "something happened" every launch and
  redraw the tab for nothing.
- **A comment that moved counts even when the state did not** - a maintainer can reword or add a
  note without the state changing, and comparing only the state would leave that with no badge.
- **One unreachable row does not abandon the rest.**

The window keeps its own loop deliberately: it reports "checked N, updated M, could not reach K",
and the shared method returns only the changed count. The SKIP rule
(`SubmissionReceiptPresenter.IsStillOpen`) is shared and commented in both places.

(+11 tests; suite 4730 to 4742.)

#### ONE email address for the whole application (2026-09-22)

Project owner: "I think I ask for the email already in Feedback... log one email address for the user,
and use that everywhere it asks for an email address."

**The setting already existed.** `UserSettings.ContactEmail` has been written by the Feedback tab
since long before submissions did, and `SubmitDraftWindow` simply never read it - so a contributor
who had already given their address was asked again, with no sign the app knew it. Grep confirms
these are the only two email inputs in the app; the Contribute tab has none.

- **Prefilled, not locked.** A contributor may legitimately want a different address on a
  contribution than on a bug report, and this is the one a person actually acts on.
- **Not overwritten if the box already holds something**, so a caller that pre-populated the dialog
  is not silently overridden.
- **Saved back on SEND, not per keystroke**, so an address abandoned by closing the dialog is never
  persisted - and only when `EmailAddressRules.IsPlausible`, since storing "dennis@" would poison
  the prefill everywhere else.

**A test isolation trap worth remembering:** `UserSettings.LoadFrom` RETURNS EARLY when the file
does not exist, leaving the static `_data` exactly as the previous test left it. Pointing at a
fresh temp path therefore does NOT reset a value - the new test passed alone and failed in the full
run as soon as another test had saved an address. Clearing the field explicitly is what isolates;
the fresh path only stops tests overwriting each other's files.

(+6 tests; suite 4742 to 4748.)

#### Maintainer feedback is now VISIBLE: an unread badge on the Drafts tab (2026-09-22)

Owner question after the first real review: "I asked for changes - where can the user see
this?" It was already displayed, in `MySubmissionsWindow`, but only for somebody who thought to
press "My submissions". **A comment nobody notices is a comment nobody reads**, and contributing
needs no account - so that window is the entire channel back to the person who did the work.

- **`SubmissionReceipt.AcknowledgedComment` stores the TEXT that was read, not a boolean.** A
  "seen" flag has to be cleared by whoever writes a new comment, and the one time that is
  forgotten, a SECOND round of feedback arrives already marked read and is never seen at all.
  Comparing the acknowledged text against the current one makes a changed comment unread **by
  construction** - `SubmissionReceiptPresenter.HasUnreadComment` derives it, nothing maintains it.
- **`UpdateState` carries the acknowledgement through unchanged.** Every refresh re-reads the same
  comment from the server, so resetting it there would resurrect the badge on every check - and a
  badge that reappears for no visible reason is one the contributor learns to ignore. Sabotaging
  this line turns exactly one test red.
- **Explicit dismiss ("Mark as read"), not "cleared by opening the window"** (owner's
  choice): opening the list to check on a different submission must not silently mark feedback
  read that was never looked at. The comment **stays on screen** afterwards - only the "New" flag
  and the IndianRed outline go, since the contributor is about to act on it.
- **The Drafts tab now stays VISIBLE while feedback is unread**, which was a real hole: the tab
  hides itself when there are no local drafts, and discarding a draft after submitting it is the
  ordinary thing to do. Without this the badge would have been unreachable in precisely the case
  it exists for. The reverse is deliberately not true - no drafts and no feedback still hides the
  tab, so somebody who has never contributed sees no trace of the feature.
- **The empty state changes wording** when feedback is what is holding the tab open. "No local
  drafts yet" is a true sentence that answers the wrong question.

Still not done, and unchanged by this: the comment is **read-only feedback**. Acting on it means
editing the original draft and submitting again - "request changes" does not reopen the submission
as an editable draft. That remains the last Phase 5 code task.

(+17 tests; suite 4713 to 4730.)

#### The password-reset link had ALWAYS been a 404 (found in live use, 2026-09-22)

The first real reset mail was clicked and gave "This page can't be found". Not a deployment fault
- **the link had never worked, since Phase 3.**

`AccountFlows.BuildLink(options, "reset", token)` produced
`https://<host>/api/accounts/reset?token=...`. `AccountEndpoints` maps **nothing** at that path.
It maps `MapPost("/reset-password")`, which a browser click cannot reach - and cannot be made to,
because completing a reset needs the new PASSWORD, which a URL does not carry. The neighbouring
`/verify` link works precisely because it is `MapGet` and a click is all it needs.

**Why no test caught it:** `EmailTemplatesTests` asserted the reset URL was *present in the body*,
with the URL itself as a test constant. It pinned the bug in place. A test that a mail contains a
link is not a test that the link resolves.

**The fix (owner's choice): a CODE to paste, not a page to serve.** The alternative was an
HTML form served from the API; rejected because the audience is a handful of maintainers already
sitting in front of the maintainer app, and it would mean the API growing a page to style, escape and
keep accessible for one form.

It is also the **safer** shape, which was not the reason for choosing it but matters: the old link
was a GET, so a prefetching mail client or scanner following it would redeem a single-use token
before the person ever read the mail. A code pasted into an application cannot be burned that way.

- Both mails that offer a reset (`PasswordReset` and `AlreadyRegistered`) now print a bare code
  and name where to paste it.
- `CRT.Maintainer` gained the other half of the flow: a reset panel (code + new password) revealed by
  "I forgot my password", `ReviewApiClient.ResetPasswordAsync`, and `ReviewApiRoutes.ResetPassword`.
  **The panel opens on any success**, since the server answers a neutral 202 for unknown addresses
  too - a panel that appeared only for registered addresses would be the enumeration oracle that
  202 exists to prevent.
- `ParseFirstError` was added for a third refusal shape: a rejected password answers
  `{"errors":[...]}`, not `{"message":...}`, so a client reading only `message` showed a generic
  failure in the one case where the server had something actionable to say. All errors are joined,
  not just the first - fixing one rule at a time is a loop with no visible end.
- **On success the panel closes and the sign-in password box is filled in** rather than the queue
  opening. Resetting is not signing in; the server issues no session here, and pretending
  otherwise would mean logging in behind the person's back.
- `BuildLink`'s header now states that **only "verify" may be passed**, because a mail link is
  followed by a GET and only that route is mapped with `MapGet`.

Anti-vacuity: restoring the old URL wording turns `The_reset_mail_contains_NO_CLICKABLE_LINK_at_all`
red. It asserts the ABSENCE of any `http`/`token=` rather than the code's presence, so a later
"or click here" helpfully added back would be caught.

**`FakeEmailSender.ExtractTokenFromLastMail` now handles both shapes** - a `?token=` link for
verification, a bare code line for a reset - matching the code by SHAPE (long, url-safe charset)
rather than by the sentence above it, so rewording a mail cannot break every reset test.

(+13 tests; suite 4700 to 4713.)

#### Staying signed in: SLIDING EXPIRY, not a longer token (2026-09-22)

Owner request: "the password is complex and not something you want to deal with for a
low-volume thing like this" - sign in once, and forget it only on sign-out, lock or invalidation.
**The interesting part is the option that was rejected**, because the obvious two are both wrong.

- **Rejected: a 365-day token.** Costs nothing to implement (`RefreshTokenDays` has no upper
  bound), but keeps a stolen token alive for a year even if nobody ever touches it.
- **Rejected: calling `POST /api/accounts/refresh` when the token runs low.** This is the
  dangerous one, and it looks like the intended design. Refresh **rotates**: the server marks the
  old session `replaced_by_id` and issues a new token. A desktop app must then write that new
  value to disk, and a crash, power cut or failed write inside that window leaves it holding a
  spent token - whose next use hits reuse detection, **revokes every session for the account**,
  and writes a `session.reuse_detected` audit entry that reads as an attack. The user resets a
  password because their laptop lost power.
- **Built: extension without rotation.** `SessionExtensionRules` + `IAccountStore.ExtendSessionAsync`
  push `expires_utc` forward on the row already in use. The token never changes, so there is no
  disk write to lose and nothing to poison. A failure to extend costs nothing - the session keeps
  its old expiry and the next request tries again.

**A sliding window is SHORTER than the long fixed one it replaces**, which is the part worth
keeping: 30 days from LAST USE rather than from login, so a weekly user is never asked again while
an abandoned or stolen machine goes cold on its own.

Load-bearing details:

- **Extension runs only AFTER every authentication check has passed** in `AccountFlows.AuthenticateAsync`.
  A revoked, rotated or expired session, or one whose account has since been locked, has already
  returned null - so extension can never prolong a credential that should have stopped working.
  Moving it earlier would break "locking takes effect immediately", which Phase 6 requires.
- **The store repeats those conditions in its own `WHERE` clause** rather than trusting the
  caller, so a TOCTOU gap cannot grant anything. `FakeAccountStore` mirrors the same guards
  deliberately - a permissive fake would certify a caller the real database silently refuses.
- **A half-lifetime threshold**, or every request on a read-only screen writes a row.
- **The new expiry is measured from `now`, never from the old expiry**, which would compound until
  a frequently-used session outlived any intended policy. Pinned by an invariant test over
  repeated extension, since a single call cannot tell the two rules apart.
- `last_used_utc` finally means something. Until now only rotation wrote it, so a session used
  daily for a month still read as never used.

**Client side:** `ReviewSessionStore` persists the session, with the token **DPAPI-encrypted to
the logged-in Windows user** - a copied file is inert on another machine or account. **Off Windows
nothing is stored at all**, deliberately: a plaintext fallback would look protected and not be,
which is worse than an honest refusal, so those maintainers sign in each launch as before.
`ShowSignInPanel` is the single funnel that forgets it, reached from expiry, from any 401 (which
is what a lock, a revocation and a sign-out all look like from the client) and from the new **Sign
out** button - which revokes server-side first, while the token is still in hand.

**`ReviewSession`'s header used to forbid all of this** ("never written to disk... do not add
remember-me without deciding where that file lives and who can read it"). Both questions are now
answered in that header rather than the rule being quietly broken. Its one surviving prohibition:
**this app must never call `/refresh`**.

Anti-vacuity: deleting the `expiresUtc <= now` guard turns three tests red, including the
`DateTimeOffset.MinValue` overflow case. **The plaintext test needed fixing after sabotage proved
it vacuous** - it searched the raw file text, which base64 of the token passes trivially, so it now
decodes `protectedToken` before asserting. That version fails against a plaintext fallback.

(+29 tests; suite 4646 to 4700.)

**`PublishPlanDetail.DescriptorWithWorkbook(hash)` is the executor's last step before writing
`system.json`, and it is NOT optional.** The planned `Descriptor` covers the uploaded files only,
because the generated workbook's hash is not known until it has been written. A rows-only change
uploads nothing, so its file list is byte-identical to the previous publish's - without the
workbook folded in, the content hash does not move and **no client ever re-downloads the board
that just changed**.

This started life as a sentence in this document telling the executor's author to remember it,
which is the same shape as every instruction that eventually gets forgotten. It is now a method
that refuses a blank or malformed hash rather than silently hashing without the workbook, with
tests covering the rows-only case specifically.

**Not yet wired anywhere.** `BoardWorkbookWriter` has no call site outside its tests - the same
state `SystemDescriptorStore.Write` is in, and for the same reason: publishing is the step that
calls both, and it does not exist yet. Do not wire either into a request path to "finish" them.

### Tasks

1. Create `src/CRT.Maintainer/`, Avalonia, referencing `CRT.Data`. Follow CRT's own conventions: no
   MVVM, code-behind, partial classes by area with header comments, pure logic in `Handlers/`.
   **[DONE 2026-09-21]** - plus `tests/CRT.Maintainer.Tests/`, both in the solution.
2. Log in against `/api/`; list the queue. **[DONE 2026-09-21]** - `ReviewEndpoints` server-side,
   `ReviewApiRoutes`/`ReviewApiParser`/`ReviewApiClient`/`ReviewSession` client-side, and the
   window's sign-in panel and live queue. **Not yet exercised against the real server** - see the
   note in DEPLOYMENT.md's "What is deliberately not here yet".
3. **Open on a summary**, never on a whole board: "3 components changed, 1 added, 2 images added, 1
   highlight moved". Drill in from there. **[DONE 2026-09-21]** - `ReviewSummary` in CRT.Data
   computes it, `ReviewEndpoints` serves it, `ReviewApiParser`/`ReviewSummaryPresenter` read and
   word it, and the window draws it. "Drill in from there" is task 4's visual diff and is next.
4. **Render changes visually** - the reason this is a desktop app. **[DONE 2026-09-22]**
   - a moved highlight drawn on the schematic, before and after. **[DONE 2026-09-22]**
   - images side by side. **[DONE 2026-09-22]**
   - scope baselines plotted. **[DONE 2026-09-22]** - but see the note above: a baseline is a
     captured IMAGE, not samples, so there is nothing to re-plot. What ships instead is the thing
     the picture cannot show - that the SETTINGS moved underneath it, which makes two traces
     non-comparable while looking like a change in the circuit.
   - text rows as a field-level diff. **[DONE 2026-09-21]**

   **The ASSET ENDPOINT the visual sub-parts are built on** (2026-09-22) -
   `GET /api/review/submissions/{id}/submitted/{hash}` and `.../published/{**path}`, with
   `ReviewAssetLocator` owning the containment and `ReviewImageComparison` deciding which pairs
   are worth showing. `ReviewApiClient.GetSubmittedAssetAsync`/`GetPublishedAssetAsync` fetch the
   bytes, and `schematicImages` tells the app which picture a highlight belongs on.

   **Read `ReviewAssetLocator`'s header before touching either route.** The two sides take
   different input and are guarded differently, and the published side is a genuine
   file-disclosure boundary - the containment tests there are written to fail when the guard is
   removed, because that exact class of test has been caught vacuous twice in this phase.
5. Three outcomes, all first-class:
   - **Approve** - merges, increments revision, regenerates `dataChecksums.json`, notifies.
     **[DONE 2026-09-22]** - `ApprovePublishFlow` owns the chain and
     `POST /api/review/submissions/{id}/approve` is mapped. Administrator only; every refusal
     before the write leaves the tree untouched.
   - **Reject** - with a reason. **[DONE 2026-09-22]**
   - **Request changes** - returns the submission to the contributor's CRT as an editable draft with
     the comment attached. **Do not skip this.** Most imperfect contributions are fixable by their
     author in a minute, and a reject that could have been a conversation costs a contributor.
     **[DONE 2026-09-22]** - server side. The contributor-side half (a returned submission
     arriving back in CRT as an editable draft) is not built yet.
6. Merge writes into the BETA tree using `CRT.Data`, producing the same file layout as today.
   **[DONE 2026-09-22]** - `PublishMerge` assembles the board, `PublishPlan` decides the paths and
   the generation, `PublishExecutor` writes the workbook AND the sidecar.
   **Into the NEWEST GENERATION ONLY** - see open question 3. A new system entering the `v2.0.0`
   generation is `Data {Hardware} {Board} v2.0.0.xlsx`; the unversioned tree is never written.
   `DataGenerationRules` owns that decision. **The board workbook WRITER does not exist and must
   be built** - `BoardDataWriter` was retired in Phase 2 (see open question 8). [The writer landed
   2026-09-21 as `BoardWorkbookWriter`.]

   **THE JSON SIDECAR - found missing 2026-09-22, FIXED the same day.** A board is TWO files: the
   `.xlsx` and a `.json` beside it. `BoardDataReader` loads component highlights from the sidecar
   (`BoardComponentHighlightStorage.LoadComponentHighlights`), and `KiCadCalibrationEntry` lives
   there too under its own `"KiCad calibration points"` root - `BoardData` does not even carry
   calibrations as a section.

   `BoardWorkbookWriter` writes only the workbook, so publishing would have **dropped every
   component highlight on the board** - the rectangles the entire app is built around - plus every
   KiCad calibration. Nothing was broken in practice only because nothing called the publish path
   yet.

   **`BoardSidecarWriter` (CRT.Data, +20 tests) closes it**, and `PublishExecutor` now writes the
   sidecar immediately after the workbook, before anything advertises the revision.

   - **It writes the COMPLETE state**, which is why it is not
     `BoardComponentHighlightStorage.SaveComponentHighlights`. That method replaces ONE schematic
     and preserves every other - right for the app's label editor, wrong here, where the manifest
     is by contract the complete intended state. A schematic whose highlights a submission DELETED
     must actually lose them; a preserving writer would leave them published forever with nothing
     reporting it. Making it preserve turns three tests red.
   - **Both roots are written in ONE pass over one document.** Writing them independently - open,
     set a root, save, reopen, set the other - is the obvious shape and loses one of the two.
   - **Round-tripped through the SHIPPED reader**, never a test-only parser, so a writer that
     agreed with a broken reader could not pass.
   - **Invariant-culture parsing.** Under `da-DK`, `100.5` read culture-sensitively becomes 1005 -
     a highlight ten times across the board, written into a file every user downloads. The
     project's fifth encounter with that bug and the worst place for it, since the damage ships.

   **`DescriptorWithWorkbook` now takes the SIDECAR's hash too, and it is equally required.** A
   submission that merely MOVES A HIGHLIGHT changes no workbook row (highlights are not in the
   workbook) and uploads no file, so neither the uploaded-file list nor the workbook hash moves.
   Without the sidecar folded in, the content hash would be identical to the previous publish's,
   **no client would re-download, and the moved highlight would never reach anybody.** Pinned end
   to end by `A_HIGHLIGHT_ONLY_change_MOVES_the_published_content_hash`.

   Removing the sidecar write turns **ten** executor tests red.
7. ~~**Retain every published revision** so any merge can be rolled back by republishing the
   previous one.~~ **STRUCK 2026-09-21 by the project owner - see open question 5.** No publish
   history is kept; a published file is overwritten in place. A bad merge is undone by publishing
   a correction. `systems.current_revision` still increments, because a submission diffs against
   it - that is not rollback.
8. Ship it through the existing Velopack pipeline, as a **separate** workflow with its **own**
   version. Do not entangle it with CRT's release. **[DONE 2026-09-22]** -
   [build-and-release-maintainer.yml](../.github/workflows/build-and-release-maintainer.yml). Prefixed tag
   and packId, released into its OWN repository (`HovKlan-DH/Classic-Repair-Toolbox-Maintainer`) on its
   own Velopack CHANNELS (`maintainer-win`/`maintainer-linux`) - the packId alone does not keep it out of
   CRT's update feed; see task 8's section; its own generated release body, never `CHANGELOG.md`.

### Definition of done

- A real submission can be reviewed and merged end to end into BETA.
- The merged board opens correctly in CRT after a sync from BETA.
- The merge wrote ONLY the newest generation, and the older generation's files are byte-identical
  afterwards. **This replaces the struck rollback line**, and it is the one that now protects the
  compatibility trees.
- "Request changes" round-trips back into the contributor's drafts.

### Traps

- **Never merge into Production in this phase.**
- **Regenerating the manifest is not optional.** The app syncs only files whose checksum changed, so
  a merge without a manifest update is invisible to every user. The current PHP does this
  automatically after each merge; preserve that behaviour exactly.
- **Orphan cleanup must fail closed.** The existing PHP refuses to delete when the referenced set
  cannot be determined. Keep that. Read the "Orphan cleanup" section of
  [the webserver README](Webserver/app-contribution/README.md) before touching deletion.

---

## Phase 6a - Drafts as real board folders [DONE 2026-09-23]

**Goal.** A draft on the contributor's machine is indistinguishable from a published board
folder, so they can edit it in CRT or in Excel and it makes no difference which.

Numbered 6a because it is not the project owner work Phase 6 describes - it landed between 5 and 6
at the owner's request and has nothing to do with per-system maintainers.

### Why this reverses a Phase 2 decision

Phase 2 stored a draft as row deltas (`draft.json`), and explicitly rejected writing an `.xlsx`
under `Drafts/` as "the same second-mechanism mistake the label editor's redirect already
rejected". That reasoning was sound for what it addressed: an EMPTY workbook sitting ALONGSIDE
`draft.json` would have been two mechanisms for one job.

This is the opposite. The workbook REPLACES `draft.json` as the single source, so there is exactly
one mechanism - and the project owner's actual requirement could not be met any other way:

> "I want to have exact same format as any normal/released board [...] Then the user can decide
> himself, if they want to do things inside of app or edit directly the Excel file."

A delta list cannot satisfy that. There is nothing to open, and an edit-time ledger cannot survive
the contributor editing the workbook with the application closed.

### What a draft folder holds

```
<DraftsRoot>/Commodore/C64/250407/
    Data C64 250407.xlsx      the board workbook, published schema, published file NAME
    Data C64 250407.json      the sidecar: component highlights + KiCad calibrations
    Sheet1of5.png             images at their real relative paths, no Files/ subfolder
    KiCad data/
    .crt-draft.json           LOCAL ONLY - never submitted, never published
```

### The four ideas the whole change rests on

1. **Seeding is a COMPLETE copy.** A draft workbook starts as the published one, so a row that is
   absent has been DELETED. Seed partially and every untouched row reads as a deletion - the change
   list becomes noise and a submission asks the server to delete most of the board.
2. **What changed is DERIVED, never recorded** (`BoardDataDiffer`). A ledger cannot survive an edit
   made in Excel; a comparison is correct whichever route made the change. This is what makes the
   two editing routes genuinely interchangeable rather than merely both possible.
3. **Every save is read-modify-write against the FILE** (`DraftWorkbookStore.Edit`), never against
   a board the application read earlier. Writing back a stale copy would silently discard an Excel
   edit, which is data loss the contributor could not anticipate.
4. **The MARKER is what makes a folder a draft**, not the presence of a workbook. A stray `.xlsx`
   copied into `Drafts/` by hand is not a draft, and reading it as one would substitute unknown
   data for the published board.

### What was lost, stated plainly

The drift report could say two things it no longer can: *"the row you edited has since disappeared
officially"* and *"a row you added now exists officially too, and theirs is the one shown"*. Both
were consequences of the MERGE - `BoardDraftApplier` silently dropped the first and failed closed on
the second - and the report existed to surface outcomes that were otherwise invisible. With no
merge, neither outcome exists.

The revision comparison still says the published board has moved; nothing now says which of your
edits that collides with. Recovering it needs the board as it was when the draft was SEEDED, which
the draft folder keeps no copy of. That is a feature in its own right, not an oversight - the
"needs attention" panel is kept in the markup, hidden, so it has somewhere to go.

### Three real bugs this surfaced

- **The board cache was keyed by the SYSTEM, and a drafted system has two workbooks.** Safe while a
   draft was an overlay applied after the cache lookup; not safe once the toggle switches between
   two files. Turning "view as officially published" off again kept showing the published board.
   Now keyed by the resolved PATH, and three call sites that cleared the old key were clearing
   nothing.
- **Component highlights were being silently dropped on save.** `BoardWorkbookSchema` has no sheet
   for them - they live in the JSON sidecar - so writing only the workbook lost every rectangle the
   label editor had just saved, with the save reporting success. Found by asking why nothing in
   production called `SaveComponentHighlights` any more.
- **`UuidV4` was write-dead across the whole codebase.** Nothing generated one; the only remaining
   consumer was a validator warning that a dead field needed fixing. Removed entirely at the
   owner's instruction (2026-09-23), including from the workbook schema.

### What was retired

`BoardDraft`, `DraftDataStore`, `BoardDraftApplier`, `DraftDriftReport`, `DraftBaseRevision`, and
the four `*DraftWriter` classes, with their tests. `BoardDataReader.LoadAsync` no longer takes a
draft. `BoardDraft.KiCadCalibrations` - "the eleventh section" - retired with them: calibrations
always had a published home in the sidecar, which a draft folder now carries unchanged.

### Definition of done

- A draft folder opens in Excel and round-trips through CRT. **Not yet verified against a real
  board by the project owner** - the suite covers it, running the app does not.
- No migration exists and none is needed: nothing had been released.
- The server contract is untouched. `SubmissionManifestBuilder` always wanted the complete intended
  state, so no wire format changed and no server work was required.

### Follow-on: editing a draft as a table [CLIENT DONE 2026-09-24; maintainer side DONE 2026-09-25]

A draft being a real workbook is what made this possible: the Drafts tab's **"Edit in table
format"** shows the draft's nine sheets as an editable grid, coloured against the published board
(green added, orange modified cell, red deleted row shown where it was, and violet for a row
flagged "!" - a duplicate key or an incomplete row, which is not a change and not counted as one).
CLAUDE.md's Drafts paragraph has the rules; the short version is that ALL logic is in `CRT.Data`
(`BoardTableDocument`, `BoardTableSheet`, `DraftTableSession`, `BoardTableClipboard`, and
`BoardTableHistory` for Ctrl+Z / Ctrl+Y) and the app's `BoardTableEditor` only paints it.

**The project owner wants the same table in `CRT.Maintainer`, so a maintainer can make the same edits.** Built
2026-09-25 - see Phase 6, "The maintainer's table". What it needed, as planned here:

- **The control moves to a shared Avalonia library** referenced by both apps (it touches nothing of
  `Main` or `DataManager`, by design). Its `BoardTable_*` theme keys then have to exist in BOTH apps'
  resources, and ProDataGrid's theme include with them.
- **The server is the real work, not the grid.** A maintainer's edit is a new version of the
  submission: it needs an upload route, a `ReviewAuthority` rule for who may amend, an audit entry,
  and a way for the contributor to see what the maintainer changed. None of that exists yet, and all
  of it is covered by "One change, every side of it".
- **The table would compare against the submission's BASE revision** on the server, where
  `BoardDataDiffer` already runs - not against a local published copy as the app does.

**A shared identity rule changed with it (2026-09-24): a component row's natural key is now its
label PLUS its region** (`BoardDraftNaturalKeys.ForComponent`). This reaches the server and the
maintainer app through `ReviewSummary`, deliberately: before it, a submission adding U1/NTSC beside an
existing U1/PAL showed the maintainer nothing at all. Rows without a region key exactly as before. A
component GIVEN a region now reads as a removal plus an addition in review, the same as a label
change. Nothing persisted a component key, so no stored data needed migrating.

---

## Phase 6 - Maintainers [DONE 2026-09-25: roles, two-stage publish, two-person approval, orphan removal, the maintainer's table]

**Goal.** Per-system maintainers who approve and publish work on their own systems.

### Roles - TWO, by the owner's decision (2026-09-25)

The plan below this line was written for FOUR roles. The project owner collapsed it in one sentence:
*"Only those two roles. An Administrator will probably be only ONE person, me, having access to
everything and can also do review and whatever. Then a Reviewer [now called Maintainer] is someone I assign specifically to
a system, and then that person can review and publish changes for that specific system. The person
may be able to maintain multiple systems, if I associate him to multiple systems."*

| Role | May | Blast radius if the account is stolen |
| --- | --- | --- |
| **Administrator** | Everything: review and publish every system, assign and remove maintainers, co-approve shared-file changes | Everything |
| **Maintainer** | Review AND publish the systems they are assigned to. Nothing on any other system | That maintainer's systems only |

Mapped onto the old table: the new role is the old **Maintainer** (for a few hours on 2026-09-25
it was called "Reviewer" - see the naming note below), and the old recommend-only **Reviewer**
(blast radius "none") **no longer exists**. That is a change of security
model, made by the project owner explicitly - which is exactly what the old traps said such a request
must be - and it is NOT the dangerous case they warned about (a global publisher): a stolen maintainer
account reaches only that person's systems. Contributor is not a role at all; contributing needs
no account (Phase 4).

Four properties survive from the four-role design and are load-bearing:

- **A system has a POOL of maintainers, not an owner.** Any maintainer of a system may act; there is no
  rank and nothing to transfer. The `maintainers` table (renamed to `reviewers` by migration 0006
  and back by 0010) is a set of (system, account) pairs.
- **The administrator is in every pool by definition**, computed, never by rows and never by an
  override path. `accounts.is_administrator` is still granted by hand only (DEPLOYMENT.md).
- **A shared-file change needs TWO approvals: a maintainer of the board AND the administrator**
  (owner decision, 2026-09-25: *"in case of changes to any shared file, then both the maintainer
  and the admin should approve before publishing to BETA or production. If there is no shared files
  changed, then normal maintainer is sufficient."*). A submission that adds or changes a file under
  `Shared files` or `Generic shared files` reaches every board citing it. Decided at create
  (`SubmissionSharedFiles`, CRT.Data), stored as `submissions.touches_shared_files`. The rule is
  `ApprovalRules` (CRT.Data, pure) - see "Two approvals for a shared-file change" below. This
  REPLACES the first version of the day, where such a submission was the administrator's ALONE and
  hidden from maintainers.
- **A system with an empty pool is normal** - it means the administrator handles it, which is also
  how "a NEW system routes to the administrator, always" falls out with no special case.

### What was built (2026-09-25)

- **`ReviewAccess` + `ReviewAuthority` (CRT.Server).** One object per request - the account and
  the set of systems it reviews, read from the pool table in `ReviewEndpoints.AuthenticateAsync` on
  EVERY call - and one rule: "administrator, or in THIS system's pool". `CanReview` and
  `CanPublish` are two questions with one answer today, pinned by a
  property test so a later split is deliberate. Every route naming a submission checks it against
  that submission (threat 3, "check the object"); the queue is filtered by it; the two asset routes
  refuse another system's bytes. `ReviewDecisionRules` and `ApprovePublishFlow` take the access
  object and refuse per system with a sentence naming the system.
- **Removal bites on the next request**, proven by
  `MaintainerAssignmentFlowsTests.Removing_a_maintainer_takes_effect_on_the_very_next_request` -
  Phase 6's definition of done, verbatim.
- **Migration 0006**: `maintainers` -> `reviewers` (named back by 0010); `accounts.is_reviewer` dropped (it meant the
  role that no longer exists); `submissions.touches_shared_files` added.
- **The administrator's API and screen.** `/api/admin/systems`, `/api/admin/accounts`,
  `POST /api/admin/maintainers` and `/maintainers/remove` (`AdminEndpoints`, a rim over
  `MaintainerAssignmentFlows`), administrator-only with a negative test. The systems list is the
  `systems` rows UNIONED with the boards in the data tree (`PublishedSystemLister`), so a shipped
  board can get a maintainer before its first submission; the first assignment creates its row as
  'shipped'. An unverified, locked or administrator account is refused with a sentence
  (`MaintainerAssignmentRules`), and the maintainer app's list says the same sentence before the button
  is pressed (`MaintainerAssignmentDisplay`). Every grant and revocation is an audit row. In
  CRT.Maintainer: a **Maintainers** button above the queue, shown only when the queue answer says
  `isAdministrator`, opening `MaintainersWindow`.
- **Maintainers are told.** On finalise, whoever must approve is e-mailed (`SubmissionRouting`,
  `EmailTemplates.SubmissionWaiting`): the system's maintainers; the maintainers AND the administrators
  on a shared-files change; the administrators alone when nobody is assigned. Never fails the contributor's request. The "somebody else handled
  it" mail of task 11 is not built - a decided submission simply leaves the queue.
- ~~**`system.json` mirrors the pool** (task 1)~~ - withdrawn the same day with `system.json`
  itself (see its section above). The pool lives in the database only.
- **Vocabulary.** "Maintainer" everywhere a user reads it: CRT Maintainer's sign-in text, CRT's
  submit dialog and the new-system agreement (`NewSystemMaintainerWindow`), the Wiki pages.
- **Naming note (owner decision, later on 2026-09-25).** The role was first renamed from Maintainer
  to "Reviewer", then back: "Reviewer" became "maintainer" everywhere - code, database (migration
  0010 renames the pool table back to `maintainers` and the stored approval roles), API and text -
  and the review application became **CRT Maintainer** (`src/CRT.Maintainer/`). What this document
  used to call "the maintainer" - the person who owns the project - is now "the project owner", so
  "maintainer" only ever means the role. Older passages that quote the intermediate "Reviewer" name
  were rewritten with the rest.

### Deliberately NOT built, and the accepted risk

- **TOTP two-factor (task 7, threat 2).** The project owner chose to defer it (2026-09-25) and to open
  publishing to maintainers without it. **Recorded here as an accepted risk:** a stolen maintainer
  account can publish to that maintainer's systems - to BETA today, and to Production once the
  two-stage publish exists - with a password as the only factor. The remaining safeguards are the
  server-side validation, the review itself, the audit rows, and the project owner's own backups. It
  stays on the security review's open list until it is built.
- **The administrator feed (task 9)** and the anomaly alerts of threat 2. The audit rows exist;
  nothing renders them yet.
- **Credits from account identity (task 10).**

### Two-stage publish: BETA, then Production [DONE 2026-09-25]

The project owner: *"it should be a two-fold process, where it is first published to BETA and then it
is published to the real production. The maintainer is still allowed to do this, but only after he
has checked that the data looks correct in BETA."* **This REVERSES open question 6's answer** (BETA
to Production was a manual copy, and the service could never write Production) - by the
owner's explicit decision, and only when switched on. **The environment stays named BETA**: it
is what users see in CRT's Configuration tab and what every installed build's sync URL carries.

- **Approve writes BETA, exactly as before.** "Publish to production" is a second, per-SYSTEM act
  in the maintainer application (a **Production** window), allowed to whoever may publish that system.
  Per system rather than per submission because BETA is one tree - two merged submissions to a
  board are in one workbook.
- **`ProductionPromotionPlan` (CRT.Data, pure)** decides what is copied: every file under the
  system's BETA folder that Production lacks or holds differently, plus every SHARED file the BETA
  board cites that differs. It refuses a board citing ANOTHER board's file that is not already
  identical in Production (that board goes first), a cited file BETA itself lacks, and a case-only
  collision with Production. What the board no longer uses is removed from production, from a list
  shown before approving - see "Orphan files" below. Order: content, then workbooks and sidecars; a `system.json` is never
  carried and one already in production is removed. The same records go to the maintainer app as `PromotionFile`, so the maintainer sees the
  exact list the server then performs.
- **`ProductionPromoter`** resolves and link-checks every source and destination before the first
  write, then copies each file through **`VerifiedFileCopy`** (hashed as written, renamed in only
  on a match - the same helper `BlobStore.TryCopyToAsync` now uses) against the hash the plan saw.
- **"Only after he has checked it in BETA", in code:** the request carries back the BETA content
  hash the maintainer was shown, and `ProductionPromotionFlow` refuses (409) if a publish has landed
  in BETA since. The maintainer app adds the human half - a box the maintainer ticks - and the button
  follows the server's `canPublish` AND the tick (`ProductionDisplay.CanPress`).
- **One `PublishLock`** serialises the BETA publish and the promotion, so a promotion can never
  copy half a publish and the hash check means something.
- **A copy list holding a shared file needs the maintainer AND the administrator** here too
  (`TouchesSharedFiles` on the plan; see the next section).
- **Off until configured.** `ProductionDataTreeRoot`, `ProductionManifestPath`,
  `ProductionPublicDataBaseUrl`: all three or none; none may carry the `-BETA` marker or equal its
  BETA twin; the data root must sit inside `ProductionTreeRoot` and be writable at startup.
  DEPLOYMENT.md step 13 is the procedure, including the permissions and `ReadWritePaths` change
  that undoes step 3 for the production data folder only.
- **Migration 0007**: `systems.production_revision`, `production_content_hash`,
  `production_published_utc`. "Waiting for production" is BETA's `content_hash` differing from
  `production_content_hash` (`ProductionPromotionRules.IsAwaitingProduction`) - the hash, not the
  revision date, because two publishes in a day share a date.
- **After a promotion** (none of it may fail the request): Production's `dataChecksums.json` is
  regenerated; the contributors of every submission merged since the previous promotion are mailed
  "published"; the administrators are mailed when a maintainer did it (the feed's stand-in); an audit
  row names who.
- **The contributor's view changed with it (the "every side" rule).** A maintainer's approval mail
  now says "published to the BETA source", with one more mail to come
  (`SubmissionPublishedToBeta`); the promotion sends "published to the source"
  (`SubmissionPublishedToSource`) - the project owner's own words for the two stages, matching the
  "source" / "BETA source" names in CRT's Configuration tab. The server reports a merged submission as `published` once its
  system has been promoted since (`ContributorFacingState`), without touching the stored state.
  CRT shows "merged" as "Published to BETA source" and "published" as "Published to source", and
  keeps asking about a merged one at launch until it reads "published" - at every launch for 30 days
  after the decision (`SubmissionReceiptPresenter.MergedRecheckWindow`), since the second publish
  may never come, and weekly after that (`MergedLateRecheckInterval`, 2026-09-27) because a BETA
  rollback can turn a merged submission "returned" at any time; "Check for updates" in My
  submissions still asks whenever pressed. `DraftRetirement` needed
  no change: it already waits for the contributor's own synced data to carry the change.
- **~~Known gap: a NEW system is never added to the master workbook~~ - CLOSED 2026-09-27**, see
  "A new system's place in the drop-down lists" below. A maintainer places the system on the
  Systems screen before it can be approved, and the publish and the promotion each add its row.

### Orphan files [DONE 2026-09-25]

The project owner: "My goal at least is that there must be no orphan files." Until now a publish never
deleted anything, so a file a board stopped using (replaced under a DIFFERENT name, or no longer
cited) stayed in BETA and production and was synced to every user for ever. A file replaced under
the SAME name is simply overwritten and leaves nothing behind.

**The rule is `DataTreeUsage` (CRT.Data), one definition for every caller.** A file is USED when it
is (1) a master workbook, (2) a board workbook - one a master lists, OR one found at the top of a
board folder - of every generation, and its `.json` sidecar, (3) cited by any of those workbooks,
(4) inside a board's `KiCad data` folder, (5) inside a folder CRT reads by name
(`FoldersReadByName`, today only `Generic shared files/MiniPro/IC tests`; `IcTestCatalogue` now
builds its paths from the same constant, and `DataFoldersReadByNameTests` in CRT.App fails if the
app and the rule disagree on it or on the KiCad folder name), or (6) a file whose name starts with
`!`. Matching ignores case, the safe direction for a rule that deletes. **Fail closed:** an
unreadable master or board workbook, a master listing a workbook the tree lacks, or a folder that
cannot be walked makes the result incomplete, and an incomplete result removes nothing. A board
found in the tree counts even when no master lists it, because a new system the server publishes is
not added to any master (done by hand; since 2026-09-27 the publish adds it, but a board copied
into a tree by hand still has none) - trusting the masters alone would delete it.

**The shipped data was cleaned first.** Run over `Assets/Data` the rule found 2 masters, 22 board
workbooks, 10,971 files and **50 orphans (8.7 MB)**: 11 `.fsc` image-editor files, 3 VGG Image
Annotator project files (`CPC664_*.json`), 29 component images/PDFs no board uses, 5 scope captures
no row cites and 4 readme/introduction texts. The project owner reviewed the full list and approved
removing all of them (2026-09-25); they are deleted from `Assets/Data`. A further 8 files are cited
only by an older generation and stay. **`DataTreeUsageShippedDataTests` now fails, naming the file,
whenever a file enters the shipped tree that nothing uses.** The live BETA and production trees are
cleaned through the administrator's list below, not by hand.

**When files are removed - always from a list somebody has seen:**

- **A BETA publish** removes what the board stops citing that nothing else in BETA uses
  (`ApprovePublishFlow.PreviewRemovals`: the candidates are `DataTreeUsage.NoLongerCited`, and the
  tree is read with the workbook the plan WRITES standing in for its new citations). The project owner
  required the list to be "visible BEFORE the maintainer/admin approves it ... so it is clear what will
  happen": the submission detail carries it as `removals` (CRT.Data's `FileRemovalPreview`), the
  maintainer app lists each file as REMOVED and says so beside the Approve button, and the approval
  sends the list back. **The server refuses (409) when the list it would now remove differs** -
  another publish may have started or stopped citing a shared file - so what goes is exactly what
  was on screen. After the write, `UnusedFileRemover` removes only those files, and only if the
  tree, read again, still does not use them.
- **A production promotion** does the same for production (`ProductionPromotionFlow.PreviewRemovals`:
  production's workbooks for the system are compared with BETA's, which replace them). The plan
  carries `removals`, the Production window lists them first in red, and the publish request sends
  them back with the same refusal rule.
- **Files that were already orphans** are the administrator's: the **Unused files** window in the
  maintainer app (`/api/admin/unused-files`, `UnusedFileFlows`), per tree, administrator-only. It lists
  every unused file with its size; Remove needs the administrator's tick in "I have looked through
  this list"; the server removes only the files sent that it still finds unused, regenerates that
  tree's `dataChecksums.json`, and writes an audit row (`data.unused_removed`) naming every file.

**Every removal** goes through `SubmissionPathRules` and `PublishPathSafety` (no link is followed),
considers only files the sync manifest would list (never a dot-file or a half-written `.tmp_`),
removes a folder it leaves empty, runs under the `PublishLock`, and never fails the publish it
follows. **A removed shared file needs no second approval**: it is removed only when no board in
that tree uses it, so it changes nothing any board shows. (Adding or changing one still needs both
approvals - see above.)

**Known limits, deliberate for now:** a board workbook with a MISSING sheet reads as citing nothing
from that sheet (`BoardDataReader`'s existing behaviour, which CRT's own cleanup shares), not as
unreadable; and a file removed from a board's `KiCad data` folder in BETA stays in production,
because that folder is kept whole.

**On users' machines:** a file removed on the server drops out of that tree's dataChecksums.json,
but CRT deletes its local copy only when "Delete orphan and non-used files" is on - and that setting
is OFF by default with its checkbox disabled in the Configuration tab. Whether to switch it on is
the owner's decision, not yet taken.

### The maintainer's table [DONE 2026-09-25]

The project owner: *"make the same 'Edit in table format' (maybe call it 'View in table format')
available in the maintainer app ... The maintainer should be able to also edit whatever, if he chooses to
publish it afterwards."* The button is **"View in table format"**, beside the decision buttons.
(Since 2026-09-26 there is no button: the table IS the submission view - see "The table is the
submission view" below.)

- **The editor is shared, not copied.** `BoardTableEditor` and `UnsavedTableEditsWindow` moved from
  CRT.App into a new Avalonia library, **`src/CRT.UI/`**, referenced by both applications, with
  `ThemeResources` (the one two-step theme lookup). The `BoardTable_*` colours moved into
  `CRT.UI/BoardTable/BoardTableColors.axaml`, merged by both apps' `App.axaml`, so the two tables
  cannot drift. The maintainer app defines the six general keys the editor borrows (`Bg`, `Fg`,
  `Table_Bg`, `Table_BorderRowLine`, `Text_Fail_Fg`, `Button_Cancel_*`) and includes ProDataGrid's
  theme. The Drafts tab is unchanged - its 138 tests passed across the move.
- **Document mode.** `BoardTableEditor.Open(BoardTableDocument)` shows a table with no draft file
  behind it: "Save changes" raises `SaveRequested` and the host saves; Reload and the draft-file
  notices are hidden; it opens on the first sheet with a change. The review window
  (`ReviewTableWindow`) builds the document from the published board and the submission
  (`BoardTableDocument.Create`, the Drafts tab's own rule), and asks before unsaved changes are lost
  (the prompt's `LeavingSubmission` wording). **Since 2026-09-26 there is no separate window**: the
  table opens in the submission panel in the summary's place (`MaintainerMain.Table.cs`), and every
  way of leaving it asks the same question - see the follow-up below.
- **An amendment is the server's decision, not the client's** (`AmendSubmissionFlow`,
  `POST /api/review/submissions/{id}/amend`; the table is `GET .../table`, CRT.Data's
  `ReviewTableData`). In order: authority over THIS board; still undecided (pending or waiting for
  its second approval); still at the amendment VERSION the maintainer opened (a second maintainer's save
  is refused naming who changed it); only the table's nine sheets are taken
  (`SubmissionRowsBoard.WithTableSections` - highlights, calibrations and the revision date stay as
  submitted); FILES are rebuilt from what the rows cite - a file the submission carries is kept, a
  file already PUBLISHED is taken from the tree (imported into the blob store, as create does), and
  anything else is refused (`amend.file_unknown`: a maintainer edits rows, and cannot bring in a file
  nobody sent); then the same path, file and row rules a new submission passes.
- **Stored beside the original.** `submission_payloads`/`submission_files` hold the current content,
  so nothing that reads a submission changed; migration 0009's `submission_amendments` keeps what
  each amendment replaced - its first row is the contributor's original. The store's `AmendAsync`
  is one transaction, and it **clears the approvals already given** (they were given to other
  content), moves 'approved' back to 'pending', and re-decides `touches_shared_files`, so the
  two-person rule follows the content actually published. An audit row (`submission.amended`)
  records who.
- **The contributor is told.** The BETA-publish mail says a maintainer changed some details; the status
  answer carries `amendedByMaintainer` (CRT.Data's `SubmissionStatus`), stored on the receipt and
  shown in "My submissions" (`SubmissionReceiptPresenter.DescribeAmended`). The maintainer app's
  submission view names who last changed it.

**Known limits, deliberate for now:** a maintainer cannot add a NEW file through the table (rows
only). Deleting or renaming a SCHEMATIC takes its highlights and calibrations with it
(`SubmissionRowsBoard.WithTableSections`): a rename is recognised by the same image file, and an
ambiguous one drops them rather than guess. Until the code review below, deleting a schematic with
highlights was refused and nothing in the table could fix it.

### A deleted component takes everything with it [DONE 2026-09-25]

The project owner: *"if a component really is deleted, then it should remove EVERYTHING related to
this component."* Deleting one used to leave its highlights behind everywhere, and in the table its
image, file and link rows too.

- **The table (both apps):** deleting a Components row also deletes that component's rows on the
  image, local file and link sheets - red on their own sheets, ONE undo step
  (`BoardTableHistory` steps can now span sheets) - and the save drops its highlights
  (`BoardTableDocument.ApplyTo`). A rename keeps them: the document remembers the component rows
  it opened with, so a row object still live under a new label is a rename, not a delete. A
  regional variant deleted beside its twin takes only its own region's images. The status line
  says what else went (`BoardTableDeletedWith`).
- **The Contribute window's "Delete this component"** now drops the highlights too
  (`ComponentBoardWriter.ApplyComponentDelete`). Its notice already said they would go.
- **The maintainer's amendment:** `WithTableSections` drops a highlight only when the edit left it
  out AND no component in the edit has its label - so the route still cannot remove a highlight of a
  component the board keeps.

### A new system's draft is retired like any other [DONE 2026-09-25]

It used to be refused outright. Now it is retired once its published board is on the machine and
matches (`DraftRetirement`), which can only happen after the master workbook lists the system -
CRT downloads a board workbook only then - so retiring it never hides the contributor's board. The
draft and the published workbook have different names (the draft's has no generation suffix), so
`DraftStatusReader.ResolveForSystem` finds the draft by its folder and keeps the draft's own key.
It depends on the master row, which the publish has added since 2026-09-27 (see "A new system's
place in the drop-down lists"). **And since 2026-09-27 it is retired only once the submission is
PUBLISHED TO PRODUCTION** - see "Drafts stay until production" below.

### Code review follow-up [DONE 2026-09-25]

A review of this phase's work found fifteen problems; all were fixed the same day. The ones that
change a design, for a later session:

- **The review API's bodies are CRT.Data's `ReviewApiContract` records.** The requests were
  records inside the server's endpoint classes while the maintainer app sent anonymous objects, and the
  answers were anonymous objects on the server. Now both ends build the same records, the JSON
  settings are one method both apply (`ApplyWireSettings`), and CRT.Maintainer.Tests'
  `ReviewWireContractTests` serialises each answer with the server's settings and parses it with
  the real parser. It caught a startup crash in the shared settings on its first run.
- **Saving the maintainer's table was refused with 413 for every real board**: the amend route fell
  to the 64 KB default body limit. It now gets the manifest's 8 MB; the approve, production-publish
  and unused-file routes, which send a file list, get 2 MB (`RequestBodyLimits`).
- **A pool row counts as "the board has a maintainer" only when that account can approve as one**
  (`ReviewAuthority.CanGiveMaintainerApproval`: verified, not locked, not an administrator). An
  account granted a pool and later made administrator by hand made a shared-file change wait for
  ever. Routing uses the same rule.
- **A board's second maintainer is no longer told "you have already approved"** for a colleague's
  approval. `GivenApproval.AccountId` (kept off the wire) and `ApprovalStatus.YouApproved` tell the
  two apart; the server's refusal names who approved, and the button reads "Another maintainer
  approved - waiting for ...".
- **Creating a submission holds the blob reference gate only for its last step.** Imports from the
  published tree used to run inside the service-wide gate. Now they run outside it; inside, each
  "held" blob is confirmed still there (else asked for) and the row is created. A budget refusal
  after the imports takes back what this create imported.
- **The removal preview no longer re-reads the whole tree on every click.** It builds no plan when
  the board stops citing nothing, and what it reads goes through `WorkbookReadCache` (once per
  file version). The removal itself never uses the cache.
- **The published-tree probe hashes synchronously** instead of blocking a request thread on an
  async read.
- **A "merged" receipt is checked at launch for 30 days at most** (see above).
- **Every shipped board's folder names are run through the identity rules**
  (`SubmissionRulesShippedDataTests`), and a path the path rules refuse is no longer reported a
  second time by the file rules.

### Second code review follow-up [DONE 2026-09-25]

A second review found twelve more; all were fixed the same day. The ones that change a design:

- **"Changes a shared file" is judged against the tree as it is NOW** (`ApprovePublishFlow.
  TouchesSharedFilesNow`). The flag stored at create compared the submission with the tree of that
  moment, so a submission citing a shared file unchanged stayed a one-approval item even after
  another publish changed that file - and would then have reverted it for every board on one
  maintainer's say. The stored flag is now a floor; the approval and the detail screen re-check,
  and the approval stores a raised flag (`MarkTouchesSharedFilesAsync`) so the queue agrees.
- **No approval is recorded for what cannot be published.** The payload and the plan (both
  read-only) now run BEFORE the first of two approvals is recorded; they used to run after, so
  the second approver was mailed for a submission the plan then refused.
- **An item whose required approvals shrank is not stranded** (`ApprovalRules.Status`): if every
  required role has already approved - the board's last maintainer left its pool after the
  administrator approved - the next approval by a role it needs publishes. BETA and production.
- **An amendment takes the `PublishLock`, and the store re-checks the version and the state
  inside its transaction** (`AmendAsync(expectedVersion, ...)`, `AmendStoreResult`). Two
  amendments at once both passed the flow's early checks and the second overwrote the first; one
  could also land mid-publish, leaving the tree with the old rows.
- **An older maintainer application that sends no removal list is told to update**
  (`RemovalsNotSentMessage`, 400) instead of "the list changed since you opened it" (409), which
  was false and could never be fixed by reopening. Approve and production publish alike.
- **Body limits live on the routes** (`.WithBodyLimit(...)` where each is mapped;
  `RequestBodyLimits.For(endpoint)` after `UseRouting`). `RequestBodyLimitsTests` builds the
  server's real route table through `Program.AddServerServices`/`MapServerEndpoints` and fails on
  any route that reads a body without a decided limit.
- **The three lists are contract records too**: `ProductionListAnswer`, `MaintainerSystemsAnswer`,
  `MaintainerAccountsAnswer`, each with a `ReviewWireContractTests` case.
- **A publish no longer rewrites files that are already there byte for byte** - the board's own
  and shared files now follow the rule another board's files did (`PublishPlanDetail.
  UnchangedFiles`, formerly `UnchangedForeignFiles`). A typo fix writes the workbook and sidecar
  and nothing else.
- Smaller: one shared-folder-name rule (`SubmissionFileScopes.IsSharedFolderName`) for the
  validator, `DataTreeUsage` and the maintainer list; the blob store's import uses
  `VerifiedFileCopy`; the Maintainer app's windows share `WindowMessage`.

### The table in the submission panel, and files in the table [DONE 2026-09-26]

The project owner: *"could it instead open in the existing right-side panel, just alike it does in
the CRT app?"* and *"whenever it is a file column, and it is an image, then it should show a
hover-helper with the image ... both the removed and the new image side-by-side ... If the file is
something else, e.g. PDF, then there should be a link that will open the PDF"* - the second for the
CRT app's Drafts table too.

- **No `ReviewTableWindow` any more.** "View in table format" swaps the summary for the table in
  the submission panel (`MaintainerMain.Table.cs`) and reads "Close table" while it is open. The
  modal window guaranteed nothing unsaved was left behind by being modal; the panel asks instead,
  on every way out - closing it, choosing another submission (Cancel puts the selection back),
  signing out, closing the window. A decision is refused while the table has unsaved changes
  (`ReviewTableWording.SaveTableBeforeDeciding`): it would decide the SAVED version. A save reloads
  the table and the detail, so the decision bar beside it acts on the amended content. The queue
  now KEEPS its selection across a refresh (`ApplyQueue`), where it used to blank the panel - which
  would have thrown the table away.
- **The hover card is CRT.UI's** (`BoardTableEditor.FilePreview.cs`, content `BoardTableFilePreview`),
  so both tables have it. Which cells are files and which file each side is, is CRT.Data's
  `BoardTableFileCells`; the bytes come from each host's `IBoardTableFileSource` -
  `DraftTableFileSource` (the local data folder, and the draft's own copy first) and
  `ReviewTableFileSource` (the server's published and submitted asset routes). Not a tooltip,
  because a tooltip closes before its link can be clicked (a flyout at first; an overlay popup since
  - see the next section). The same path on both sides is still compared, by bytes - a picture
  replaced under its own name is not coloured in the table and is the change most worth seeing.
- **The server now serves a published file the PUBLISHED board cites** as well as one the
  submission cites (`ReviewAssetLocator.TryLocatePublishedFile`'s `publishedBoardFiles`, read
  through `ApprovePublishFlow.PreviewReads` only when the submission does not cite the path). The
  old picture of a changed or deleted row is cited by the published board alone, so it answered
  404 - which had also been making every REMOVED image in the change summary read "No published
  file at this path".
- **The maintainer app opens a contributor's file by saving it to the temp folder**, and only a
  type a submission may carry (`ReviewTableFiles.TryGetOpenName`); a web page is saved as `.txt`,
  so its script never runs in the maintainer's browser.
- **One list of drawable image types** (`ImageFileTypes`, CRT.Data), where the component editor and
  the maintainer's image comparison each kept a copy.

### The table is the submission view, and the hover card is instant [DONE 2026-09-26]

The project owner: *"make the table the default first view, as this is the most helpful one. I am
not sure if I can use the other view anymore, as it is confusing to look at, so just scrap that
information"*; *"can it then show image/tooltip instantly, instead of the small delay? And likewise
instantly NOT show the helper when moving mouse outside the helper area"*; the top-left buttons
overlapped the headline; and a smaller header - *"[New system] [Awaiting your review] {ID}"*, then
Manufacturer, Hardware and Board, then the contributor's comment.

- **No change summary in the maintainer app.** Choosing a submission opens its table straight
  away; "View in table format" / "Close table" are gone. Leaving a table with unsaved changes still
  asks - on choosing another submission, signing out and closing the window. A refresh reloads the
  table too unless it has unsaved changes.
- **What the table cannot show stays, as short lines above it** (`ReviewNotInTable`): component
  highlight and KiCad calibration-point changes (both in the `.json` beside the workbook, both
  published by an approval - the summary's calibration section was added precisely so they are not
  approved unseen), and the automatic checks' warnings and errors. The picture of a moved highlight
  on its schematic went with the summary.
- **Now unused by the app, left in place for an owner decision**: `ReviewImageComparison` (its
  file also holds the `ReviewSubmissionAssets` records the parser uses), `ReviewFileComparison`,
  `ReviewHighlightGeometry`, `ReviewHighlightCanvas`, and most of `ReviewSummaryPresenter` and
  `ReviewScopeBaseline` - each still tested. The server still sends what they read.
- **The details are in the QUEUE ROWS, not over the table** (owner, same day: *"I actually meant
  for the [Awaiting...] info and so on to be shown in the left-side menu/list"*). Each row
  (`MaintainerMain.QueueItems.cs`, words from `ReviewQueueDisplay`) shows "New system" /
  "Published system" and "Awaiting your review" badges with the id, Manufacturer / Hardware /
  Board, the contributor's comment and the wait. **The queue answer is now a contract record**
  (`ReviewApiContract.ReviewQueueAnswer` / `ReviewQueueEntry`, built by the server's
  `ReviewQueueFlow`, pinned both ends by `ReviewWireContractTests`), carrying `isNewSystem`
  (`PublishedBoardLocator.LocateSystem` - no manifest loaded) and `awaitsYou`
  (`ApprovalStatus.CanApprove` with the stored shared-files flag; one approvals read per row). Both
  are null from an older server, which shows no badge. The opened row takes the detail's own
  answers, judged against the tree as it is now, so it never says "awaiting" beside an Approve that
  is off. The queue's buttons wrap under the title, and Maintainers / Unused files stay
  administrator-only (`ApplyQueueResponse`, pinned).
- **Every table cell's text tooltip is instant as well** (*"The instant-show should also work for
  texts"*): `ToolTip.ShowDelay` 0 in both cell themes (`BoardTableEditor.CellToolTipDelay`).
- **Then, the same day:** "Show changes only" also hides the tabs of sheets with nothing to show
  (`BoardTableDocument.SheetsShown` / `SheetToShow`, both tables), and the user's choice of it
  survives a table with nothing published; the maintainer's next submission opens on the sheet
  last looked at; a file card's side is labelled only beside another one; and the queue rows lost
  the "#4" (*"I do not see any value in showing the #3 and #4 data"*).
- **And then:** the **queue is grouped by board** (a heading per board, "New system" on it when
  nothing is published; each submission two lines; one not waiting for this account dimmed, "with
  the other approver" - `ReviewQueueDisplay.Group`); the **sheet is remembered per submission**;
  **flagged rows show on their sheet's tab** in a violet pill beside the change count
  (`BoardTableSheetTabHeader`); and an **important signal is keyed on display name AND net**
  (`BoardDraftNaturalKeys.ForKiCadImportantSignal`) - one display name covers several nets, so
  every second one was being flagged a duplicate. That key is shared by the table, the differ and
  the server's review summary, so all three changed together; a changed net now reads as a removal
  plus an addition.
- **Last round of the day:** no Refresh button - the queue checks itself every minute while the
  window is in front and on returning to it, updating only the list (`MaintainerMain.QueueRefresh.cs`;
  an open submission decided elsewhere stays, undecidable); a new system's table is compared with
  the submission itself as opened, so only the maintainer's own edits are coloured (compared with
  nothing it showed a lone "0 Flagged"; an all-green version was turned down); Approve sits at the
  far right; the window's place
  and "Show changes only" are remembered (`MaintainerSettings`); the footer names the account
  without "Signed in as"; and the three decision buttons wear CRT's red (`Button_Cancel_*`).
- **Then, for a new system:** its file card reads both sides from the SUBMISSION
  (`ReviewTableFiles.HashToRead`) - read from the published tree, every picture sat beside "There
  is no file at this path" under "Before (published)" - so it shows one picture, and two only when
  the maintainer names another file ("As submitted" / "Your change"). Its lines above the table
  count instead of naming ("Highlights included for [2] components", "KiCad calibration points included for [1]
  schematic"). The maintainer application's program file now carries CRT's icon (`ApplicationIcon`),
  which is what Windows shows in its title bar and taskbar.
- **And then:** no line says "(not in the table)" any more (it read as if the components were
  missing from the table), each count reads "[2]" with the number bold (`ReviewNoteRun`), a new
  system's file card drops "Unchanged" (`IBoardTableFileSource.SaysUnchanged`), and the table's
  scroll bars stay full size in both applications (`ScrollViewer.AllowAutoHide` off), so a sheet
  wider than the window visibly has more columns. A "fit the table to the window" button with
  wrapped cells was considered and not built: a sheet of ten columns squeezed into one screen
  wraps paths and descriptions into tall rows that are harder to read than a scroll.
- **The contributor, above the submission's table** (owner request, 2026-09-26: "It should be
  possible to see who it is (email) and how many contributions the contributor has done, and some
  statics about accepted and rejected"): "From dh@hinet.dk - [5] other submissions: [3] published,
  [1] waiting, [1] rejected", or "- no other submissions" for a first contribution. The server's
  `ContributorHistory` counts the contributor's OTHER submissions
  (`ISubmissionStore.GetContributorSubmissionsAsync` - by account, or by email, any case, among
  those sent without one) into CRT.Data's `ReviewContributorFacts`, sent as the detail's `contributor`; the
  maintainer app words it (`ReviewContributorLine`). Published = `merged`; a rejection counts only
  when a maintainer made it (`decided_by` set); a replaced (`withdrawn`) or never-finished
  submission is not counted.
- **The production panel says WHOSE WORK a promotion carries** (owner request, 2026-09-27: "I am not
  sure if the shown information in the right-side panel is any helpful ... can you propose
  something?"). It was the file copy list and nothing else - the mechanics of a copy, which cannot
  tell a maintainer whether the data is right, while the thing the button actually does (push named
  people's accepted work to every CRT user) was not on screen at all. The panel is now, in order:
  the merged submissions this carries (contributor, their own description, how long ago it was
  accepted), then what needs a second look (approval, removals, shared files), then the file list
  COLLAPSED behind a header that counts it - kept, because it is the audit trail.
  The facts are the server's `carrying` on the plan answer (CRT.Data's `CarriedSubmission`), built
  by `ProductionPromotionRules.Carrying` from **the same query and the same window** that
  `AfterPublishAsync` uses to email those contributors once the promotion succeeds - so the screen
  cannot name someone the mail will not reach, or stay silent about someone it will. It never fails
  the plan: context beside a decision, not part of it.
  **Deliberately NOT a diff of the data itself** (components added, schematics renamed): every
  submission was already reviewed in the table, and the tick is the maintainer saying they checked
  the board in BETA, so re-deriving it would re-ask an answered question at real cost.
- **"PUSH BACK TO QUEUE" IS BUILT** (owner decision, 2026-09-27). It was first judged impossible,
  and that judgement was WRONG on its central fact: the claim that a merged submission's blobs are
  garbage-collected. `SubmissionCollectionStates.Live` contains `Merged`, so
  `DeleteRetiredPayloadsAsync` never touches one - the project owner corrected it ("everything stays
  shadowed ... data is still there"), and proposed the mechanism: overwrite BETA from production.
  That works because production holds a COMPLETE board, not a partial set.

  `BetaRollbackFlow` (server) over `BetaRollbackPlan` (CRT.Data, pure) is the mirror of the
  promotion: production's bytes go back over BETA, files only BETA had are deleted, the system's
  recorded BETA state follows the tree, and every submission merged since the last promotion returns
  to `pending` with the maintainer's comment and its approvals cleared - all three in ONE store
  transaction (`RecordRollbackAsync`, the one writer of the BETA columns that is not a publish) -
  each contributor mailed. `POST /api/review/production/rollback`
  and `/rollback/plan`; the button sits beside "Publish to production" in CRT's red.

  **The limit that survives is structural: a rollback is PER SYSTEM.** `PublishMerge` replaces rows
  wholesale, so nothing records whose row was whose and one contributor's work cannot be picked out
  - `ProductionPromotionPlan` says the same in the other direction. So the confirmation NAMES every
  submission it takes back and says "ALL N submissions ... a board cannot be rolled back one
  contribution at a time". That sentence is the feature's safety: without it a maintainer returning
  one contributor's work would silently discard two others'.

  **A system never promoted is a REMOVAL, not a restore** (owner decision): production has nothing
  to restore from, so the board leaves the BETA tree entirely - "restore nothing" would leave the
  bad board exactly as it is while reporting success. **Shared files are deliberately never
  restored**: one reaches every board citing it, so rolling THIS board back would silently revert
  others. The comment is required on both sides - it is the contributor's only feedback.

  **Still not possible, and correctly so:** restoring a board whose submissions were merged BEFORE
  the last promotion. That work is already in production and live for everyone; a corrective
  submission is the only honest route.

  **Code review of the rollback (2026-09-27), eleven findings, all fixed.** The ones worth
  remembering: the BETA checksum manifest is rebuilt the moment the tree moves (it was left
  advertising the rolled-back bytes, so BETA clients kept them); the bookkeeping is ONE transaction
  (`RecordRollbackAsync`) that also CLEARS the returning submissions' approvals - left in place, a
  shared-file submission was republished by a single approval; a SHARED file goes back too, but
  only when BETA still holds exactly the bytes a returning submission carried, and a shared file it
  added goes through `UnusedFileRemover` (leaving the shared change in BETA leaked it to production
  with the next promotion of any board citing it); paths are compared ORDINALLY, as the
  case-sensitive server tree needs; every restore is a `VerifiedFileCopy`; comparisons check length
  first and cache hashes per file version, async; a failed record answers `NotRecordedMessage`
  instead of a 500, and pushing back again completes it; the contributor mail promises only what is
  true (the draft may already be retired; a newer submission does not always replace it); and CRT
  hears the rollback as `returned` ("Taken back out of BETA - waiting for review again"), with a
  merged receipt re-asked weekly past its 30-day window. **Known gap, left open:** a file in the
  rolled-back board's OWN folder that another board has since cited is deleted with the board's
  BETA-only files.
- **THE ADMINISTRATOR'S SECOND APPROVAL ONLY FOR REPLACING A SHARED FILE, AND NO AUTOMATIC
  REMOVAL OUTSIDE A SYSTEM'S OWN FOLDER** (owner decision, 2026-09-27: "change the review process
  slightly, so it will NOT require a second acknowledgement from the admin. BUT for this to work,
  then it should not delete any files in any of these folders ... It can delete any files inside
  its own system main folder - not outside it"). Asked and answered: a REPLACEMENT of an existing
  shared file (same path, different bytes) still needs the administrator too, because it changes
  what every board using it shows; adding a new shared file, and everything else, needs one
  approval. `SubmissionSharedFiles` counts replacements only (an unreadable tree still counts - the
  safe side), `ProductionPromotionPlan.TouchesSharedFiles` counts replaced shared copies only, and
  the approval recomputes from the tree now rather than trusting the flag stored at create (so
  submissions queued under the old rule are not held back). `AutomaticRemovalScope` limits every
  automatic removal - publish, promotion, push-back - to the system's own folder; shared or other
  systems' files a board stops citing become orphans for Admin > Unused files. A push-back no longer
  removes a shared file the returning submission added. Server 3.1.0.
- **ONE SUBMISSION IN BETA PER SYSTEM** (owner decision, 2026-09-27: "it should be possible only
  to submit ONE contributor submission to Beta per system ... Wouldn't this be optimal, or otherwise
  suggest better option"). It is the right rule given what a push-back can do: `PublishMerge`
  replaces a board's rows wholesale, so a push-back is per SYSTEM and returns everything merged since
  the last promotion - with two contributors' work in BETA it could only take both back. The price is
  throughput on a busy system (its queue waits for Beta > Prod), and reviewing, amending, requesting
  changes and rejecting all go on meanwhile. `ApprovePublishFlow` step 3b refuses any approval of a
  system `ProductionPromotionRules.IsAwaitingProduction` (409), with CRT.Data's
  `OneSubmissionInBeta.BusyMessage`; CRT Maintainer's `ApprovalGate` turns Approve off beside the same
  sentence, read from its "Beta > Prod" list. ON ONLY when production publishing is configured -
  without it nothing ever leaves BETA and every system would close after its first approval. The
  same gate turns Approve off for an UNPLACED new system (the server refused that since 2.0.0; the
  button now says so before it is pressed). Server 3.0.0 (MAJOR).
- **A SYSTEM'S HISTORY ON THE SYSTEMS SCREEN** (owner request, 2026-09-27: "I would like to see the
  date, newest first, to understand what has happened to a system"). `SystemHistoryRules` merges the
  system's submissions (sent; decided - merged, rejected, changes requested, replaced - and by whom,
  now that the per-system query reads `decided_by`) with the audit rows naming the system or one of
  its submissions (promotion, push-back, pool changes, invitations, amendments, placement), newest
  first, 100 at most; `SystemDetailAnswer.History` carries it and CRT Maintainer shows it, date first,
  in place of the bare submission list. Three events became audited for it: a saved placement
  (`system.placed`), an accepted invitation (now one row per system), and a removal's address.
- **"PLEASE WAIT" WHILE PUSHING BACK OR PUBLISHING TO PRODUCTION** (owner request, 2026-09-27):
  the main window fades and a layer takes every click until it ends, however it ends
  (`BetaView.RunBusyAsync`, `MaintainerMain.ShowBusy`).
- **SEVERAL MAINTAINERS PER SYSTEM** were always possible (the pool's key is system + account); the
  Systems panel now says so under its controls.
- **MAINTAINERS ARE SET ON THE SYSTEMS SCREEN, AND A NEW ONE IS INVITED BY EMAIL** (owner request,
  2026-09-27: "As admin I should be allowed to set a system maintainer in the 'Systems' list - so that
  should be moved from 'Admin' section. I should be able to either select an existing maintainer or
  invite a new maintainer via email."). "Set maintainers" left Admin; the Systems panel carries the
  controls for an administrator. Inviting needed an ACCOUNT PATH, because none existed - accounts
  were registered with curl, neither application can make one. So an invitation is also account
  creation: `POST /api/admin/maintainers/invite` stores the hash of a one-time code in migration
  0012's `maintainer_invitations` (a table of its own, so nothing reading the pool sees an invited
  person early) and mails the code; `POST /api/accounts/accept-invitation` (code, name, password),
  from "I have an invitation" on CRT Maintainer's sign-in screen, creates the account VERIFIED - the
  code proved the mailbox - and adds it to the pool of every system that address has an open
  invitation to, in one transaction. Hashes the password only after the code checks out (the
  reset's order), so the unauthenticated route cannot be used to run Argon2 without a real code.
  An address that has an account is refused (choose it from the list, which says why an account
  cannot be granted); inviting again replaces the open invitation; `.../invitations/withdraw`
  cancels one; codes last 14 days. Open invitations ride on `POST /api/review/systems/detail` for
  an administrator only. The buttons were also renamed: Systems, Contributor Submissions (Review),
  Beta > Prod (BETA), Admin, in that order. Server 2.1.0.
- **FOUR SCREENS INSTEAD OF THREE WINDOWS** (owner request, 2026-09-27). CRT Maintainer's top left is
  four buttons - **Review** (the queue, as it was), **BETA** (the "Publish to production" window),
  **Systems** (new) and **Admin** ("Set maintainers" and "Unused files", the administrator's two
  windows) - each a list on the left and the chosen item on the right. Review and BETA carry a badge
  counting the SYSTEMS that wait for this account (the server's `awaitsYou`: the queue's, and a new
  optional one on `GET /api/review/production`, false once this account has given its production
  approval for that BETA state); Systems a discreet count of all systems; Admin shows only for an
  administrator. Switching screen HIDES rather than closes, so an open table - unsaved changes
  included - is where it was on coming back. The BETA list and the systems are read with the
  queue's own minute check, each keeping the "never disturb what is open" rule: the tick on the
  shown BETA system survives a check unless its row changed.
  **The Systems screen is for EVERY maintainer, every system, contributors' addresses included**
  (owner decision, 2026-09-27, asked and answered "Everything for everyone"). That is a deliberate
  widening: until then a maintainer saw a contributor's address only on a submission they could
  decide. `GET /api/review/systems` and `POST /api/review/systems/detail` (`SystemOverviewFlow`,
  authority `CanReviewAnything`) give each system's maintainers, its contributors with how their
  submissions to it went (ContributorHistory's rules), and its 50 newest submissions in the word the
  contributor is told - which the maintainer app words with CRT's own `DescribeState`, so both apps
  describe a submission identically. Server 1.3.0; no migration.
- **A NEW SYSTEM'S PLACE IN THE DROP-DOWN LISTS** (owner request, 2026-09-27: "When a system is
  added to BETA, and it is a NEW system, can you then make sure it gets added also to the main Excel
  data file in the Data root? The maintainer should order the new system, so it becomes visible in
  the right location for the drop-down lists. This must be done before it can be pushed to BETA.").
  Asked and answered: the maintainer sets the POSITION and the NAMES (hardware name, board name,
  hardware notes); production gets the row "at the same place"; the place is chosen by DRAGGING the
  new system in the full list on the Systems screen.
  **The rule for the file:** `MasterListing` (CRT.Data) reads and writes the NEWEST master generation
  only (`NewestMasterPath` - older generations are frozen, as for boards), locates the sheet and the
  header by name in the first 20 rows, inserts ONE row after the named one (formatted like its
  neighbour) and changes nothing else; a blank hardware cell below the new row is given its own name
  so CRT's carry-forward does not move it under the new hardware. Written atomically.
  **Where it happens:** the placement is saved per system (`SystemListingFlow`, migration 0011's
  nullable `systems.listing_*` columns) and does not touch a file - unless the board is ALREADY in
  BETA (copied by hand, or merged before this existed), when BETA's file gets the row at once. The
  approval refuses a new-to-the-tree or unlisted system with no placement
  (`ListingForPublishAsync`, before anything is written, with the executor re-checking under the
  lock); the publish writes the row after the sidecar and before the database. Promotion inserts
  BETA's row into production's file after the nearest row above it that production lists
  (`TryResolvePlacement`), and refuses - with nothing copied - when neither file lists the system.
  Pushing a never-promoted system back out of BETA removes its row there; the placement is kept.
  **Unreadable masters:** a board already in the tree publishes exactly as before when the file
  cannot be read (the file was never part of publishing one); only a board NEW to the tree is refused.
  Promotion with an unreadable master copies as before and lists nothing.
  **No two systems under one name pair** (`MasterListing.NamesTakenBy`), at the placement, the
  approval, the publish and the promotion: CRT keys a board by "hardware name|board name",
  case-insensitively, so a second row with the same pair is one board twice sharing every setting.
  **The maintainer app:** `SystemPlacementView` on the Systems screen, dragged by CRT.UI's
  `ListRowDrag` - the Drafts tab's schematic-images drag, lifted out unchanged so both use one copy
  (frozen slots, re-entrancy guard, capture, edge auto-scroll). The Systems list leads with the
  systems waiting for a place, its badge counts those this account can place, and the review screen
  warns above an unplaced new system's table. Server 2.0.0 (MAJOR: the approval refuses what it
  accepted).
- **DRAFTS STAY UNTIL PRODUCTION, and a BETA tester is told when to switch back** (owner request,
  2026-09-27: "it will not remove anything from the users Draft tab, until the data has been
  migrated to production ... if the user has selected BETA as source ... there should be some kind
  of notification"). `DraftRetirement.IsPublishedState` is "published" only - "merged" (in BETA) no
  longer retires a draft, since "Push back to queue" can take it out of BETA again. When a receipt
  reaches "published" while CRT is downloading from the BETA source, a banner under the tabs
  (`Main.SourceSwitchNotice.cs`, words from `SubmissionReceiptPresenter.DescribeSourceSwitchNotice`)
  says the work is in production and to switch "Download data from test source" off; dismissing it
  is stored on the receipt (`SourceNoticeDismissed`), and it goes by itself when the source is
  switched. Retiring a draft now also removes the folders it leaves EMPTY above it
  (`DraftWorkbookStore` - never the Drafts root, never a folder with anything in it).
- **THE REVISION DATE IS THE SERVER'S, at both stages** (owner decision, 2026-09-26: "when the
  maintainer publish it to BETA, the revision date gets updated from server. Same happens when it
  gets published to real production, so server always wins, and what is typed by user is not
  important"). `ApprovePublishFlow.BuildPlan` stamps `BoardWorkbookStyle.FormatRevisionDate(nowUtc)`
  into the plan, and `PublishExecutor` writes THAT into the workbook rather than computing its own -
  one value, used by both. Production promotion copies the BETA bytes and touches nothing, so the
  BETA stamp carries through.

  This fixed two defects at once. A NEW SYSTEM could not be published at all: its seeded workbook
  carries no revision date (`DraftSeeder.CreateNewSystem` - nothing to inherit one from, and CRT
  never asks), `PublishMerge`'s fallback to the published board finds none, and `PublishPlan` then
  refused `publish.no-revision` at the one irreversible step (reported with a screenshot; it failed
  CLOSED, so nothing was corrupted). And on an EXISTING board the workbook and the database
  disagreed: the workbook had been stamped with the publish date since 2026-09-23 while the plan's
  descriptor - which is what reaches `systems.current_revision` - still carried the SUBMITTED date.
  That row is the base a contributor's next draft is diffed against, so the drift check was
  comparing against a revision no board ever held. Pinned by
  `ApprovePublishFlowTests.A_NEW_system_that_carries_no_revision_date_can_still_be_planned`, its
  `..._is_ignored_in_favour_of_the_publish_date` twin, and `PublishExecutorTests`' assertion that
  the workbook and the `systems` row now agree.

  **The client deliberately still stamps nothing.** A draft that stamped its own date would look
  newer than the board it came from and `DraftRevisionComparer` would report drift on every draft -
  see `PublishExecutor`'s header.
- **The decision buttons are coloured by direction, and a new system's tooltip stops saying
  "Published"** (owner requests, 2026-09-26). "Approve and publish to BETA" is CRT's green
  (`Button_Ok_*`; `MaintainerApp.axaml` defines them now - it had only `Button_Cancel_*`, and a
  missing `DynamicResource` draws an unstyled button in silence, so `SharedTableColourKeysTests`
  covers this application's own windows too), while "Request changes" and "Reject" stay red: "two
  red ones ... and then one accepting it, being green. Should be logical." Approve keeps its place,
  last and separated. Separately, a changed cell's tooltip is now TOLD what to call the value it
  replaced (`BoardTableDocument.Create`'s `baselineLabel`): a new system is compared with the
  submission itself, so its `HasBaseline` is true and the default "Published value: (empty)" named
  a board that does not exist - it reads "As submitted: (empty)" there, the file card's own word for
  that side. A draft names nothing and is unchanged.
- **A NEW SYSTEM gets no "Files included" line** (owner request, 2026-09-26: "if that is the
  component/board files, as typed in the table, then there is no reason to show this. Only data that
  is NOT otherwise visible should be shown here"). Every file of a new system is cited by a row that
  is itself on screen, so the count repeated the table's own image / local file / link cells. The
  line stays for a PUBLISHED board, where a file replaced under its own unchanged path moves no cell
  and so genuinely is invisible; a new system's KiCad data keeps its own line too, since no row
  cites it. `ReviewNotInTableTests` fails against the version that counted them.
- **Code review of the day's work (2026-09-26), ten findings, all fixed.** The ones worth
  remembering: the submission lines now count **what the approval would write** (`FilesLine` -
  a file replaced under its own path colours no cell, so it was approved unseen); the previous
  submission's detail is dropped on SELECTION, not when the next arrives, or the file card read the
  old submission's hashes while the table waited; signing out clears the queue, selection and
  detail, or the next sign-in re-selected the old one with the old account's badges; a window
  restored maximized no longer overwrites its remembered NORMAL position with the placement's own
  synthetic point (`thisRestoring`); `SubmissionKiCadFiles.Collect` keeps the PUBLISHED spelling of
  a case-variant KiCad file, so a publish replaces it instead of writing a second file beside it on
  the case-sensitive server; a conflict message is re-shown in the QUEUE's message when the refresh
  hid the decision panel; a dead session under unsaved table edits stops the checks and says the
  edits cannot be saved, rather than repainting the same dead end every minute; and files opened
  from the table are swept from the temp folder a day later (`ReviewTableFiles.OpenedFileLifetime`).
- **A board's KiCad data travels in its submission** (owner decision, 2026-09-26: "when a person
  submitting anything from his local PC, then I expect that it will send everything the server does
  not already have"). The `KiCad data` folder is the one part of a board no row cites - CRT reads it
  by name - so the rows-only manifest silently left it behind: a NEW system was published to BETA
  without its traces. `SubmissionKiCadFiles` is the one rule for it: only the types CRT reads
  (`ComponentListBuilder.IsSupportedKiCadRawFile`: .kicad_pcb/.kicad_sch/.kicad_pro - the shipped
  trees' stray KiCad-traces.json report and one legacy .sch are read by nothing and stay behind),
  only inside the board's OWN folder, exempt from "every file is cited by a row" at create AND in
  `PublishPlan` (`TryCheckName`'s `kiCadProjectFile`). Content signatures: s-expression text opening
  "(" for pcb/sch, JSON text opening "{" for pro (`SubmissionContentRules.OpensWith`). The client
  collects the UNION of the draft's and the synced official folder, draft winning
  (`SubmissionKiCadFiles.Collect`, `SubmissionManifestBuilder.Build`'s `kiCadFiles`), so a published
  board's untouched KiCad files travel as already-held (imported server-side, zero upload) and the
  manifest stays the complete intended state. *** An AMENDMENT carries them over ***
  (`AmendSubmissionFlow`): its file list is rebuilt from rows, which would silently strip them. The
  maintainer sees "KiCad data included: [3] files" (new/changed counts against the tree for a
  published board) above the table (`ReviewNotInTable.KiCadLine`, from the detail's
  `submittedFiles`). Publish writes them like any file; removals stay out (the folder is kept whole,
  the known limit above). With this, EVERYTHING in a draft the app reads now travels: rows, the
  sidecar's highlights and calibrations, every cited file (scope baselines included - rows cite
  them), and the KiCad folder; the draft marker and lock files are local bookkeeping, and an uncited
  stray file deliberately never travels.
- **A newer submission from the same contributor REPLACES the older one** (owner decision,
  2026-09-26: "if it is from same person, then the newest one always wins"). Every submission is
  the contributor's whole draft, and the draft stays after sending, so the newer one holds
  everything the older one did; left in the queue, the older one approved after the newer would
  publish the older content back over it. When a submission is QUEUED (finalise - never at create,
  so a send that never finishes replaces nothing), the server withdraws the same contributor's
  older submissions of the same system (`SubmissionReplacementRules`,
  `SubmissionFlows.WithdrawReplacedAsync`) - only those still `pending` and never amended: a
  maintainer's table edits
  or a first approval keep it ("IF this is the case, it is valid and it should probably stay").
  The amendment is checked inside the withdrawing transaction, under the row lock `AmendAsync`
  takes first (`ISubmissionStore.WithdrawReplacedAsync`). "Same contributor" is the same signed-in
  account, or the same contact email (any case). **The email is not verified**: anyone who knows a
  contributor's address and board can replace their waiting submission with their own - accepted
  by the project owner; the replacing one still has to pass review. No new state and no
  migration: the older one is `withdrawn` (nothing else sets that state) with a decision comment
  for the contributor, and CRT reads `withdrawn` as "Replaced by a newer submission", painted
  neutral rather than as refused. The maintainer app's "no longer in the queue" line names both
  causes. Pinned by `SubmissionReplacementTests`, one of which holds the server's state and CRT's
  wording together.
- **The hover card opens and closes AT ONCE, and sits beside its cell.** It is a `Popup` in the
  window's overlay layer (`ShouldUseOverlayLayer`), light dismiss off: the flyout in
  `TransientWithDismissOnPointerMoveAway` mode closed only ~100 px away from itself, and a
  light-dismissed popup spends the next click in the grid on closing. The card follows the pointer
  from file cell to file cell; beside the cell (not under it) it no longer covers the next row's
  file. A card the pointer has left decodes nothing, and decoding runs off the UI thread.

### Next in this phase

Nothing planned. Open owner decisions are listed under "Open questions" and in the orphan
section (CRT's own "Delete orphan and non-used files" setting).

### Definition of done (roles - met 2026-09-25)

- A maintainer can approve only the systems they are in the pool for, proven by a server-side denial
  test (`A_maintainer_of_ANOTHER_system_cannot_approve_and_NOTHING_is_written`).
- Two maintainers on one system can both act, and the audit trail names which one did
  (`SetDecisionAsync` records the account; the pool is a set).
- The administrator can act on every system without being added to any pool, and without a
  distinct override code path.
- A new system cannot be published by anyone but an administrator (its pool is empty), and a
  shared-files change is published only once a maintainer of the board AND the administrator have
  both approved it (`ApprovalRules`; the administrator alone when the board has no maintainers).
- Removing a maintainer takes effect immediately, proven by a test using the same access path a
  request takes.
- Every grant, revocation and decision is in the audit trail.

### Two approvals for a shared-file change [DONE 2026-09-25]

The project owner: *"in case of changes to any shared file, then both the maintainer and the admin
should approve before publishing to BETA or production. If there is no shared files changed, then
normal maintainer is sufficient."* The worked case that prompted it: a submission that edits a text
on the board AND replaces a shared image is ONE submission, so both approve that one submission -
it is not split.

- **`ApprovalRules` (CRT.Data, pure) is the rule, used for both stages.** `Required(touchesShared,
  systemHasMaintainers)`: an ordinary change needs any one approval (empty list, today's behaviour);
  a shared-file change needs `[Maintainer, Administrator]`, or `[Administrator]` alone when nobody
  reviews the board - the administrator does not approve twice. `Status(required, given, yourRole)`
  says who is still awaited, what this account may do and whether its approval publishes; the
  server sends that record to the maintainer app as-is (`approval` on the submission detail and on the
  production plan) and `ApprovalWording` writes the button and the line from it.
- **Either may approve first.** The first approval records a row (`submission_approvals` /
  `production_approvals`, migration 0008, keyed by role so one role cannot count twice) and
  publishes NOTHING; the submission's state becomes `approved`, it stays in the queue marked "one
  of two approvals given", and the other side is e-mailed (`EmailTemplates.ApprovalNeeded`). The
  second approval publishes, under the same `PublishLock` and with every check the first ran.
- **A production approval is tied to the BETA content hash it was given for.** If BETA changes
  between the two, the first approval no longer counts and both approve again - what was approved
  is what is copied.
- **Reject or Return by either one ends it.** An approval is agreement to publish, not a lock.
- **An administrator who is also in a board's pool approves as the administrator.** One account
  never supplies both halves.

### Traps

- **Do not build a permission matrix.** Two roles plus a per-system pool is the whole model. The
  authority question is always "is this account in this system's pool, or an administrator?"
- **Do not implement the administrator as a special case at each call site.** `ReviewAuthority` is
  the one place; a route that needs a new question adds it THERE and takes a `ReviewAccess`.
- **Do not reintroduce a role flag on the account.** `is_reviewer` was dropped precisely so no code
  path can grant the old global-queue behaviour by reading it.
- **A shared-file change needs both approvals.** `TouchesSharedFiles` is set at create and must be
  honoured by every rule that PUBLISHES; a new publishing route that forgets `ApprovalRules` lets
  one person change a file every board cites.
- **A system with an empty maintainer pool is normal**, not an error state.

---

### Draft discard notice [DONE 2026-09-28]

Owner request: a contributor who discards their own draft after submitting it must be "clearly visible
on the system and for the maintainer(s) - both in the BETA to PROD queue, but also in the normal queue",
so the maintainer can push it back and ask them (by email, for now). CRT's Discard reports each of that
board's still-open submissions to `POST /api/submissions/{id}/draft-discarded` (token-proved, no body;
CRT.Data `DraftDiscardContract`), marked on the receipt first so an offline discard goes at the next
launch (`DraftDiscardReporter`). The server keeps the first notice per submission
(`submission_draft_discards`, migration 0014, `DraftDiscardFlow`) and audits it as
`submission.draft_discarded` for the system's history. `draftDiscardedUtc` rides on the queue entry,
the detail's `submission`, the production plan's carried submissions and the Systems detail;
`carriesDiscardedDraft` on each production list entry. Server 3.3.0. The discard never withdraws the
submission, and every sentence says so. Not built: any channel back to the contributor.

### Hand-copied files, and the file tree [DONE 2026-09-28]

**Hand-copied files (server 3.3.1).** The project owner copies data between the trees as root; such
files refuse an open-for-write, and an approval answered 500 half-way through C128 (new workbook,
old `.json`). Board files are now REPLACED by rename (CRT.Data `FileReplacer`), and the publish, the
promotion and the push-back check every folder first (`TreeWriteAccess`), refusing with the
folders named and the fix command in the log. DEPLOYMENT.md step 3 says how to copy by hand.

**The file tree (server 3.4.0).** Owner request: "a file-structure for all existing files in BETA or
PROD, and then a highlighting of files changed ... 'show only changed files' ... including their
parent folder", and a way to open a system's folder and check "the Excel file ... and JSON". Both
CRT Maintainer queues draw the system's files as a tree (`FileTreeView`; production after the
publish on Beta > Prod, the BETA data after approving in "Files..."'s window, from the new
`GET /api/review/submissions/{id}/files`). Double-clicking a file opens it for every maintainer,
from the trees' public data addresses. The highlight file is now written only when its content
changes (`BoardSidecarWriter.WriteIfChanged`), so it no longer shows as replaced after a rows-only
change. Not built: a server-generated preview of the workbook a submission WOULD publish - before
approval the tree opens BETA's current workbook and says so; the new one is checked on Beta > Prod.

**Same day, after using it (server 3.5.0).** The owner's follow-up: the tree got the table's hover
card (the path and the picture or an "Open PDF file" link, nothing about the change), plus/minus
boxes on folders and "Expand all" / "Collapse all"; the administrator's "Open BETA folder" was
removed again ("with this new folder view you can scrap that"). Beta > Prod got a "Reject" beside
"Push back to queue" - the same rollback, the submissions `rejected` instead of `pending`. And a
defect the owner found by opening the workbook: every approval wrote it without the "# Hardware:" /
"# Board:" caption on each sheet, because the rows do not carry it and `PublishMerge` left it out;
it now keeps the replaced board's, or takes the drop-down names when there is none. CRT's own
"Save to draft" and label-editor save dropped it from drafts too - fixed, and guarded by
`BoardDataCaptionTests`.

## Phase 7 - Retire the PHP contribution path

**Goal.** New submissions arrive only through the new pipeline.

### Tasks

1. Confirm with the project owner that BETA has run on the new pipeline long enough to trust, and get
   explicit approval to enable Production.
2. Raise `$minimumContributionVersion` in
   [api/index.php](Webserver/app-contribution/api/index.php) to the first CRT release speaking the
   new contract. **Include any pre-release suffix** - `version_compare` ranks `2.7.0-beta.1` below
   `2.7.0`, so a bare version locks out every beta of it.
3. Leave the `OUTDATED_VERSION` response shape untouched - it is a contract with
   `ContributionPackaging.TryParseOutdatedVersionResponse` in the app.
4. Remove the old Contribute tab UI from CRT once the new path is the only one.
5. Archive `Assets/Webserver/app-contribution/` review code with a note saying what replaced it.
   Keep `app-feedback` and `app-checkin` - they are unrelated.
6. Update the Wiki: `Contribute-tab.md`, `Contribute-data-via-CRT.md`,
   `Contribute-data-via-GitHub.md` (likely removable), `Explanation-of-data-files.md`, plus any
   sidebar entry. Tell the project owner which files are ready to paste - **never** say a Wiki page has
   been updated.

### Definition of done

- Old-version submissions are rejected with the tailored update message.
- No supported CRT version posts to the old endpoint.
- Wiki files are updated in the repo and the project owner has been told which to paste.

---

## Phase 8 - Maintainer tab inside CRT [DONE 2026-09-29]

**Goal (owner decision, 2026-09-29).** No separate CRT Maintainer application: a maintainer uses CRT
itself, toggling to a Maintainer tab. One application, one version, one release. The detailed plan
the work followed is `Assets/MaintainerTabMergePlan.md`.

**Decided with the project owner:**

- **A "Maintainer" tab** in CRT's tab row, between Drafts and Configuration. While it is selected
  the sidebar and the worklog bar collapse, so the four screens (Systems, Contributor Submissions,
  Beta > Prod, Admin) get the window's width; leaving it puts them back.
- **Hidden unless "Enable Maintainer tab" is ticked** in the Configuration tab (default off, with
  a "?" to the new `Maintainer-tab` Wiki page). A hobbyist never sees a sign-in screen.
- **Sign-in inside the tab**, as the panel swap the application had. A remembered sign-in is
  restored the first time the tab is shown.
- **CRT is the UI reference**: the moved screens take CRT's theme and global styles.

**What moved where.** UI to `src/CRT.App/Tabs/Maintainer/` (namespace `CRT`; `MaintainerMain`
became `TabMaintainer`, a `UserControl`), logic to `src/CRT.App/Handlers/Maintainer/` (namespace
`Handlers.MaintainerHandling`), tests to `tests/CRT.App.Tests/Maintainer/` and
`tests/CRT.App.Tests/Ui/Maintainer/`. `Main.Maintainer.cs` holds the tab's visibility and the layout
while it is selected. `CRT.Data` did not change; `CRT.UI` gained one event
(`BoardTableEditor.OnlyChangesWantedChanged`).

**What the window did, and where it went:**

| Window behaviour | In the tab |
| --- | --- |
| Restore the remembered session in `OnOpened` | On the tab's FIRST attach (`TabMaintainer.Session.cs`); `ReviewSessionStore.Initialise()` at CRT's start-up |
| Minute queue check while the window is active | While the tab is attached AND CRT's window is active; on returning to either if due |
| `Closing` asks about unsaved table edits | `Main.OnWindowClosing` asks, after the Drafts tab's table, in one continuation |
| Its own "please wait" overlay | None; `Main`'s one overlay covers it |
| Window placement and "Show changes only" in `CRT-Maintainer-Settings.json` | Placement dropped (CRT's window has its own); "Show changes only" is `UserSettings.MaintainerShowChangesOnly`, carried over once by `MaintainerSettingsMigration`, which deletes the old file |

**What was retired:** the `src/CRT.Maintainer/` project and `tests/CRT.Maintainer.Tests/` (their
tests run in CRT.App.Tests), `build-and-release-maintainer.yml`, the release repository
`HovKlan-DH/Classic-Repair-Toolbox-Maintainer` and its `REVIEW_RELEASES_TOKEN` secret (both for the
project owner to archive and delete), the separate version (`1.0.0-alpha.4`), and
`MaintainerReleaseSeparationTests`. **The server** changed only its wording (3.5.1): the mails and one
refusal name CRT and its Maintainer tab, and the invitation mail tells the invitee how to show the tab.

**Traps for the next agent:**

- **Two base addresses, on purpose.** `AppConfig.CrtServerRootUrl` has no "/api" (the review routes
  append "/api/review/..."); `CrtServerBaseUrl` adds it (`SubmissionClient` appends only
  "/submissions"). Both are built from the one host.
- **The DPAPI entropy string still says "Review"** (`ReviewSessionProtection`). Changing it makes
  every remembered session undecryptable - it signs every maintainer out, silently.
- **The minute check is gated on ATTACHMENT**, not just on the window: a `TabControl` detaches an
  unselected tab's content, and CRT is in front far more often than the tab is on screen. Each
  check extends the session.
- **The exit prompt asks the Drafts table, then the Maintainer table, in ONE continuation** with one
  settled flag - two separate rounds would each need their own.
- **`LeftPanelWidth` must never be saved while the tab has collapsed the sidebar** - the splitter
  handler reads a width of 0 then.
- **`UserSettings.LoadFrom` on a missing file keeps the settings already in memory** - a test that
  wants defaults writes `{}` first.
- **`ReviewHighlightCanvas` is moved but used by nothing** (it drew for the change summary retired on
  2026-09-26). Left for the project owner to decide on.

---

## Security model

**Read this before implementing any part of Phase 3, 4, 5 or 6.** It is not a checklist to satisfy
at the end; it constrains how those phases are built.

### Secure by design

This project's stated principle is **secure by design**, which is a stronger commitment than the
more common "secure by default" and is worth stating precisely, because the difference decides how
code gets written:

- **Secure by default** means the system ships with safe settings. Insecure states are still
  reachable - someone just has to make a mistake to get there.
- **Secure by design** means insecure states are **structurally unreachable**. Security is a
  property of the architecture, not of remembering to configure it correctly.

The practical difference, in the one place it matters most:

| Secure by default | Secure by design |
| --- | --- |
| "Remember to authorise every endpoint" | Endpoints are deny-by-default; an unauthorised one cannot serve a request |
| "Do not put secrets in the client" | The client has no code path that could use one |
| "Reviewers should not publish" | The publish action did not exist in the (since retired) recommend-only Reviewer's authorisation model |
| "Validate uploaded files" | Nothing reaches the data tree except through the validating writer |
| "Do not write to Production yet" | The production path has no default, so a misconfigured service refuses to start |

**Apply these rules when building anything in Phases 3-6:**

1. **Deny by default.** Authorisation is opt-in per endpoint. A new endpoint with no explicit rule
   must refuse everyone, including the administrator. The failure mode of forgetting is an outage,
   never an exposure - an outage is noticed immediately and harms nobody.
2. **Make the dangerous thing impossible, not discouraged.** Where a capability should not exist for
   a role, remove the capability rather than hiding the button. See the retired recommend-only Reviewer role in
   [Phase 6](#phase-6---maintainers).
3. **One way in.** Every write to the data tree goes through a single validating writer. If a second
   path exists, it will eventually be the unvalidated one.
4. **Fail closed, always.** Unknown file extension, unparseable workbook, ambiguous permission,
   missing configuration: refuse. The existing code already works this way - keep it. See
   `ExternalTargetLauncher`'s allowlist and the orphan cleanup that refuses to delete when the
   referenced set cannot be determined.
5. **No ambient authority.** Nothing is permitted because of where the request came from, what
   happened earlier in the session, or what the client asserts about itself. Authority is looked up
   per request, from the database, for the specific object being acted on.
6. **Least privilege, structurally.** The service user can write the data tree and its blob store
   and nothing else. A maintainer's token is useless against systems they do not maintain. The
   retired recommend-only Reviewer's token could not change published data at all.
7. **Assume every input is hostile.** Contributed data, manifests, file names, images, board rows,
   API parameters. There is no trusted input in this system - not even from an administrator.

When a requirement and this principle conflict, raise it with the project owner rather than quietly
choosing convenience. The answer is sometimes "accept the risk", but that is the project owner's call
to make explicitly, and it belongs in this document when it happens.

### The governing rule

> **The Maintainer tab contains no authority. It is a rendering surface for decisions the server has
> already made.** (Written of the separate maintainer application; it holds unchanged for the tab
> it became on 2026-09-29.)

Every permission decision happens on the server, against the database, on every request. The
desktop app hiding a button is a convenience for honest users, never a control. If a single feature
is ever protected only by the client, the whole model collapses.

### Why open source does not weaken this

CRT and the maintainer app are public on GitHub, so an attacker can read exactly how access works. That
is fine, and it is the normal condition for security software (OpenSSL, SSH and Signal are all
public). The design must be secure **because of its structure**, not because attackers cannot see
it - Kerckhoffs's principle. Only keys are secret; the design is not.

Concretely, an attacker reading the source learns that `POST /api/submissions/{id}/approve` exists
and takes a bearer token. They could have learned that in minutes with an HTTP proxy against a
closed-source build. What they still need is a valid token for an account the server's own database
lists as a maintainer of that specific system. **No amount of source reading produces that**,
because the authority lives in MariaDB rows, not in the binary.

This yields a concrete test to apply to every endpoint built:

> If the attacker knows the full source, the exact request shape, and every system id, but has only
> an ordinary contributor account - can they do anything a contributor should not?

If the answer is ever yes, the endpoint is wrong.

**Two consequences an implementer must internalise:**

- **Never ship a secret in the client.** No API key, no shared password, no signing key, no
  admin-bypass token. Anything in the binary is public, whether or not the source is. If a design
  seems to need a client-side secret, the design is wrong.
- **Never trust a client-supplied claim of identity or role.** The token identifies the account;
  roles are looked up server-side from that account. A request saying "I am a maintainer" or
  carrying its own `role` field is ignored.

### Threat model

Ordered by real-world likelihood multiplied by damage - not by how alarming they sound.

| # | Threat | Likelihood | Worst outcome | Primary defence |
| --- | --- | --- | --- | --- |
| 1 | **Malicious contributed data** reaching users' machines | Moderate | Harmful files on thousands of installations | Upload-time content validation; the existing launcher allowlist |
| 2 | **Stolen maintainer account** | Moderate | Attacker publishes to every user of that system | Strong auth (2FA), scoped authority, audit feed, maintainer's own BETA backups. **NOT retained revisions - struck 2026-09-21, open question 5** |
| 3 | **Privilege escalation** by an ordinary contributor | Moderate | Approving their own or others' work | Server-side authorisation on every request |
| 4 | **Rogue maintainer** | Low | Same as 2, without the theft | Scope limits, audit trail, revocation. **Not rollback - struck 2026-09-21, open question 5**; recovery is a corrective publish plus backups |
| 4b | ~~**Rogue or stolen Reviewer**~~ | - | **Folded into 2 and 4 on 2026-09-25**: the recommend-only Reviewer role no longer exists; a Maintainer is the per-system publisher | See 2 and 4 |
| 5 | **Server compromise** via the new API | Low | Total | Small attack surface, localhost binding, no shell-outs |
| 6 | **Denial of service / disk exhaustion** by upload | Moderate | Service unavailable, disk full | Quotas, rate limits, blob garbage collection |

Note that threat 1 outranks everything. **The most valuable thing to attack here is not the review
app - it is the data pipeline into thousands of CRT installations.** Breaking into the maintainer app
is merely one route to that end; simply submitting hostile content is the cheaper route, and it
needs no account theft at all.

### Threat 1 - malicious contributed data (the most important)

Contributed files become files that thousands of installations download and open.
`Handlers/Security/ExternalTargetLauncher` already reasons about this on the consuming side: its
allowlist comment states that `.exe`/`.bat`/`.lnk` inside the "network-synced, community-contributed
data root must never" be handed to the shell, and anything not on the document/image allowlist is
rejected fail-closed.

**Extend that same reasoning to the upload side, where it is cheaper to catch:**

- Reject, at upload, any file whose extension is not on an explicit allowlist. Derive it from
  `ExternalTargetLauncher.OpenableFileExtensions` plus the image and KiCad types board data legitimately
  uses. **Allowlist, never blocklist** - fail closed on anything unrecognised, including files with
  no extension.
- **Verify file content, not just the name.** Check magic bytes: a `.png` must actually be a PNG.
  Decode every submitted image server-side; a file that fails to decode is rejected. This also
  catches decoder-exploit attempts aimed at the Avalonia/Skia image path on users' machines.
- Reject archives entirely unless a board format genuinely requires one. If one does, guard against
  zip-slip (entries escaping via `..`) and zip bombs (compression-ratio and total-size caps).
- Treat every path in a manifest as hostile: no absolute paths, no traversal, nothing resolving
  outside the target system folder. Reuse the reasoning in
  `OnlineServices.TryResolveValidatedLocalPath`. Test each case.
- Cap file size, file count and total system size per submission.
- Remember that board data is **rendered as HTML** in places (`OverviewHtmlBuilder`). Contributed
  text reaching an HTML surface must be encoded, or a contribution becomes script injection.
- Links in contributed data must stay confined to the schemes `ExternalTargetLauncher` already
  permits. Never widen that allowlist to accommodate a contribution.

### Threat 2 - stolen maintainer account (the most damaging)

Because maintainer approval publishes directly, **a stolen maintainer account is a software supply
chain attack on CRT's users.** This is the scenario to design against hardest.

- **Require two-factor authentication for any account holding maintainer or administrator rights.**
  TOTP is sufficient, needs no third party, and works offline. Contributors do not need it;
  maintainers do. Make this a precondition of being granted maintainership, not an option.
- Keep authority **narrow**: a maintainer's token must be useless against systems they do not
  maintain. Scope every check to the specific system id in the request. The retired recommend-only
  Reviewer's token had to be useless for publishing anything at all.
- **A pool raises the value of revocation, not the risk.** Several maintainers per system means more
  accounts that can publish to it, so removal must be immediate and a departing maintainer's tokens
  must stop working at once - not at next login.
- ~~Keep authority **reversible**: retained revisions (Phase 5) mean any publish can be rolled
  back.~~ **NO LONGER TRUE - open question 5, answered 2026-09-21.** Publishing overwrites in
  place and keeps no history, so a publish is NOT reversible by the system. Recovery is a
  corrective publish, or the project owner's own backup of the BETA tree. This was listed as what
  "makes direct publishing tolerable at all", so anything in Phase 6 that leaned on it needs
  re-arguing on the remaining safeguards: validation, review before publishing, and the
  administrator feed.
- Keep authority **visible**: the administrator feed (Phase 6) exists precisely so that an
  unexpected publish is noticed. A silent compromise is the dangerous one.
- Sessions expire; tokens are revocable; revoking maintainership takes effect immediately, not at
  next login.
- Alert on anomalies worth a human glance: a first-ever publish from a new location, a burst of
  approvals, an approval on a system the account has never touched.

### Threat 3 - privilege escalation

- Authorise **every** request server-side. Never infer authority from a prior request, a session
  flag set at login, or anything the client sends.
- Check the **object**, not just the verb: "may this account approve **this** submission, for
  **this** system?" - not "is this account a maintainer somewhere?" (Insecure Direct Object
  Reference is the classic failure here, and it is easy to write by accident.)
- Write a deliberate **negative test per protected endpoint**: an authenticated contributor, and a
  maintainer of a *different* system, must both be refused. Per project rules these tests ship with
  the endpoint, not afterwards.
- Administrator-only actions (appointing maintainers, and the administrator's half of a shared-file
  approval) are checked the same way, with no back door.

### Threat 5 - server compromise

- Bind the service to `127.0.0.1` only; reach it through the existing reverse proxy. It must never
  listen on a public interface.
- Run as a dedicated unprivileged user that can write only the data tree and its own blob store.
- Parameterised SQL everywhere - no string-built queries.
- **Never shell out** with any contributed value. No `Process.Start`, no image-conversion binaries
  invoked with a contributed filename.
- Keep dependencies few and patched; the .NET runtime on the server needs the same update
  discipline as the rest of the box.
- Log auth failures and authorisation denials. An attacker probing endpoints should be visible.

### Threat 6 - resource exhaustion

- Per-account quotas on submission count, size and rate.
- Rate limits on login, registration and password reset (these also blunt credential stuffing).
- Garbage-collect blobs from abandoned submissions - otherwise the disk fills quietly.
- Cap decompressed sizes before writing anything to disk.

### What is deliberately NOT relied upon

State these plainly so no future change quietly depends on them:

- **Obscurity.** The source is public; assume the attacker has read all of it.
- **Client-side checks.** Useful for user experience, worthless as controls.
- **Network location.** The current `review/.htaccess` IP restriction does not scale to maintainers
  and must not be reintroduced as a security boundary. It may remain as defence in depth for
  administrator-only endpoints, but nothing may *depend* on it.
- **Good intentions.** The model must hold when a maintainer account is hostile, because sooner or
  later one will be.

### Security review, 2026-09-25 - what was closed, and what is still open

A full read of the server, CRT.Data and the maintainer app, asking how an anonymous contributor, a
maintainer or an administrator could change, damage or overwrite the published data. What it found
and what was done, so the next session does not re-derive it:

| Finding | Fixed by |
| --- | --- |
| Submitted paths were contained to the DATA ROOT only, so a submission to one board could overwrite any non-workbook file of any other board (its highlight sidecar, its `system.json`, its PDFs) once approved | `SubmissionFileScope` + `SubmissionFileRules` (CRT.Data), at create AND in `PublishPlan`: a submission may change only its own folder and the two shared folders. Another board's file may be cited only byte-identical to the published copy, and is then never written |
| A file no row used was carried and published, and the maintainer app drew only images - so it was approved unseen | Every file must be cited by a row (`SubmissionFileRules`, `PublishPlan`). The maintainer app lists EVERY changed file (`ReviewFileComparison`) from new per-file facts the server sends (`SubmittedFileFact`, a CRT.Data type both ends share) |
| No file-type or content check (threat 1 asked for both); a dot-file such as `.htaccess` passed, and the data tree is under `public_html` with `AllowOverride All` | Allowlist of the types rows actually cite (`.png .jpg .jpeg .gif .bmp .webp .pdf .txt .html .htm`), no dot-segments, and a signature check at finalise and before publish (`SubmissionContentRules`). Apache hardening is DEPLOYMENT.md step 12a - **the project owner applies it by hand** |
| The system id was case-insensitive and the `systems` row is created by the first anonymous submission, so a case-variant could hijack a board's row and make every real submission fail | Migration 0005 (`utf8mb4_bin` in all three tables); case-variants of published paths and folders refused against the real tree (`PublishedTreeView.FindCaseVariant`) |
| No quota on anonymous submissions; completed blobs and payload rows were never collected | Per-address limit (`SubmissionRateLimitPolicy`: 20 a day, 4 GiB), a free-disk reserve (`BlobStore.HasRoomFor`), per-route body limits (`RequestBodyLimits`), and the hourly sweep now collects unreferenced blobs and ended submissions' rows |
| A blob was verified once, on arrival; an append racing completion could poison it, and publishing copied on trust | Per-upload lock in `BlobStore`; every blob re-verified before any write, and each copy hashed and renamed into place only on a match |
| `AccessTokenMinutes` promised a short-lived token that did not exist | Removed; the one token is the sliding session token (owner's 2026-09-22 decision) |
| Smaller: upload-state oracle for any hash; over-long summary/revision/finding subject failing an INSERT with a 500; `is_accepting` never read; no symlink check on publish | Each closed - see the section headers in `SubmissionFlows`, `SubmissionValidator`, `PublishPathSafety` |

**The check that shaped the rules.** `SubmissionRulesShippedDataTests` runs every new rule over every
board in `Assets/Data`. It found that C128DCR 250477 cites two texts in the C128 310378 folder (hence
"foreign, unchanged" is allowed) and that three shared images are misnamed - a PNG saved as `.jpg`,
two JPEGs saved as `.png` (hence an image name accepts any image signature). A stricter rule would
have been the third time this pipeline rejected published data.

**Still open, deliberately, each a owner decision:**

- **Images are signature-checked, not decoded.** Decoding on the server needs an imaging library
  there - a dependency and licence choice (threat 1 asked for it).
- ~~**Publishing never deletes a file.**~~ Resolved 2026-09-25: a publish removes the files the board
  stops citing that nothing else in that tree uses, from a list shown before approving and checked
  again at approval - see Phase 6, "Orphan files". A shared file another board cites is never removed.
- **No second factor, by decision (2026-09-25).** Phase 6's per-system scope IS built, and
  publishing is now open to maintainers on their systems - with a password as the only factor. The
  project owner chose to defer TOTP; the accepted risk is written into Phase 6. A stolen maintainer
  account publishes to that maintainer's systems - and, once publishing to production is switched
  on, to PRODUCTION for them. Mitigations: only bytes already in BETA, no shared file without the
  administrator's own approval too, an administrator mail on every maintainer's production publish,
  and the audit row.
- **Blobs of merged submissions are kept for ever**, so the next edit to a board uploads only what
  changed. Growth is bounded by what an administrator approves.

### Review checklist for each of Phases 3-6

Before declaring any of those phases done:

1. Does any client hold a secret? (Must be no.)
2. Is every protected endpoint authorised server-side, against the specific object?
3. Is there a negative test per protected endpoint - wrong role; right role but wrong system; and, while it
   existed, a recommend-only Reviewer attempting to publish?
4. Is all contributed content validated by extension **and** content before it is stored?
5. Is every action attributable in the audit trail?
6. ~~Can every publish be rolled back?~~ **No, by decision - open question 5, 2026-09-21.** The
   replacement question: **is a bad publish RECOVERABLE, and does the design say how?** Today the
   answer is a corrective publish plus the project owner's own BETA backups. Do not answer this one
   "yes" by quietly building a revision store; if rollback becomes necessary, re-open question 5
   with the project owner first.
7. Would the design still hold if the attacker knew everything except the passwords and tokens?
8. Is the insecure state **unreachable**, or merely **not the default**? (See
   [Secure by design](#secure-by-design). If it is only the latter, say so explicitly and get the
   project owner's agreement rather than leaving it implied.)

---

## Cross-cutting concerns

**Testing.** Every phase adds tests in the same change - a hard project rule, enforced by a `Stop`
hook. Pure logic (overlay merge, diff, manifest negotiation, validation, permissions) belongs in
`CRT.Data` or a `Handlers/` folder and must be unit tested. Never write a test needing hardware, a
network call, a display, or a spawned process; for the server, test against an in-memory or
temp-folder seam rather than the live box. Headless UI tests go through `UiTest.Run(...)`.

**Logging.** `CRT.Data` must not depend on the app's static `Logger`. Introduce a minimal `ICrtLog`
with a no-op default, adapted by each host. No test may call `Logger.Initialize()`.

**Security.** All input from a submission is untrusted: paths, file names, image bytes, board rows.
Reuse the existing validation reasoning in `OnlineServices` and `ExternalTargetLauncher` rather than
inventing new rules. Passwords hashed with Argon2id or bcrypt. Rate-limit auth endpoints. The
service listens on localhost only and is reached through the existing reverse proxy.

**Backups.** Before the first Production merge, confirm with the project owner that both the data trees
and the MariaDB database are backed up, and that a restore has actually been tested. The whole
design assumes published data is recoverable.

**Performance.** Board images run to tens of megabytes; C64 250407 alone is 76 MB across 1100 files.
Never load a whole system into memory to compute a diff. Hash streaming, upload streaming, and
prefer reading image headers over decoding images - `WorkbookPdfExporter.TryReadImageSize` already
demonstrates the technique and its subtleties.

**Documentation.** Wiki pages ship in the same commit as the code whose behaviour they describe.
`remind-wiki-mirror.sh` will name candidates; its `MAP` needs new entries as new code appears
(contribution, drafts, the Maintainer tab). Never claim a page is live - the project owner pastes them by hand.

---

## Open questions for the project owner

These need answers before the phases that depend on them. They are not blocking earlier work.

1. **Shared files ownership** (blocks Phase 6). `Commodore/Shared files` is 232 MB and
   `Generic shared files` 48 MB; neither belongs to one system, but contributions will want to add
   to them. Proposal: shared files are always administrator-owned, and a submission adding one is
   flagged for the administrator specifically. Confirm or redirect.
   **[ANSWERED 2026-09-25] Redirected: a shared-file change needs the board's maintainer AND the
   administrator**, for the BETA publish and again for production; with no shared file changed,
   one maintainer is enough. See Phase 6, "Two approvals for a shared-file change".

2. **Cross-system contributions** (affects Phase 4). A fact affecting five C64 revisions is five
   submissions, potentially to five maintainers. Acceptable, or should the app support a
   cross-system change explicitly?

3. **The legacy `Data <board>.xlsx` files** (affects Phase 5). Every board ships a legacy unversioned
   workbook beside the versioned one, which is why orphan cleanup almost never deletes anything. Is
   that duality still needed? Retiring it would simplify merging and make cleanup meaningful again.

   **Surveyed 2026-09-21, and the mechanism is not what this question assumes.** The numbers:
   13 versioned board workbooks, 9 unversioned, and **all 9 sit beside a versioned twin** - there
   is no board reachable only through a legacy file.

   **The duality is a WHOLE-TREE pairing, not a per-board one.** There are two MASTER workbooks,
   and each names its own generation of board files in its `ExcelDataFile` column:
   `Classic-Repair-Toolbox.v2.0.0.xlsx` names the 13 versioned ones,
   `Classic-Repair-Toolbox.xlsx` names the 9 unversioned ones. A board file is therefore never
   resolved by scanning for a version - `DataManager` picks the MASTER by version
   (`ResolveMainExcelFile`: highest version <= the app's own, else the unversioned one as
   fallback), and the board files follow from whichever master won.

   **So the legacy tree is a COMPATIBILITY TARGET for app builds older than 2.0.0**, not a
   redundant copy. Deleting it does not simplify a merge; it drops those builds. Two consequences
   for Phase 5 task 6: a merge must write into the generation its master names (and, if both are
   to keep working, into BOTH), and orphan cleanup is "almost never deleting anything" **because
   both generations are genuinely referenced** - it is working correctly, not failing.

   **[ANSWERED 2026-09-21] The merge writes ONLY the newest workbook generation, and never
   touches an older one.** The project owner's words: "I only want to have the merge write into the
   newest versionized Excel [...] for sure it should not merge into older Excel files."

   **An older generation is FROZEN, not stale.** `Classic-Repair-Toolbox.xlsx` and its 9
   unversioned board files still serve app builds older than 2.0.0, and they keep working
   precisely because nothing writes to them any more. So this is not "retire the legacy tree" -
   it is **leave it alone permanently**. Phase 5's merge must have no write path to it at all, and
   orphan cleanup must not treat its files as unreferenced.

   **The generation may itself be bumped for this work** (2.0.0 to 3.0.0, or 2.6.0 - the
   project owner has not decided which, and it is their call, not a derivation). Whatever the number,
   the rule is the same: publishing writes the newest generation only, and every older generation
   is immutable from that moment.

   **[ANSWERED 2026-09-21] The target generation is DISCOVERED, never configured.** The
   project owner: "You should always use the newest version, so if newest version is 2.0.0, you
   should use that - if it is 3.0.0 you should use that."

   So publishing resolves the newest master workbook present in the tree and writes that
   generation, exactly as `DataManager.ResolveMainExcelFile` already resolves one for reading.
   **This is deliberately not a setting.** A configured generation is a second place the truth
   lives, and the failure it produces is silent: the tree moves to 3.0.0, the setting still says
   2.0.0, and published contributions land in a frozen generation that no current build reads,
   with nothing throwing. Discovery cannot disagree with the tree because it reads the tree.

   Note the asymmetry with the READ path, and keep it: reading picks the newest generation **at or
   below the app's own version** (an app must not read a workbook newer than itself), while
   publishing picks the newest generation that exists, full stop. Bumping the generation is
   therefore a deliberate act - create the new master - and publishing follows it automatically
   from the next publish onward.

4. **Conflict policy** (affects Phase 5). Two contributors edit the same component simultaneously.
   Proposal: last-approved-wins plus a warning to the second maintainer, and nothing more until it
   demonstrably hurts. Confirm.

5. **Revision retention** (affects Phase 5). Keeping every published revision of every system costs
   disk - 1 GB today, hundreds of systems later. Keep all, or collapse to the last N?

   **Measured 2026-09-21: 999 MB across 23 systems**, the largest single ones being C64 250469 at
   109 MB, C128 310378 at 94 MB and C64 250407 at 76 MB; `Commodore/Shared files` accounts for
   210 MB of the total on its own.

   **The arithmetic that makes this decidable is that a revision is not a copy of a system.** A
   typical contribution is a handful of rows and perhaps one image - the workbook is kilobytes and
   the images are unchanged. If revisions are stored CONTENT-ADDRESSED, as Phase 4's blobs already
   are, an unchanged 76 MB of images is stored once however many revisions reference it, and a
   typo fix costs the size of one workbook. Under that storage the honest answer to "keep all or
   collapse to N" is **keep all** - the cost is roughly the size of what actually changed, which
   is what a contributor changed, which is small.

   It only becomes an N-revisions question if revisions are stored as whole-tree snapshots, and
   Phase 4 already built the mechanism that makes that unnecessary.

   **[ANSWERED 2026-09-21] NO PUBLISH HISTORY IS RETAINED. A published file is overwritten in
   place, and there is no rollback store.**

   **First, TWO DIFFERENT THINGS were being called a "revision", and the question above conflated
   them. The agent asked the project owner a question that mixed the two and had to re-ask.** Whoever
   reads this next must keep them apart:

   - a **workbook GENERATION** (no-version, `v2.0.0`, a future `v3.0.0`) - a compatibility target
     for a range of app builds. It is in the FILE NAME and the project owner sees it every day;
   - a **published REVISION** - `systems.current_revision` in `0001_initial.sql`, incremented on
     each publish and recorded by a submission as the `base_revision` it was built against. It
     lives in the DATABASE and never appears in the data tree at all.

   **They have OPPOSITE answers, which is exactly why mixing them was harmful:**

   - **Generations: never written, never deleted** (question 3). `Classic-Repair-Toolbox.xlsx`
     serves every app build from the first up to but NOT including 2.0.0;
     `Classic-Repair-Toolbox.v2.0.0.xlsx` serves 2.0.0 and newer. Each names its own board files
     internally. A new system entering the `v2.0.0` generation is
     `Data {Hardware} {Board} v2.0.0.xlsx`.
   - **Publish history: not retained.** Publishing overwrites the board workbook of the newest
     generation in place. Nothing archives what it said before.

   **What this removes from Phase 5:** task 7 ("retain every published revision so any merge can
   be rolled back") is **struck**. There is no revision store, no content-addressed archive of
   published trees, and no rollback action in the maintainer app. `systems.current_revision` still
   increments - it is what a submission diffs against, which is a different job from rollback and
   is still needed.

   **What carries the risk instead, and it must not be quietly weakened:** a bad merge is undone
   by publishing a correction, not by reverting. That puts the weight on automated validation
   (Phase 4 task 4) and on human review catching a problem BEFORE it publishes, plus the
   project owner's own backups of the BETA tree. Phase 6's "direct maintainer publishing" was argued
   partly on retained revisions making it safe; with no revisions retained, that argument is gone
   and the safety rests on review and backups. Revisit before widening who may publish.

6. **BETA to Production promotion** [ANSWERED 2026-09-20]. **A manual file copy the project owner
   performs.** There is no script, no rsync job and no automation to fit into. Two consequences,
   both already applied in Phase 3: the service must never write Production **by any code path**,
   and because it has no legitimate reason to, the service user is denied write permission on that
   tree at the filesystem level rather than merely being configured away from it. Do not design a
   promotion step, and do not re-ask this.

7. **`_UserContribution` vs `Drafts/`** [ANSWERED 2026-09-20]. **Drafts supersede it.** `Drafts/`
   is the only mechanism new authoring ever writes through; the existing `_UserContribution`
   sidecar stays **readable indefinitely** so boards already authored that way keep loading exactly
   as before, but nothing new ever creates one. Confirmed by the project owner - do not re-ask this in
   Session 2c, build task 10 on this basis directly. See
   [Phase 2](#phase-2---local-first-drafts-in-crt-done-2026-09-20) task 10.

8. **EPPlus licensing** [DECIDED 2026-09-21 - build the writer on EPPlus and proceed. The
   project owner will settle the licence later and has judged it not a current issue: "Use EPPlus for
   now, and I WILL take this later." It is a licensing action, not a code one, and it does NOT
   block Phase 5. Do not re-ask it, and do not swap the library on your own initiative].

   **The prediction below was wrong, and the reason is worth keeping.** It expected Phase 4 task 3
   ("server-side diff, using `CRT.Data`") to read workbooks. It does not, because the diff was
   built to work on the MANIFEST - the rows the client already sends as JSON - rather than on the
   published `.xlsx`. `SubmissionValidator`'s own header states the split: "it opens nothing and
   reads no images". That held right through Phase 4: `src/CRT.Server/` contained no reference to
   `EPPlus`, `OfficeOpenXml`, `EpplusLicense` or `BoardDataReader`, and the assembly sat inert as
   a transitive reference of `CRT.Data`.

   **[IT IS NOW REAL, 2026-09-21.] EPPlus EXECUTES ON THE SERVER from Phase 5**, in two places,
   neither of which is the submission diff the prediction pointed at:

   - **Reading** - `PublishedBoardReader` (task 3) opens the published workbook so a maintainer can
     be shown what actually changed. Only the server has the data tree, so only the server can
     read that half.
   - **Writing** - `BoardWorkbookWriter`, called by `PublishExecutor` (task 6). The first EPPlus
     WRITE anywhere in this codebase.

   `CRT.Server` was added to `CRT.Data`'s `InternalsVisibleTo` for the read half. The project owner
   has decided to proceed on the current licence and settle it separately - it is a licensing
   action, not a code one, and it does not block Phase 5.

   **`BoardDataWriter` NO LONGER EXISTS - checked 2026-09-21, and this is a trap for task 6.** It
   was retired during Phase 2 session 2c: authoring now produces draft ROWS
   (`LabelEditorDraftWriter`, `ComponentDraftWriter`, `KiCadCalibrationDraftWriter`) instead of
   mutating a workbook, because a shadow `.xlsx` under `Drafts/` would have been a second,
   inconsistent draft mechanism beside the `draft.json` one. Verified by grep: no `class
   BoardDataWriter` anywhere in `src/` or `tests/`, and **every surviving EPPlus call site is a
   READ** (`BoardDataReader`'s three `new ExcelPackage(stream)` calls and `DataManager`'s two).

   Two consequences for whoever builds publishing. First, **the workbook writer has to be written
   from scratch** - task 6 reads as though a writer is sitting there to be called, and it is not;
   budget for it. Second, **CLAUDE.md's coverage table still lists `BoardDataWriter` and the
   "deliberate quirks" note still cites its dead `region` argument** - both are stale and should
   go when someone is next editing that file for other reasons.

   The original deferral, kept for its reasoning:
   `SetNonCommercialPersonal` is used today for a desktop app. **Verified that Phase 3 executes no
   EPPlus code at all**: `EpplusLicense.Ensure()` is called only from `BoardDataReader`'s EPPlus
   entry points, never from a static constructor or module initializer, so a service that reads no
   workbook never invokes it. The assembly ships as a transitive reference of `CRT.Data` and sits
   inert. The project owner has decided to keep using EPPlus for now and revisit later, with replacing
   it a real possibility. **This becomes a genuine question at [Phase 4](#phase-4---submission-pipeline-done-2026-09-21)
   task 3**, where the server-side diff does read workbooks - confirm the licence covers server-side
   use, obtain one, or replace the dependency before that task.
