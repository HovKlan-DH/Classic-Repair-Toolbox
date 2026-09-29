# Plan: merge CRT Maintainer into CRT as a "Maintainer" tab

> **STATUS: ALL FOUR STEPS DONE (2026-09-29).** The outcome is recorded in Phase 8 of
> `Assets/NewContributeStrategy.md`, which is now the place to read. What is left is the project
> owner's own list at the end of Step 4. This file can be deleted.

This file's home in the repository is `Assets/MaintainerTabMergePlan.md` (owner decision,
2026-09-29). The executing agent reads it from there. It is a working document for this merge:
once Step 4 has recorded the outcome in `Assets/NewContributeStrategy.md` (Phase 8), the owner
may delete it.

Written for: the agent (Opus 5.5, high effort) executing this migration in the
`Classic-Repair-Toolbox` repository, one step per session. Read `.claude/CLAUDE.md` in full before
Step 1, and re-read the "Tests" section before every step - the Stop hook builds Release and runs
the whole suite, and a red suite blocks the handover.

## Context

Classic Repair Toolbox (CRT) ships two desktop applications built from this repository: CRT itself
(`src/CRT.App/`, every hobbyist) and CRT Maintainer (`src/CRT.Maintainer/`, a handful of
maintainers who review and publish contributions through `src/CRT.Server/`). Both are Avalonia
apps on the same shared libraries (`src/CRT.Data/` for board data and the review API contract,
`src/CRT.UI/` for the table editor and the busy overlay). The maintainer app has its own version
(`1.0.0-alpha.4`), its own release workflow, its own GitHub release repository, its own settings
file and its own copy of the headless test harness.

The project owner has decided (2026-09-29) that this split is not worth it: a maintainer is also a
CRT user, and the best experience is ONE application where they toggle between ordinary use and
maintainer work. The end state:

- **No CRT Maintainer application.** `src/CRT.Maintainer/`, `tests/CRT.Maintainer.Tests/`, its
  workflow and its release repository are retired.
- **One version**: the maintainer surface reports and ships with CRT's version.
- **The maintainer surface becomes a `Maintainer` TAB in CRT**, shown only when a Configuration
  checkbox is ticked, holding the sign-in and the four screens it has today.
- **CRT is the UI reference.** Where the two differ in look or behaviour, the maintainer surface
  adopts CRT's conventions, not the other way round.
- **What the app and the server share stays where it is** (`CRT.Data`, `CRT.UI`). Nothing moves
  into or out of the shared libraries because of this merge; only the maintainer app's OWN code
  moves, into `CRT.App`.

### Decisions taken with the project owner (2026-09-29) - do not relitigate

| Question | Decision |
| --- | --- |
| Where the maintainer surface lives | A **"Maintainer" tab** in `MainTabControl`. While it is selected, `Main` collapses the left sidebar (hardware/board/component list) and the worklog bar so the four screens get the full window width; selecting any other tab restores them. The four screen buttons (Systems / Contributor Submissions / Beta > Prod / Admin) stay as they are inside the tab. |
| How the tab becomes visible | A **Configuration checkbox** "Enable Maintainer tab" in a new "Maintainer" section, default OFF, persisted in `UserSettings`, with a "?" help icon - exactly the Oscilloscope/Workbooks pattern. Hobbyists never see a sign-in screen. |
| Where sign-in lives | **Inside the tab**, as a panel swap (sign-in panel until a session exists, then the screens) - what `MaintainerMain` does today. A remembered session signs in silently when the tab is first shown. |
| Tab position | **After Drafts, before Configuration.** Plain header text "Maintainer". |
| Version | The separate version is removed; the tab is CRT's version. |
| Old app | Hard cut. Maintainers uninstall CRT Maintainer and use CRT (owner communicates this; see the manual steps at the end). |

### Judgement calls made by the planner (routine; the owner can veto any of them)

- The Configuration "?" icon opens a NEW Wiki page, `Maintainer-tab`, added to `Assets/Wiki/`
  and the sidebar (every other tab has a page, and the Workbooks/MiniPro icons open pages). The
  owner pastes it like any other page.
- "Show changes only" is carried over from `CRT-Maintainer-Settings.json` once and that file is
  then deleted; the old window placement is not carried over (CRT's window has its own).
- The server's mails and one refusal message that name "CRT Maintainer" are reworded to CRT and
  its Maintainer tab, with a PATCH server version bump (`3.5.0` -> `3.5.1`).
- The review API client sends `User-Agent: CRT <version>` like CRT's other HTTP calls (it sent
  none before). Files opened from the file tree go under `%TEMP%\Classic-Repair-Toolbox\Maintainer\`
  instead of `%TEMP%\CRT Maintainer\`.
- Unticking "Enable Maintainer tab" hides the tab but neither signs out nor closes an open table
  (the "hide, never close" rule the four screens already follow); the exit prompt still catches
  unsaved edits.
- No badge on the "Maintainer" tab header; the minute queue check runs only while the tab is
  selected and the window is in front (the same "in front" rule as today, applied to the tab).

### What this plan deliberately does NOT do

- No change to `CRT.Server` routes, contracts or rules. The tab calls exactly the routes the app
  called. The ONE server change is wording: mails and one refusal message that tell a maintainer
  to "open CRT Maintainer" or download it from its releases page (see "The server knows the old
  app by name") - a PATCH bump of the server version, done in Step 3.
- No change to `CRT.Data`. `CRT.UI` changes in exactly two ways: the `InternalsVisibleTo` entries
  for the retired projects go, and `BoardTableEditor` gains one `internal` event
  (`OnlyChangesWantedChanged`, see "Settings") so the tab can persist "Show changes only".
- No redesign of the four screens. They move as they are; only what a Window had and a UserControl
  cannot (title, placement, its own BusyOverlay, `Closing`, `Activated`) is re-homed on `Main`.
- No badge on the "Maintainer" tab header (possible follow-up; the four buttons inside carry theirs).
- `CHANGELOG.md` is never touched (CLAUDE.md: hands off). Say in the final summary that an entry
  is warranted and leave it to the owner.

---

## The inventory (what exists today)

### `src/CRT.Maintainer/` - 60-odd source files, two folders

- **`Main/`** (UI): `Program.cs` (Velopack install hooks + Avalonia bootstrap), `MaintainerApp.axaml(.cs)`
  (resources: `FontAwesomeSolid/Regular` under the maintainer assembly's avares URI, the theme keys
  listed below, merges `avares://CRT.UI/BoardTable/BoardTableColors.axaml` and the ProDataGrid
  theme), `MaintainerMain.axaml` (sign-in panel + queue panel + its own `<ui:BusyOverlay>`) and its
  partials `MaintainerMain.axaml.cs` (sign in, queue, decisions), `.QueueItems.cs`, `.QueueRefresh.cs`,
  `.Table.cs`, `.Settings.cs`, `.Modes.cs`, `.Beta.cs`, `.Systems.cs`, `.Admin.cs`, `.Invitation.cs`,
  `.Files.cs`; the right-hand panels `BetaView`, `SystemView` (+ `SystemView.Maintainers.cs`),
  `SystemPlacementView`, `UnusedFilesView`, `FileTreeView` (+ `FileTreeView.FilePreview.cs`); the
  windows `FileTreeWindow`, `RollBackBetaWindow`; the helpers `ServerWait.cs` (BusyOverlay wrapper
  turning a timeout into `ReviewApiFailure.TimedOut`), `WindowMessage.cs` (one-line red/grey
  message), `FileTreeFiles.cs` (`IFileTreeFiles` host), `ReviewHighlightCanvas.cs`.
- **`Handlers/`** (pure, unit tested): `ReviewApiClient.cs`, `ReviewApiRoutes.cs`,
  `ReviewApiParser.cs` (+ `.Systems.cs`, `.Production.cs`), `ReviewApiViews.cs`, `ReviewSession.cs`,
  `ReviewSessionStore.cs`, `ReviewSessionProtection.cs`, `MaintainerSettings.cs` (also holds
  `WindowPlacementRules`), `MaintainerModes.cs`, `QueueRefreshRules.cs`, `ReviewQueueDisplay.cs`,
  `SystemsDisplay.cs`, `SystemPlacementDisplay.cs`, `ProductionDisplay.cs`, `UnusedFilesDisplay.cs`,
  `MaintainerAssignmentDisplay.cs`, `ReviewNotInTable.cs`, `ReviewContributorLine.cs`,
  `ReviewDecisionWording.cs`, `ReviewTableWording.cs`, `ApprovalGate.cs`, `ApprovalWording.cs`,
  `DraftDiscardWording.cs`, `FileRemovalWording.cs`, `MaintainerWaitWording.cs`, `FileTree.cs`,
  `FileTreePreviewSource.cs`, `OpenedFiles.cs`, `ReviewTableFiles.cs`, `ReviewTableFileSource.cs`,
  `ReviewFileComparison.cs`, `ReviewImageComparison.cs`, `ReviewHighlightGeometry.cs`,
  `ReviewScopeBaseline.cs`, `ReviewSummaryPresenter.cs`.
- Namespaces: `CRT.Maintainer` (UI) and `CRT.Maintainer.Handlers` (logic). CRT.App uses `CRT` for
  UI and `Handlers.<Area>Handling` for logic (`Handlers.DataHandling`, `Handlers.OnlineHandling`).
- `CRT.Maintainer.csproj`: AssemblyName `Classic-Repair-Toolbox-Maintainer`, version
  `1.0.0` / `1.0.0-alpha.4`, packages Avalonia + Fluent + Inter + `System.Security.Cryptography.ProtectedData`
  (DPAPI, Windows-only at runtime) + Velopack; links the same two Font Awesome .otf files CRT.App
  links; `InternalsVisibleTo CRT.Maintainer.Tests`; references CRT.Data and CRT.UI.
- Theme keys defined ONLY in `MaintainerApp.axaml` (Light and Dark): `Badge_NewSystem_Bg/Fg/Border`,
  `Queue_Divider`, `Mode_Selected_Bg/Border`, `Badge_Attention_Bg/Fg`, `Badge_Count_Bg/Fg`. Keys it
  defines that CRT's `App.axaml` ALREADY defines with the same meaning: `Bg`, `Fg`, `Table_Bg`,
  `Table_BorderRowLine`, `Text_Fail_Fg`, `Button_Cancel_*`, `Button_Ok_*`.
- Styles in `MaintainerMain.axaml`'s `Window.Styles`: `Border.Badge`, `ListBoxItem.QueueHeading`,
  `Border.QueueDivider`, `ListBoxItem.QueueEntry`, `ListBoxItem.ListEntry`, `Button.Mode`,
  `Button.Mode.Selected`, `Border.ModeBadge(.Attention/.Count)`.
- Per-app state on disk, all in `%LOCALAPPDATA%\Classic-Repair-Toolbox\` (CRT's own folder):
  `maintainer-session.json` (DPAPI-protected bearer token, `ReviewSessionStore`; entropy string
  "Classic-Repair-Toolbox.Review.Session.v1" MUST NOT change) and `CRT-Maintainer-Settings.json`
  (window placement + `ShowChangesOnly`, `MaintainerSettingsStore`).
- Window-only behaviour to re-home: `Title="CRT Maintainer"`; its own window placement
  (`MaintainerMain.Settings.cs`); `OnOpened` restoring the session; the minute queue check gated on
  `this.IsActive` and re-run on `Activated` (`MaintainerMain.QueueRefresh.cs`); `Closing` asking
  about unsaved table edits (`MaintainerMain.Table.cs`, `thisClosingSettled`); its own
  `<ui:BusyOverlay x:Name="BusyOverlay" Margin="-12">`.
- The base address: `ReviewApiRoutes.DefaultBaseAddress = "https://classic-repair-toolbox.dk"`
  (NO `/api`; every route appends its full path). CRT's `AppConfig.CrtServerBaseUrl` is
  `"https://classic-repair-toolbox.dk/api"` (WITH `/api`; `SubmissionClient` appends `/submissions`).
  The route tests pin that difference; keep both shapes, derive both from one host constant.
- `ReviewApiClient` sends no `User-Agent`. CRT's `OnlineServices`/`BoardViewReporter` send
  `"CRT {AppConfig.AppDisplayVersionString}"`.
- `OpenedFiles` (files opened from the file tree) writes to `%TEMP%\CRT Maintainer\<guid>\`.
- `MaintainerApp.axaml` is plain `<FluentTheme/>`; CRT's `App.axaml` restyles `Button`, `TextBlock`,
  `ListBox`, `TextBox`, `ComboBox` and `GridSplitter` globally and has a brown dark palette
  (`Bg` dark `#201C1A` vs the maintainer's `#1F1F1F`). Moved markup inherits all of that - which is
  the point ("CRT is the UI reference") - but it means every screen must be LOOKED AT after the move.
- **`CRT.App.csproj` is stricter than `CRT.Maintainer.csproj`:** warnings are errors in Debug AND
  Release, and `SonarAnalyzer.CSharp` runs with a hand-kept exemption list (`CRT.App.csproj` ~line
  52). Maintainer code has never been built under either. Expect Sonar findings on the moved code.
- The only type-name clash between the two apps is `Program`, which is not moved.

### The server knows the old app by name (user-visible text, so it changes too)

- `src/CRT.Server/Handlers/Email/EmailTemplates.cs:178-180`:
  `MaintainerDownloadUrl = "https://github.com/HovKlan-DH/Classic-Repair-Toolbox-Maintainer/releases"`,
  printed in the invitation mail (line ~215); "Open CRT Maintainer" in mail bodies at ~111, ~164,
  ~373, ~574. Pinned by `tests/CRT.Server.Tests/EmailTemplatesTests.cs:42,241`
  (`Assert.Contains("CRT Maintainer", ...)`) and `MaintainerInvitationFlowsTests.cs:89`
  (the releases URL) - both tests change WITH the wording.
- `src/CRT.Server/Handlers/.../ApprovePublishFlow.cs:408-415` `RemovalsNotSentMessage`: "your copy
  of CRT Maintainer did not send the list ... Update CRT Maintainer" (reused by
  `ProductionPromotionFlow.cs:230-235`) - the old-client detection message.
- Comment-only mentions in `AmendSubmissionFlow.cs:141`, migrations 0010/0012, `appsettings.Example.json:40`,
  `Configuration/ServerOptions.cs:50`, `CRT.UI.csproj:4-17`, `BoardTableColors.axaml:6`,
  `CRT.Data/ReviewApiContract.cs:12-24`, and many server/data test comments.
- Nothing on the server gates on client identity or version: no user-agent check, no minimum
  client version, no version header. The `User-Agent` is only stored truncated on the session row.

### `tests/CRT.Maintainer.Tests/`

References ONLY `CRT.Maintainer` (not CRT.Server - `ReviewWireContractTests` builds the server's
JSON settings itself via `ReviewApiContract.ApplyWireSettings`). Pure tests at the root
(`ApprovalGateTests` ... `UnusedFilesDisplayTests`, 35 files) and headless UI tests in `Ui/`
(`AnsweringHttpHandler.cs` - a fake `HttpMessageHandler`, `BetaViewTests`, `DraftDiscardShownTests`,
`FileTreeViewTests`, `MaintainerMainModesTests`, `MaintainerMainQueueTests`,
`MaintainerMainSettingsTests`, `MaintainerMainTableTests`, `RollBackBetaWindowTests`,
`SharedTableColourKeysTests`, `SystemPlacementViewTests`, `SystemViewTests`, `UnusedFilesViewTests`,
plus its own copies of `TestAppBuilder.cs` and `UiTest.cs`). The UI tests build `MaintainerMain`
but never `Show()` it (its `OnOpened` calls the live server) - where layout matters they move its
`Content` into a plain window.

### CRT.App integration points (all real names, verified)

- `src/CRT.App/Main/Main.axaml`: `RootGrid` (columns `200 / 4 / *`): `LeftPanel` (col 0),
  `MainSplitter` (col 1), `RightPanel` (col 2) with rows 0-5 = banners and `WorklogBar`, row 6 =
  `MainTabControl` (Schematics, Overview, Resources, `WorkbooksTabItem` [hidden at start],
  `OscilloscopeTabItem`, Contribute, `DraftsTabItem` [hidden at start], `ConfigurationTabItem`,
  Feedback, About). `DataSyncStatusIconBorder` overlays row 6 top-right. `ModeHintBorder` and
  `<ui:BusyOverlay x:Name="BusyOverlay">` are the last children of `RootGrid`.
- `src/CRT.App/Main/Main.axaml.cs`: `ApplyOscilloscopeTabVisibility()` (line ~606) and
  `ApplyWorklogBarVisibility()` (~661) are the two conditional-tab switches, both ending in
  `MoveSelectionOffHiddenTab(tabItem)`; `ShouldReturnFocusToComponentSearch()` (~968) is the
  focus-steal predicate (backs off on Feedback/Configuration/Workbooks and on Drafts while its
  table is open); `OnWindowClosing` (~1007) already asks about the Drafts table's unsaved edits via
  `TabDrafts.ConfirmLeavingTableAsync(this)` guarded by `_unsavedTableEditsSettledForExit`;
  `RootGrid.ColumnDefinitions[0].Width` is restored from `UserSettings.LeftPanelWidth` (~261) and
  saved by `OnMainSplitterPointerReleased` (~500).
- `src/CRT.App/Main/Main.Worklog.cs` `OnMainTabControlSelectionChanged` (~236): guarded by
  `ReferenceEquals(e.Source, MainTabControl)`; hands focus to Workbooks' search box and Drafts' table.
- `src/CRT.App/Tabs/Configuration/TabConfiguration.axaml(.cs)`: sections are a `TextBlock` heading +
  controls; the "Workbooks" section (axaml ~229-254) is the model: a `CheckBox` and a
  `Button.HelpIconButton` (Font Awesome Regular `&#xf059;`) in a horizontal StackPanel; the handler
  writes `UserSettings.X` then `if (TopLevel.GetTopLevel(this) is Main mainWindow) mainWindow.ApplyX();`
  (`OnEnableWorklogChanged` ~443); the help click goes through `ExternalTargetLauncher.TryOpen(AppConfig.WikiPageUrl(AppConfig.WikiPageWorkbooks))`.
- `src/CRT.App/Handlers/Data/UserSettings.cs`: property pattern `get => _data.X ?? default; set { _data.X = value; Logger.Info(...); Save(); }`
  over `UserSettingsData` (`[JsonPropertyName("camelCase")]`, `[JsonIgnore(WhenWritingNull)]`);
  `LoadFrom(path)` is the test seam; `AppConfig.AppFolderName = "Classic-Repair-Toolbox"`.
- `src/CRT.App/Main/App.axaml.cs`: `AppConfig` (constants: `CrtServerBaseUrl`, `WikiPage*`,
  `WikiPageUrl(name)`, `AppDisplayVersionString`, `AppFolderName`); `StartApplicationAsync` loads
  the stores in order (`UserSettings.Load()`, `WorklogManager.Load(args)`, `DraftManager.Load(args)`,
  `SubmissionReceiptStore.Load()`, `BoardViewReporter.Load()`) - `ReviewSessionStore.Initialise()`
  joins that list.
- `src/CRT.App/Main/App.axaml`: `FontAwesomeSolid`/`FontAwesomeRegular` keys (avares under
  `Classic-Repair-Toolbox`), theme dictionaries Light/Dark, merges `BoardTableColors.axaml` and the
  ProDataGrid theme - the maintainer's two includes are already there.
- `src/CRT.App/Tabs/Drafts/TabDrafts.Table.cs`: `IsTableOpen`, `HasUnsavedTableEdits`,
  `FocusTableIfOpen`, `ConfirmLeavingTableAsync(Window? owner)`, `AskAboutUnsavedTableEditsAsync`
  (+ `UnsavedTableEditsAnswerForTests`) - the shape `TabMaintainer` mirrors for `Main`.
- Tests: `tests/CRT.App.Tests/` (namespace `ClassicRepairToolbox.Tests` / `.Ui`) has the harness
  (`Ui/TestAppBuilder.cs` - `HeadlessTestApp : CRT.App` with an empty override, `Ui/UiTest.cs`,
  `[assembly: AvaloniaTestIsolation(PerAssembly)]`, the `"HeadlessUi"` collection,
  `xunit.runner.json` with `parallelizeTestCollections: false`), `TabConstructionTests` (every tab
  constructs), `Ui/SharedTableColourKeysTests` (scans ONLY `src/CRT.UI/BoardTable/*.axaml` today -
  the maintainer's copy had a second fact scanning `src/CRT.Maintainer/Main/*.axaml` that must be
  carried over against the new folder), `Ui/BusyOverlayHostsTests` (builds three CRT dialogs AND
  reads the markup of `Main.axaml`, `src/CRT.Maintainer/Main/MaintainerMain.axaml` and
  `src/CRT.Maintainer/Main/FileTreeWindow.axaml` by HARDCODED path for `<ui:BusyOverlay` and the
  `xmlns:ui` line), `WikiHelpPageNamesTests` (every `AppConfig.WikiPage*` names a file in
  `Assets/Wiki/`), `TestPathLiteralTests` (scans every `*.cs` under `tests/`), and
  `MaintainerReleaseSeparationTests` (reads both release workflows + AppConfig; asserts the
  maintainer releases go to their own repository and channels - its reason disappears with the app).
  `tests/CRT.Server.Tests/ReviewDecisionRulesTests.cs:337-348` and the maintainer's
  `ReviewDecisionWordingTests.cs:13-24` pin the minimum decision-comment length (10) from both
  sides without a project reference - that pairing survives the move unchanged.
- Release: `.github/workflows/build-and-release-maintainer.yml` (packId
  `Classic-Repair-Toolbox-Maintainer`, channels `maintainer-win`/`maintainer-linux`, tag
  `maintainer-v<version>`, publishes to `HovKlan-DH/Classic-Repair-Toolbox-Maintainer` with secret
  `REVIEW_RELEASES_TOKEN`, no macOS build); `build-and-release.yml` (CRT) does not mention the
  maintainer app; `build-and-unittest.yml` finds coverage reports with `find`, so 4 -> 3 test
  projects needs no functional edit, but its comment (~line 79) and `coverage-summary.sh` line 18
  say "four test projects". `.vscode/` and `.gitignore` have no maintainer entries. There is no
  `global.json`/`Directory.Build.props`. `Classic-Repair-Toolbox.slnx` lists all nine projects.
  `InternalsVisibleTo` entries naming `Classic-Repair-Toolbox-Maintainer` / `CRT.Maintainer.Tests`:
  `src/CRT.Data/CRT.Data.csproj` (~line 77) and `src/CRT.UI/CRT.UI.csproj` (~46-49).
- Docs that name the app: `.claude/CLAUDE.md` (many places), `Assets/NewContributeStrategy.md`
  (intro lines 5/16/22, the architecture diagram 114-142, Decisions rows 169 and 174, line 323,
  Phase 5 from 2142 with the release-workflow section ~2876-2921 and task 8 ~3414, Phase 6
  ~3626-3650, the security model 4549-4558, line 4793), `src/CRT.Server/DEPLOYMENT.md` (~111,
  1159-1180, 1267, 1300-1316, 1361-1378, 1458, 1553, 1558-1578, 1622), `src/CRT.Server/VERSION.md`
  (preamble lines 6 and 16 - the history table rows are history and stay),
  `.claude/hooks/server-version-bump.sh` (a comment, line 5). No Wiki page and no README text
  names the app; `Assets/Wiki/Contribute-data-via-CRT.md:445` says "the project owner may set you
  up as a maintainer".

---

## Target design (the shape every step builds towards)

### Files

```
src/CRT.App/Tabs/Maintainer/
    TabMaintainer.axaml(.cs)          the tab: sign-in panel + queue panel  (from MaintainerMain)
    TabMaintainer.QueueItems.cs       ... one partial per former MaintainerMain.*.cs part,
    TabMaintainer.QueueRefresh.cs         same names, same header comments, "MaintainerMain" -> "TabMaintainer"
    TabMaintainer.Table.cs
    TabMaintainer.Modes.cs
    TabMaintainer.Beta.cs
    TabMaintainer.Systems.cs
    TabMaintainer.Admin.cs
    TabMaintainer.Invitation.cs
    TabMaintainer.Files.cs
    TabMaintainer.Session.cs          NEW: remembered-session restore on first show, sign-in/out
                                      plumbing that OnOpened used to do (see "Lifecycle" below)
    BetaView.axaml(.cs), SystemView.axaml(.cs) + SystemView.Maintainers.cs,
    SystemPlacementView.axaml(.cs), UnusedFilesView.axaml(.cs),
    FileTreeView.axaml(.cs) + FileTreeView.FilePreview.cs, FileTreeWindow.axaml(.cs),
    RollBackBetaWindow.axaml(.cs), FileTreeFiles.cs, ReviewHighlightCanvas.cs,
    ServerWait.cs, WindowMessage.cs                (UI helpers, internal static, unchanged)

src/CRT.App/Handlers/Maintainer/      every file from src/CRT.Maintainer/Handlers/ EXCEPT
                                      MaintainerSettings.cs (retired - see "Settings")
                                      namespace Handlers.MaintainerHandling
src/CRT.App/Main/Main.Maintainer.cs   NEW partial: ApplyMaintainerTabVisibility, the sidebar/worklog-bar
                                      collapse while the tab is selected, the closing prompt hook
tests/CRT.App.Tests/Maintainer/       every pure test from tests/CRT.Maintainer.Tests/*.cs
tests/CRT.App.Tests/Ui/Maintainer/    every UI test from tests/CRT.Maintainer.Tests/Ui/ except the
                                      harness copies (TestAppBuilder.cs, UiTest.cs) and the
                                      maintainer SharedTableColourKeysTests (folded into CRT.App's)
```

Namespaces: UI files `namespace CRT` (every CRT tab is), logic `namespace Handlers.MaintainerHandling`
(the app's `Handlers.<Area>Handling` convention). `x:Class="CRT.TabMaintainer"` etc.;
`xmlns:local="clr-namespace:CRT"`. Tests: `ClassicRepairToolbox.Tests.Maintainer` (pure) and
`ClassicRepairToolbox.Tests.Ui.Maintainer` (headless UI), matching CRT.App.Tests' own namespaces.

Type names stay as they are (`ReviewApiClient`, `ReviewSession`, `SystemView`, ...) - only
`MaintainerMain` becomes `TabMaintainer`. Before moving, grep `src/CRT.App/` for each moved type
name to confirm no collision (none is expected; `FileTree`, `SystemView`, `BetaView` are not CRT
names today - verify, do not assume).

### Lifecycle of the tab (what a Window did, and where it goes now)

| Window behaviour | In the tab |
| --- | --- |
| `Title="CRT Maintainer"` | Dropped. The window is CRT's; the tab header reads "Maintainer". |
| `OnOpened`: `ReviewSessionStore.Initialise()`, `Recall`, show queue, read lists | `ReviewSessionStore.Initialise()` moves to `App.StartApplicationAsync` beside `SubmissionReceiptStore.Load()`. The tab restores the remembered session the FIRST time it is attached to the visual tree (`OnAttachedToVisualTree`, once - a `thisSessionRestored` flag), which is the first time it is selected: a remembered, usable session goes straight to the queue panel and reads the lists under the busy overlay; none shows the sign-in panel. Never in the constructor (tests build the tab; CLAUDE.md's "never Show MaintainerMain" rule becomes "never attach TabMaintainer to a shown window with `ReviewSessionStore` pointed at a real file" - tests never call `Initialise()`, so `Recall` finds no path and returns null). |
| Minute queue check while `IsActive`, re-run on `Activated` | The timer starts on sign-in and stops on sign-out as today. Its tick runs only when the tab is ATTACHED (a `TabControl` detaches an unselected tab's content - CLAUDE.md, Workbooks section) AND the owning window `IsActive` (`TopLevel.GetTopLevel(this) is Window { IsActive: true }`). `OnAttachedToVisualTree` (every time, i.e. every return to the tab) plays the role `Activated` played: run the check if `QueueRefreshRules.IsDue(thisQueueAskedUtc, now, ActivationGap)`. Subscribe to the window's `Activated` in `OnAttachedToVisualTree` and unsubscribe in `OnDetachedFromVisualTree` (keep the handler in a field so the unsubscribe is the same delegate). |
| `Closing` asks about unsaved table edits | `Main.OnWindowClosing` extends its existing Drafts block: `TabMaintainer.HasUnsavedTableEdits` is checked the same way, the Maintainer tab is selected, and `TabMaintainer.ConfirmLeavingTableAsync(this)` asked. ONE posted continuation asks Drafts first, then Maintainer, sets `_unsavedTableEditsSettledForExit` and calls `Close()` - two separate Post-and-Close rounds would each need their own settled flag. |
| Its own `<ui:BusyOverlay>` | Removed from the tab's markup. Every `BusyOverlay.For(this)` / `ServerWait.RunAsync(this, ...)` already resolves to the nearest host up the tree, which is `Main`'s overlay. `FileTreeWindow` keeps its own (it is a window). |
| Window placement (`MaintainerMain.Settings.cs`) | Dropped entirely - `Main` persists its own placement. `WindowPlacementRules` goes with it. |
| `UseSettings` applying `ShowChangesOnly` | `TableEditor.OnlyChanges = UserSettings.MaintainerShowChangesOnly` in the tab's constructor; written back whenever it changes (see "Settings"). |
| `SignInPanel` heading "CRT Maintainer" + "if you are here to repair a machine, use the ordinary CRT application instead" | Reword: drop the "CRT Maintainer" heading (the tab header already says it); keep the boxed explanation, now "For maintainers of the hardware data only." / "This tab reviews and publishes contributions to the shared hardware data. It needs a maintainer account, given by the administrator through an invitation, and does nothing without one." / "You can hide this tab again in the Configuration tab." - the third paragraph about "the ordinary application" is wrong now that this IS it. |
| `Window.Styles` | Become `UserControl.Styles` on `TabMaintainer.axaml`, unchanged. |
| `Grid Margin="12"` root | Keep a margin consistent with other tabs (TabDrafts/TabConfiguration use `Margin="16"` on their root - use 16). |

### `Main` while the Maintainer tab is selected (`Main.Maintainer.cs`)

- `ApplyMaintainerTabVisibility()`: `MaintainerTabItem.IsVisible = UserSettings.EnableMaintainerTab;`
  then `MoveSelectionOffHiddenTab(MaintainerTabItem)` when hidden - a copy of the shape of
  `ApplyOscilloscopeTabVisibility`. Called from the constructor beside the other two `Apply*`
  calls and from the Configuration checkbox handler. Hiding the tab NEVER signs out or closes an
  open table ("switching screen hides, it never closes" - the same rule the four screens follow);
  the closing prompt still catches unsaved edits.
- `OnMainTabControlSelectionChanged` gains: entering the Maintainer tab
  (`ReferenceEquals(MainTabControl.SelectedItem, MaintainerTabItem)`) -> `EnterMaintainerLayout()`;
  leaving it (`e.RemovedItems.Contains(MaintainerTabItem)`) -> `LeaveMaintainerLayout()`. The
  items ARE the `TabItem`s, so `RemovedItems` is the whole answer - no "previous tab" field, and
  it also covers the programmatic `MoveSelectionOffHiddenTab` path (a `SelectedItem =` assignment
  raises the event with `e.Source == MainTabControl`).
- `EnterMaintainerLayout()`: remember `RootGrid.ColumnDefinitions[0].Width`, set columns 0 and 1
  to `new GridLength(0)`, `LeftPanel.IsVisible = false`, `MainSplitter.IsVisible = false`,
  `WorklogBar.IsVisible = false`, and set a `thisMaintainerLayoutActive` flag.
  `LeaveMaintainerLayout()`: restore the remembered width (or `UserSettings.LeftPanelWidth`),
  column 1 back to 4, both visible again, `WorklogBar.IsVisible = UserSettings.EnableWorklog`,
  clear the flag. **Guard `OnMainSplitterPointerReleased`** so it never writes `UserSettings.LeftPanelWidth`
  while the flag is set (it reads `LeftPanel.Bounds.Width`, which would be 0). Change the row/column
  definitions IN PLACE - CLAUDE.md records that assigning a new `RowDefinitions`/`ColumnDefinitions`
  collection leaves the Grid measuring against the old ones.
- The banners in rows 0-4 (update, sync, source switch, ...) stay visible: they are app-wide.
  `DataSyncStatusIconBorder` stays (top-right of the tab area; the tab's mode buttons are top-left).
- `ShouldReturnFocusToComponentSearch()`: return false when `ReferenceEquals(selectedTab, MaintainerTabItem)`
  - unconditionally, not only while a table is open: the sign-in panel has the password box, the
  decision panel a comment box, and the Systems panel an invitation address box.
- F11 in `OnMainKeyDownCloseSinglePopup` (Main.ComponentPopup.cs:413-418) is NOT gated on any
  tab today - it toggles the schematics fullscreen window from anywhere. Gate it: do nothing while
  `MaintainerTabItem` is selected (a fullscreen schematic over a collapsed sidebar is nonsense).
  Escape there (:420-441) only acts when worklog entry mode is active or a component popup is
  visible, so it already reaches the decision comment box untouched - no change.
- The width restore `RootGrid.ColumnDefinitions[0].Width = new GridLength(UserSettings.LeftPanelWidth)`
  (Main.axaml.cs:261) runs AFTER the `Apply*Visibility` calls; `EnterMaintainerLayout` must never
  run from the constructor (it does not - selection starts on Schematics), so the remembered
  width is always the real one.
- `ApplyMaintainerTabVisibility` runs BEFORE the tab could ever be selected, so no layout state
  leaks: on start-up the selected tab is Schematics.

### Settings

- `UserSettings.EnableMaintainerTab` (bool, default false, JSON `enableMaintainerTab`) - the checkbox.
- `UserSettings.MaintainerShowChangesOnly` (bool, default false, JSON `maintainerShowChangesOnly`)
  replaces `MaintainerSettings.ShowChangesOnly`. Written by the tab whenever the table's
  "Show changes only" CHOICE changes. `BoardTableEditor` (CRT.UI) raises no event for that today
  (its events are `Saved`, `SaveRequested`, `ReloadedFromOutside`); `OnlyChangesWanted`/`OnlyChanges`
  are `internal` (visible to CRT.App through `InternalsVisibleTo`). Add ONE event,
  `internal event EventHandler? OnlyChangesWantedChanged`, raised in the `OnlyChanges` SETTER
  (`BoardTableEditor.axaml.cs` ~512-520, the only place `thisOnlyChangesWanted` changes; the
  check box handler `OnOnlyChangesChanged` routes through it) - NEVER in `ApplyOnlyChanges` or the
  `Attach` path, which change the FILTER without touching the choice, or a draft with nothing
  published would write "off" back as the user's choice. The Drafts tab ignores the event. Add a
  `BoardTableEditorTests` case: the event fires from the setter and not from `ApplyOnlyChanges`.
- One-time migration in Step 1: `MaintainerSettingsMigration.Apply(settingsFolder)` reads
  `CRT-Maintainer-Settings.json` if present, copies `ShowChangesOnly` into
  `UserSettings.MaintainerShowChangesOnly`, and deletes the file. Soft on every failure. Called from
  `App.StartApplicationAsync` right after `UserSettings.Load()`. Unit tested with a temp folder
  (`TempWorkspace`), never against the real AppData path. Window placement is NOT migrated (Main
  has its own).
- `ReviewSessionStore`: keep file name `maintainer-session.json`, keep `LegacySessionFileName`
  adoption, keep the DPAPI entropy string. Replace its private `AppFolderName` constant with
  `AppConfig.AppFolderName` (same value - one constant, not two). `Initialise()` is called once from
  `App.StartApplicationAsync`; `InitialiseAt(path)` stays the test seam. Tests that point it at a
  temp file mutate static state, so their class joins a `"ReviewSession"` xUnit collection
  (CLAUDE.md: "Any new test class that touches one of these statics must join its collection").
- `AppConfig`: add `CrtServerRootUrl = "https://classic-repair-toolbox.dk"` and define
  `CrtServerBaseUrl = CrtServerRootUrl + "/api"` (a constant expression). `ReviewApiRoutes.DefaultBaseAddress`
  becomes `AppConfig.CrtServerRootUrl` (still `const`, still no `/api`; `ReviewApiRoutesTests`'s
  "no doubled /api" test keeps pinning it). Add `WikiPageMaintainer = "Maintainer-tab"` for the
  help icon (the Wiki page is written in Step 4; `WikiHelpPageNamesTests` fails until the file
  exists - so create the file in Step 2 with a placeholder heading and fill it in Step 4, or write
  it fully in Step 2; either way the test is green at the end of Step 2).
- `ReviewApiClient`: send `User-Agent: CRT {AppConfig.AppDisplayVersionString}` like
  `OnlineServices` does (the server log then names the client version; nothing on the server reads
  it). One line in its HttpClient construction; pin with a test using `AnsweringHttpHandler`.

### Theme

Add to BOTH theme dictionaries of `src/CRT.App/Main/App.axaml`, with the exact values from
`MaintainerApp.axaml`, under a comment naming the tab: `Badge_NewSystem_Bg/Fg/Border`,
`Queue_Divider`, `Mode_Selected_Bg/Border`, `Badge_Attention_Bg/Fg`, `Badge_Count_Bg/Fg`. Do not
add them to `UserSettings.UserThemeDefaultColors` (the JSON user-preference theme) unless the owner
asks; the "UserPreference" variant sets the Light variant and then overwrites ONLY the keys in
`UserSettings.GetUserThemeColors()` (App.axaml.cs:351-394, seeded from `UserThemeDefaultColors`),
so a key absent from that table simply keeps its Light value - verified. CRT.App's
`SharedTableColourKeysTests` scans ONLY CRT.UI's table markup today; the maintainer's copy had a
second fact scanning the maintainer app's own `*.axaml` for every `{DynamicResource X}` and
resolving each in both variants (asserting `Button_Ok_Bg` and `Button_Cancel_Bg` were found). In
Step 2 that fact moves into CRT.App's test, pointed at `src/CRT.App/Tabs/Maintainer/*.axaml`, and
also asserts `Badge_Attention_Bg` and `Mode_Selected_Bg` were found - that is what catches a
theme key left behind in the deleted `MaintainerApp.axaml`.

---

## Steps

Each step is one session. Each ends with `dotnet test Classic-Repair-Toolbox.slnx` green in
Release (the Stop hook enforces it) and a summary stating real counts. Do not commit or push - the
owner commits (memory: never auto-commit). **Steps 1 and 2 deliberately leave `src/CRT.Maintainer/`
compiling on its own copy of the code; the duplication is intended and is removed in Step 3.** Do
not try to make CRT.Maintainer consume CRT.App's copy, and do not delete anything from
`src/CRT.Maintainer/` before Step 3.

Conventions that apply to every step (from CLAUDE.md): every partial file opens with a header
comment stating what it owns and the `.axaml.cs` part carries the file map; pure logic lives in
`Handlers/`, never in a tab; tests are part of the change; plain ASCII punctuation in everything
written; user-facing text never assumes a job or income; `Logger.*` calls are inert in tests and no
test may call `Logger.Initialize()`, `UserSettings.Load()`, `ReviewSessionStore.Initialise()` or
`DataManager.InitializeAsync()`.

### Step 1 - The logic and its tests move into CRT.App (CRT.Maintainer untouched)

Goal: `src/CRT.App/Handlers/Maintainer/` holds every pure class, `tests/CRT.App.Tests/Maintainer/`
every pure test, both green, and the new settings/constants/theme keys exist. No UI yet.

1. Copy (`git cp` does not exist - copy the files, keep `src/CRT.Maintainer/` as is) every file of
   `src/CRT.Maintainer/Handlers/` except `MaintainerSettings.cs` into
   `src/CRT.App/Handlers/Maintainer/`. Change `namespace CRT.Maintainer.Handlers` to
   `namespace Handlers.MaintainerHandling`; fix `using`s. Header comments that say "this app" /
   "the maintainer app" / "CRT Maintainer" / "MaintainerMain" are reworded to "the Maintainer tab" /
   "TabMaintainer" - read each header, do not sed blindly, since several explain history that
   should stay ("the app became CRT Maintainer on 2026-09-25").
2. `ReviewSessionStore`: use `AppConfig.AppFolderName`; nothing else changes. `ReviewSessionProtection`:
   unchanged (entropy string stays). `MaintainerWaitWording`: unchanged (CRT's `CrtWaitWording` is
   for CRT's own waits; different sentences, keep both). `OpenedFiles`: the temp folder becomes
   `%TEMP%\Classic-Repair-Toolbox\Maintainer\<guid>\` (the app's own name; update its test).
   **Sonar and warnings-as-errors:** `CRT.App.csproj` fails the build on warnings in Debug
   (lines ~170-172) as well as Release (~23) and runs `SonarAnalyzer.CSharp`; the exemption list
   is the `<WarningsNotAsErrors>` property at ~line 52 (rule IDs `S108;S125;...`), with its
   rationale in the comment above it: a rule ID already in the list is exempt; a rule that fires
   for the first time must be FIXED in the code or its ID appended with the justification stated
   in the turn summary. Build after copying and fix every finding in the code (unused parameter,
   member that could be static, naming, etc.); append an ID only for a rule the project has
   already decided is noise. Say in the summary which findings were fixed and which IDs appended.
3. `ReviewApiRoutes.DefaultBaseAddress` = `AppConfig.CrtServerRootUrl` (add that constant and
   re-express `CrtServerBaseUrl` from it - see "Settings" above). `ReviewApiClient`: add the
   `User-Agent` header.
4. `UserSettings`: add `EnableMaintainerTab` and `MaintainerShowChangesOnly` (pattern: copy
   `EnableWorklog`). Add `MaintainerSettingsMigration` (Handlers/Maintainer) and call it from
   `App.StartApplicationAsync` after `UserSettings.Load()`; call `ReviewSessionStore.Initialise()`
   beside `SubmissionReceiptStore.Load()`. `AppConfig.WikiPageMaintainer` is added in Step 2 with
   its page, not here.
5. `App.axaml`: add the theme keys (both variants). `CRT.App.csproj`: add
   `System.Security.Cryptography.ProtectedData` (same version as CRT.Maintainer.csproj) with the
   same comment about DPAPI being Windows-only and guarded; `TreatWarningsAsErrors` in Release
   means the `[SupportedOSPlatform("windows")]` split in `ReviewSessionProtection` must survive
   the move exactly (CA1416).
6. Tests: copy every root `*.cs` of `tests/CRT.Maintainer.Tests/` except `MaintainerSettingsTests.cs`
   into `tests/CRT.App.Tests/Maintainer/`, namespace adjusted, `using Handlers.MaintainerHandling`.
   `ReviewSessionStoreTests` joins `[Collection("ReviewSession")]` (define the collection). From
   `MaintainerSettingsTests`, keep whatever tested `ShowChangesOnly` as a `UserSettingsTests` case
   (`LoadFrom` a temp file, round-trip both new properties) and add `MaintainerSettingsMigrationTests`
   (file present -> value copied and file deleted; absent -> nothing; unreadable -> nothing, no
   throw). Drop the `WindowPlacementRules` tests with the class. Add a `ReviewApiClient` test that
   the request carries the `User-Agent` (through `AnsweringHttpHandler`, copied to
   `tests/CRT.App.Tests/Ui/Maintainer/` now or in Step 2 - it is a plain `HttpMessageHandler`, so
   `tests/CRT.App.Tests/Maintainer/` is the better home; move it there).
   `TestPathLiteralTests` scans the moved files automatically - fix any literal it flags rather
   than exempting it.
7. Run the full suite. Expected: everything that was green in CRT.Maintainer.Tests is green again
   under CRT.App.Tests (count them: the pure tests number about 200 - report the real figure), plus
   the new ones. CRT.Maintainer and CRT.Maintainer.Tests still build and pass unchanged.

### Step 2 - The Maintainer tab, wired into Main and Configuration (CRT.Maintainer still present)

> **Carried over from Step 1 (done 2026-09-29):** `PoolActionTests.cs` was NOT moved in Step 1,
> because `PoolAction`/`PoolActionKind` live in `src/CRT.Maintainer/Main/SystemView.Maintainers.cs`
> (UI). Copy it into `tests/CRT.App.Tests/Maintainer/` together with `SystemView`. The session
> store's tests already had their own collection, `[Collection("ReviewSessionStore")]` - that name
> was kept rather than the `"ReviewSession"` this plan suggested. `AnsweringHttpHandler.cs` now
> lives in `tests/CRT.App.Tests/Maintainer/` (namespace `ClassicRepairToolbox.Tests.Maintainer`);
> the UI tests use it from there.

Goal: ticking "Enable Maintainer tab" in Configuration shows a working Maintainer tab that signs in,
restores a remembered session, shows the four screens, opens tables, decides, and asks about
unsaved edits on exit; CRT.App's headless UI tests cover it.

1. Create `src/CRT.App/Tabs/Maintainer/` and copy every UI file from `src/CRT.Maintainer/Main/`
   except `Program.cs` and `MaintainerApp.axaml(.cs)`. Rename `MaintainerMain.*` to
   `TabMaintainer.*`; `Window` -> `UserControl` in markup and code-behind (`x:Class="CRT.TabMaintainer"`,
   `xmlns:local="clr-namespace:CRT"`, drop `Title/Width/Height/MinWidth/MinHeight`); namespaces to
   `CRT`; `using Handlers.MaintainerHandling`. Move `Window.Styles` to `UserControl.Styles`. Remove
   the tab's own `<ui:BusyOverlay>`; keep `FileTreeWindow`'s. Keep the `FindControl`-by-name
   lookups as they are; DELETE the hand-written `private void InitializeComponent() => AvaloniaXamlLoader.Load(this);`
   (`MaintainerMain.axaml.cs:76`) so the generated one is used like every other CRT tab (the two
   overloads must not coexist). The complete list of Window-only members to re-home, from a
   verified sweep of `src/CRT.Maintainer/Main/`: `OnOpened` (axaml.cs:92) -> the Session partial;
   `this.IsActive` / `this.Activated` (QueueRefresh.cs:45,49) -> the attach/window gate;
   `this.Closing += OnWindowClosing`, `OnWindowClosing`, `this.Close()`, `thisClosingSettled`
   (Table.cs:76,381,391,60) -> gone, replaced by the two members `Main` calls;
   `prompt.ShowDialog<...>(this)` (Table.cs:297) and `window.Show(this)` (Files.cs:57) -> owner
   `TopLevel.GetTopLevel(this) as Window` (null-guarded; `BetaView.axaml.cs:449-481` already does
   exactly this for `RollBackBetaWindow` - copy it); `this.Launcher.LaunchFileInfoAsync(...)`
   (Table.cs:369) -> `TopLevel.GetTopLevel(this)?.Launcher` (as `FileTreeFiles.cs:116` already
   does); everything in `MaintainerMain.Settings.cs` (dropped). The views (`SystemView`,
   `SystemPlacementView`, `UnusedFilesView`, `FileTreeView`) use no Window member. The ~60
   `x:Name`s clash with no member name in the partials, so the generated fields are safe.
   **Look at every screen after the move** (Step 2.6): CRT's global `Button`, `ListBox`,
   `ListBoxItem`, `TextBlock` and `TextBox` styles (App.axaml ~452-565) now apply. The three
   places most likely to need a more specific selector in the tab's own styles: `Button.Mode`
   (an unselected mode button takes `Button_Bg/Fg/Border`; the selected one's foreground stays
   `Button_Fg` over `Mode_Selected_Bg` - check contrast in both variants), the queue headings
   (`ListBoxItem.QueueHeading:disabled` draws `Fg` while entries draw `Form_Fg` over the list's
   `Form_Bg` - the heading/row distinction now rests on those differing), and a selected queue row
   (the row's built `TextBlock`s take the global `Fg` over `Form_Selected_Bg`).
2. Apply the "Lifecycle" table above: `TabMaintainer.Session.cs` (first-attach session restore),
   `TabMaintainer.QueueRefresh.cs` (attached + window-active gate, `Activated` subscription per
   attach - keeping the existing `thisQueueTimer.IsEnabled && QueueRefreshRules.IsDue(...)` guard
   on BOTH the `Activated` handler and the on-attach check, so a return to the tab while signed
   out fires nothing; preserve the expiry funnel in `MaintainerMain.axaml.cs:464-605` exactly:
   an expired session with unsaved table edits stops the checks and says so without leaving the
   panel, everything else goes through `ShowSignInPanel`), no `Closing` in the tab - instead `HasUnsavedTableEdits` and
   `ConfirmLeavingTableAsync(Window owner)` exposed for `Main` (mirror `TabDrafts`' two members,
   including an `UnsavedTableEditsAnswerForTests` seam if the maintainer's tests need one).
   `MaintainerMain.Settings.cs` is NOT copied; its `ShowChangesOnly` half becomes the
   `UserSettings.MaintainerShowChangesOnly` read in the constructor and the write-back described
   under "Settings". Reword the sign-in explanation box as the table says.
3. `Main.axaml`: add `<TabItem x:Name="MaintainerTabItem" Header="Maintainer" IsVisible="False">`
   with `<local:TabMaintainer x:Name="TabMaintainer" />` between `DraftsTabItem` and
   `ConfigurationTabItem`, with a comment in the style of the Workbooks/Drafts ones. New partial
   `Main.Maintainer.cs` with the header and the members described under "`Main` while the
   Maintainer tab is selected"; add it to the file map in `Main.axaml.cs`'s header. Wire:
   `ApplyMaintainerTabVisibility()` in the constructor after `ApplyWorklogBarVisibility()`;
   `TabMaintainer.Initialize(this)` beside the other tabs' `Initialize` calls if the tab needs
   `Main` (it should not - it reaches the window through `TopLevel.GetTopLevel`);
   `OnMainTabControlSelectionChanged` enter/leave; `ShouldReturnFocusToComponentSearch` opt-out;
   `OnWindowClosing` extended; `OnMainSplitterPointerReleased` guarded; F11/Escape verified.
4. `TabConfiguration.axaml`: a "Maintainer" section after the "Drafts" section, modelled on the
   "Workbooks" one: `<CheckBox x:Name="EnableMaintainerTabCheckBox" Content="Enable Maintainer tab" />`
   + `<Button x:Name="EnableMaintainerTabHelpButton" Classes="HelpIconButton" ToolTip.Tip="Open help page for the Maintainer tab" Click="OnEnableMaintainerTabHelpClick">`
   with the same `&#xf059;` glyph, and one line of plain text under the checkbox: "For maintainers
   of the hardware data. The tab needs a maintainer account, given by the administrator through
   an invitation." Code-behind: set `IsChecked` from `UserSettings.EnableMaintainerTab` in the
   constructor BEFORE subscribing `IsCheckedChanged` (the file's own rule), handler writes the
   setting and calls `mainWindow.ApplyMaintainerTabVisibility()`, help click opens
   `AppConfig.WikiPageUrl(AppConfig.WikiPageMaintainer)` through `ExternalTargetLauncher.TryOpen`
   and logs a warning on refusal (copy `OnEnableWorklogHelpClick`). Add `AppConfig.WikiPageMaintainer = "Maintainer-tab"`
   and create `Assets/Wiki/Maintainer-tab.md` (full text is in Step 4; a first version with the
   page's purpose and the checkbox is enough here so `WikiHelpPageNamesTests` passes). Add a row for
   the new page to `Assets/Wiki/_Sidebar.md` under "The tabs" (look at how the other tab pages are
   listed) and an entry in `ConfigurationHelpIconTests` for the new "?" button (it asserts each
   help button exists, carries the class and glyph, and shares a row with its checkbox).
5. Tests, in `tests/CRT.App.Tests/Ui/Maintainer/`: copy every `tests/CRT.Maintainer.Tests/Ui/*Tests.cs`
   except `SharedTableColourKeysTests.cs` and `MaintainerMainSettingsTests.cs`; do NOT copy
   `TestAppBuilder.cs`/`UiTest.cs` (CRT.App.Tests has the real ones - `UiTest.Run`/`RunAsync`).
   Rename `MaintainerMain*Tests` to `TabMaintainer*Tests`, `new MaintainerMain()` to
   `new TabMaintainer()`; where a test moved the window's `Content` into a plain window to get a
   layout pass (`MaintainerMainQueueTests.cs:167-172, 448-451, 511-513`, `MaintainerMainModesTests.cs:440-443`
   - they read `main.Width`/`main.Height`, which are NaN on a UserControl), put the tab itself in
   `new Window { Content = tab, Width = 1100, Height = 720 }`, `Show()` it and `Dispatcher.UIThread.RunJobs()`
   - the pattern of the PRIVATE `BuildShownTab` helpers in `WorkbooksBoardPreviewTests.cs:140-146`
   and `SchematicsThumbnailsContextMenuTests.cs:38-46` (copy the pattern into one private helper
   per test class; they are not shared today). All join `[Collection("HeadlessUi")]`. From `MaintainerMainSettingsTests` keep the "Show changes only
   is applied from settings" idea as a `TabMaintainer` test reading `UserSettings.LoadFrom` (it
   then joins the `"UserSettings"` collection instead - a class can be in ONE collection, so if it
   must also build UI, drive `UserSettings` through `LoadFrom` inside the `"HeadlessUi"` collection
   the way `WorkbooksListTests` does and read CLAUDE.md's note on that race before choosing).
   The moved UI tests now run under CRT's `HeadlessTestApp` (CRT's global styles and palette, not
   the maintainer's plain Fluent theme), so an assertion on a padding, an opacity or a colour that
   CRT's global styles change must be re-read: if CRT's style is the intended new look, update the
   expectation and say so; if the style breaks the screen (a heading no longer distinguishable
   from a row, a badge unreadable), fix the tab's own style with a more specific selector - CRT is
   the reference, but a screen that stops reading right is a regression.
   Add: `TabConstructionTests` case for `TabMaintainer`; `BusyOverlayHostsTests` - its hardcoded
   paths change to `src/CRT.App/Tabs/Maintainer/FileTreeWindow.axaml` (kept as a host) and the
   `MaintainerMain.axaml` row is replaced by a fact that `src/CRT.App/Tabs/Maintainer/TabMaintainer.axaml`
   declares NO `<ui:BusyOverlay` (a second overlay inside Main would dim only the tab and block
   nothing above it); `SharedTableColourKeysTests` gains the "every key in `Tabs/Maintainer/*.axaml`
   resolves" fact described under "Theme";
   new `MainMaintainerTabTests` built on `new CRT.Main()` - `Main` IS constructed headlessly by
   `MainWindowTests.cs` (:122, with the `UserSettings.LoadFrom`/`WorklogManager.LoadFrom`
   redirects at :39-53 - copy that setup), `MainDraftBadgeTests.cs:50` and
   `ThumbnailsDetachPerBoardTests.cs:90-103` - covering: with `EnableMaintainerTab` on, the tab
   item is visible and sits between Drafts and Configuration; selecting it collapses columns 0
   and 1 of `RootGrid`, hides `LeftPanel`, `MainSplitter` and `WorklogBar`; selecting Schematics
   again restores the width that was there before and `WorklogBar` follows `EnableWorklog`;
   `UserSettings.LeftPanelWidth` is unchanged after an enter/leave round trip; with the setting
   off the tab is hidden and, if it was selected, the selection moved off it; and
   `ShouldReturnFocusToComponentSearch()` is false while the Maintainer tab is selected - the
   existing `MainWindowTests.cs:188-197` does exactly this for the Drafts tab, so add the case
   beside it. (No pure `MaintainerLayout` helper is needed; the arithmetic is two assignments.)
   New Configuration tests: the checkbox reflects and writes `EnableMaintainerTab`
   (`UserSettings.LoadFrom` temp file, `"UserSettings"` collection - look at how the existing
   Configuration tests are collected), and the help icon: TWO new `[Fact]`s in
   `ConfigurationHelpIconTests.cs` copying :27-57 (button exists with the `HelpIconButton` class
   and the `` glyph; it shares a visual parent with `EnableMaintainerTabCheckBox`).
   `TabConstructionTests`: one new `[Fact] The_maintainer_tab_can_be_constructed` beside the
   others (:22-113).
6. Run the app (`dotnet run --project src/CRT.App` or the VS Code `watch` task; untick "Check for
   new or updated data at application launch" first if the sync is slow) and check by hand, then
   render PNGs headlessly for the handover (memory: "Render UI to PNG before handover" - a scratch
   console app in the scratchpad folder that builds `TabMaintainer` inside a window under the
   headless Skia platform; never `Show()` anything that reaches the live server): the sign-in
   panel; the queue panel with the four buttons; the Systems screen; a table open. Check
   specifically: the sidebar and worklog bar collapse on entering the tab and come back on leaving
   with the saved width intact; typing in the password box is not interrupted by the component
   search focus steal; a remembered session signs in on first entering the tab; the minute check
   does not fire while another tab is selected (log lines); closing CRT with unsaved table edits
   asks; "Enable Maintainer tab" unticked while signed in hides the tab and re-ticking brings it
   back signed in.
7. Full suite green. Summary states the counts and that CRT.Maintainer is still present and
   duplicated on purpose until Step 3.

### Step 3 - Retire CRT Maintainer

> **Carried over from Step 2 (done 2026-09-29):**
> - `BusyOverlayHostsTests` and `SharedTableColourKeysTests` (CRT.App.Tests) already point at
>   `src/CRT.App/Tabs/Maintainer/`; nothing in them names `src/CRT.Maintainer/` any more.
> - `ReviewHighlightCanvas` was copied into `src/CRT.App/Tabs/Maintainer/` as the plan said, but
>   NOTHING references it in either application - it drew moved highlights for the change summary
>   that was retired on 2026-09-26. It was left in place (dead code is proven and removed only with
>   the owner's agreement); raise it in the Step 3 summary.
> - The "Show changes only" choice reaches the tab through `TabMaintainer.UseRememberedChoices`,
>   called by `Main`'s constructor - NOT read in the tab's constructor, so a tab built by a test
>   never touches `UserSettings`.
> - The tab's waits run under Main's overlay; the moved UI tests use `MaintainerTabHost.AddOverlay`
>   (tests/CRT.App.Tests/Ui/Maintainer/) to stand in for Main.

Goal: nothing in the repository builds, tests, ships or documents a separate maintainer application
(documentation prose is Step 4; this step removes the code and the machinery).

1. Delete `src/CRT.Maintainer/` and `tests/CRT.Maintainer.Tests/` (whole folders, including their
   `bin/`/`obj/`). Remove both from `Classic-Repair-Toolbox.slnx`.
2. Remove `InternalsVisibleTo` entries naming `CRT.Maintainer` or `CRT.Maintainer.Tests` from
   `src/CRT.Data/CRT.Data.csproj` and `src/CRT.UI/CRT.UI.csproj` (grep all csproj files).
3. Delete `.github/workflows/build-and-release-maintainer.yml`. Delete
   `tests/CRT.App.Tests/MaintainerReleaseSeparationTests.cs` (its reason is gone; say so in the
   summary and in CLAUDE.md). Check `.vscode/launch.json` and `.vscode/tasks.json` for maintainer
   launch/build entries and remove them. Check `.gitignore` for maintainer-specific lines.
4. **The server's user-facing text** (the one server change in this plan; "vocabulary the user
   reads ... must be changed in all of them" - CLAUDE.md):
   - `EmailTemplates.MaintainerDownloadUrl` -> CRT's releases page
     (`https://github.com/HovKlan-DH/Classic-Repair-Toolbox/releases`; rename the constant to
     `CrtDownloadUrl`), and the invitation mail's sentence around it now says to install CRT (or
     use the one already installed), tick "Enable Maintainer tab" in the Configuration tab and
     choose "I have an invitation" there. Every "Open CRT Maintainer" (~lines 111, 164, 373, 574)
     -> "Open CRT and go to the Maintainer tab". Keep the mails' shape otherwise.
   - `ApprovePublishFlow.RemovalsNotSentMessage` (:413-416, reused by `ProductionPromotionFlow.cs:235`):
     "your copy of CRT Maintainer did not send ... Update CRT Maintainer" -> "your copy of CRT
     ... Update CRT".
   - Update `EmailTemplatesTests` (lines ~42 and ~241 assert "CRT Maintainer" in the body -
     assert the new wording) and `MaintainerInvitationFlowsTests` (~89 asserts the old releases
     URL; ~108 passes "CRT Maintainer" as the user-agent test value - change it to `"CRT 3.0.0"`
     or leave it, it is only data). Read each server test that mentions the app before deciding.
   - **Bump `CRT.Server`'s `InformationalVersion` by PATCH**: it is `3.5.0` today
     (`src/CRT.Server/CRT.Server.csproj:65`; newest `VERSION.md` row is `3.5.0 | 2026-09-28 |
     MINOR` at line ~64), so `3.5.1` - re-read the csproj first in case it moved. "A reworded mail
     body" is the PATCH example in `VERSION.md`. Add the history row: what a caller notices is
     only that the mails and one refusal message name CRT and its Maintainer tab instead of CRT
     Maintainer. That also satisfies the `server-version-bump.sh` Stop hook.
   - `appsettings.Example.json` (~40) and `ServerOptions.cs` (~50): reword the comments.
5. Grep the whole repository (excluding `CHANGELOG.md`, `Assets/NewContributeStrategy.md`,
   `src/CRT.Server/VERSION.md`'s history table, `src/CRT.Server/DEPLOYMENT.md` and `.claude/`,
   which Step 4 handles) for `CRT.Maintainer`, `Classic-Repair-Toolbox-Maintainer`, `MaintainerMain`,
   `MaintainerApp`, `maintainer-win`, `maintainer-linux`, `REVIEW_RELEASES_TOKEN`, `CRT Maintainer`,
   `maintainer app`, `maintainer application`. Every hit is either removed, reworded to "the
   Maintainer tab", or (in a comment explaining history, e.g. "the app became CRT Maintainer on
   2026-09-25", migrations 0010/0012) left alone - decide each one by reading it. `CRT.UI.csproj`'s
   and `BoardTableColors.axaml`'s headers explaining that the table is shared by "two
   applications" are reworded to "CRT's Drafts tab and its Maintainer tab" - the sharing is
   still real and the library stays. Also fix the "four test projects" comments in
   `build-and-unittest.yml` (~79) and `coverage-summary.sh` (line 18) to three.
6. `dotnet build Classic-Repair-Toolbox.slnx -c Release` and the full suite: three test projects
   now. Report the total against the Step 2 total - the difference must be exactly the deleted
   `MaintainerReleaseSeparationTests` cases plus the retired maintainer-only tests still present
   in Step 2 (the old project's own copies), nothing else.

### Step 4 - Documentation, handover notes, memory

Goal: every document that describes the system describes the merged one; the owner has a list of
the manual steps only they can do.

1. `.claude/CLAUDE.md` (the owner's instructions file - edit carefully, keep its voice):
   - "One change, every side of it": the parts are now the CRT desktop app (which includes the
     Maintainer tab), the contribution service, the shared library, and the board data. Reword the
     sentence naming the maintainer application; keep the review-API contract paragraph
     (`ReviewWireContractTests` now lives in CRT.App.Tests).
   - "Tests": three projects in the table (drop the CRT.Maintainer.Tests row; fold its coverage
     into the CRT.App.Tests row); "What is covered" gains a "Maintainer tab" row listing the moved
     Handlers classes; the headless UI table gains rows for the moved UI tests (one row per file
     in the style of the existing ones); mention `MainMaintainerLayoutTests` and the
     `"ReviewSession"` collection under "Test seams".
   - "Tabs (`Tabs/`)": add `Maintainer` to the folder list; add a "#### Maintainer (`Tabs/Maintainer/`)"
     section carrying over the substance of today's "### CRT Maintainer" section (the four screens,
     the queue rules, the table-is-the-submission-view rules, the file tree, placement, invitations)
     with the app-specific parts rewritten per the Lifecycle table, and the "Never `Show()`
     `MaintainerMain` in a test" rule replaced by the tab's equivalent. Remove the old section.
   - "Code layout conventions" BusyOverlay paragraph: drop `MaintainerMain`; `FileTreeWindow` stays.
   - "Release process": remove the paragraph about the maintainer app's own repository and
     `MaintainerReleaseSeparationTests`; add one sentence recording that CRT Maintainer was merged
     into CRT on 2026-09-29 and its release repository is retired.
   - "CRT.Server's version is YOURS to bump": "The project owner bumps CRT and CRT Maintainer" ->
     "bumps CRT".
   - The "Contribution service" section's sentences that say "the maintainer app" -> "the Maintainer
     tab" where they describe today's behaviour.
2. `Assets/NewContributeStrategy.md`: in "Decisions already made", add a reversal note under the
   existing one about the draft format: the "Review tool: Separate Avalonia desktop app" row is
   reversed on 2026-09-29 - the maintainer surface is a tab in CRT (why: a maintainer is a CRT
   user; two apps, two versions and two release repositories for a handful of people was cost with
   no benefit); the "Maintainer server access" row's "an account in the maintainer app" -> "a
   maintainer account, used from CRT's Maintainer tab". Update the architecture diagram (one
   desktop app box) and line 142's sentence. Intro lines 5, 16 and 22, line 323 and line 4793 are
   reworded; the security model's "The maintainer application contains no authority. It is a
   rendering surface..." (4549-4550) becomes "The Maintainer tab contains no authority..." - the
   property is unchanged and still important. Add
   "## Phase 8 - Maintainer tab inside CRT [DONE 2026-09-29]" after Phase 7 with: goal, what moved
   where (the file table above), the Lifecycle decisions, what was retired (project, tests,
   workflow, repository, version, settings file), the settings migration, and traps for the next
   agent (the base-address `/api` difference; the DPAPI entropy string; the attach/detach gate on
   the minute check; the closing prompt ordering; `LeftPanelWidth` must not be saved while
   collapsed). Add it to the table of contents. Phase 5's text is history - leave it, add one line
   at its top pointing to Phase 8.
3. Wiki (`Assets/Wiki/`, WIKI syntax - no paths, no extensions in links):
   - `Configuration-tab.md`: document the new "Maintainer" section and its checkbox, in the page's
     existing style, linking `[Maintainer tab](Maintainer-tab)`.
   - `Maintainer-tab.md` (created in Step 2, completed here): what the tab is for, who can use it
     (a maintainer account comes only by invitation from the administrator), how to show it
     (Configuration checkbox), signing in / staying signed in (Windows remembers the session
     encrypted to the Windows user; other systems ask each launch), the four screens in a few
     sentences each, "Sign out". Write for hobbyists per CLAUDE.md's audience rule; no job/income
     wording. Do not paste server details from DEPLOYMENT.md.
   - `_Sidebar.md`: the entry added in Step 2 - check it renders in the right group.
   - `!sync-status.md` is generated by the Stop hook - do not hand-edit it; both pages will appear
     there for the owner to paste.
4. `src/CRT.Server/DEPLOYMENT.md`: every "CRT Maintainer" / "the maintainer application" that
   describes how a maintainer does something today -> "the Maintainer tab in CRT" (~lines 111,
   1159-1180 step 13, 1267, 1300-1316, 1361-1378, 1393, 1410, 1444, 1458, 1552-1553, 1622).
   Sentences that record what was deployed WHEN stay as history. Add a short dated note in the
   section about the app rename (~1361) that the app was folded into CRT on 2026-09-29, and that
   the session file (~1558-1578: `maintainer-session.json`, DPAPI on Windows, nothing stored
   elsewhere) is unchanged and now written by CRT. `src/CRT.Server/VERSION.md`: preamble lines 6
   and 16 only (the Step 3 history row is already there). `.claude/hooks/server-version-bump.sh`
   line 5 comment. `Assets/Wiki/Contribute-data-via-CRT.md:445`: link the new page where it says
   the project owner may set you up as a maintainer. No further server bump: this step is docs
   and comments only.
5. Memory (`C:\Users\Dennis\.claude\projects\...\memory\`): update
   `naming-maintainer-and-project-owner.md` (the review app is no longer a separate app; "the
   Maintainer tab"), `render-ui-to-png-before-handover.md` (the "never Show() MaintainerMain" line
   becomes the tab rule), `crt-server-version-is-mine.md` ("Dennis handles CRT"), and add a
   `maintainer-tab-merge.md` project memory pointing at Phase 8 of the strategy document. Update
   `MEMORY.md` lines accordingly.
6. Final summary for the owner, including the manual steps only they can do:
   - Bump CRT's `InformationalVersion` when releasing (the tab ships with it) - not done by the agent.
   - Archive (or leave a final "moved into CRT" release note on) the
     `HovKlan-DH/Classic-Repair-Toolbox-Maintainer` repository; delete the `REVIEW_RELEASES_TOKEN`
     secret from this repository's settings.
   - Tell the maintainers to uninstall CRT Maintainer and, in CRT, tick Configuration >
     "Enable Maintainer tab"; their remembered session carries over on Windows (same file, same
     folder).
   - Paste the two Wiki pages listed in `!sync-status.md`.
   - Write the CHANGELOG entry.

---

## Verification (end to end, after Step 4)

1. `dotnet build Classic-Repair-Toolbox.slnx -c Release` warning-free; `dotnet test Classic-Repair-Toolbox.slnx`
   green across the THREE test projects (quote the counts).
2. Run CRT with `--simulate-update` off and data sync unticked. Configuration > tick "Enable
   Maintainer tab" -> the tab appears after Drafts; select it -> sidebar and worklog bar collapse;
   the sign-in panel shows the reworded explanation; sign in with a maintainer account -> the queue
   panel with Systems / Contributor Submissions / Beta > Prod (/ Admin for an administrator);
   choose a submission -> its table opens in the panel with "Show changes only" as last saved;
   switch to Schematics -> layout restored, sidebar width unchanged; come back -> table still open;
   close CRT with an unsaved edit -> the prompt, on the Maintainer tab; relaunch -> signed in
   without a password (Windows), tab still enabled, "Show changes only" remembered in
   `Classic-Repair-Toolbox.settings.json`, `CRT-Maintainer-Settings.json` gone.
3. `GET /api/health` reports the PATCH-bumped server version; the server log shows the review
   calls arriving with `User-Agent: CRT <version>`; an invitation mail (owner sends one to a test
   address after deploying) names CRT's releases page and the Maintainer tab. Deploying the server
   is the owner's step (DEPLOYMENT.md); the agent builds it (`dotnet publish` per DEPLOYMENT.md)
   and says so.
4. `git grep -i "CRT.Maintainer\|Classic-Repair-Toolbox-Maintainer\|MaintainerMain"` returns only
   history (CHANGELOG, VERSION.md's table, NewContributeStrategy.md's Phase 5, dated notes).
5. Headless PNG renders of the four states listed in Step 2.6, attached to the handover.
