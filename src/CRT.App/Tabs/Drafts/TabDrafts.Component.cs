using Handlers.DataHandling;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace CRT
{
    // ###########################################################################################
    // THE COMPONENT PANE on the right of the Contribute tab (CRT 4.0 step 2, 2026-10-25 - the plan:
    // "A component opens in a pane, never in a maximized window").
    //
    // Owns: which component of which draft is open (thisComponentEntry), opening and closing the pane -
    // on a component, or on its form for a new one ("Add component", slice 3c) - and what follows its
    // changes - the left column reads the draft as the pane left it, the draft's row follows each save,
    // and the board on screen catches up once the pane closes or the tab is left (CatchUpBoardOnScreen,
    // as for the table: reading the whole board after every typed field would make typing lag). And
    // the pane's two requests of Main: "Show on board" and "Preview popup".
    //
    // The right side shows ONE of the table, the pane and the start panel (ApplyTableMode). A
    // component opened closes the table, and the table opened closes the pane - each having saved
    // what it holds.
    //
    // Part of the TabDrafts partial class - see TabDrafts.axaml.cs for the file map.
    // ###########################################################################################
    public partial class TabDrafts
    {
        // The draft the pane is open on - null while it is closed.
        private HardwareBoardEntry? thisComponentEntry;

        // ###########################################################################################
        // What the pane changed, per label - "Changed" in the left column follows the pane at once,
        // while the board on screen (where the marks come from) catches up later. Forgotten whenever
        // Main shows the board again: its marks are then up to date.
        // ###########################################################################################
        private readonly Dictionary<string, bool> thisChangedHere = new(StringComparer.OrdinalIgnoreCase);

        internal bool IsComponentOpen => this.thisComponentEntry is not null;

        private bool IsComponentOpenFor(HardwareBoardEntry entry) =>
            this.thisComponentEntry is not null &&
            string.Equals(this.thisComponentEntry.ExcelDataFile, entry.ExcelDataFile, StringComparison.OrdinalIgnoreCase);

        internal ComponentPane ComponentPaneForTests => this.ComponentPane;

        private void WireComponentPane()
        {
            this.ComponentPane.CloseRequested += (_, _) => this.CloseComponent();
            this.ComponentPane.ContentChanged += (_, _) => this.OnComponentPaneChanged();
            this.ComponentPane.Saved += (_, excelDataFile) => this.OnComponentPaneSaved(excelDataFile);
            this.ComponentPane.ShowOnBoardRequested += (_, label) => this.RunGuarded(() => this.ShowComponentOnBoardAsync(label), "Show on board");
            this.ComponentPane.PreviewRequested += (_, label) => this.PreviewComponent(label);
            this.ComponentPane.ManagePicturesRequested += (_, _) => this.ManagePictures();
            this.ComponentPane.PictureToolRequested += (_, asked) => this.RunGuarded(() => this.StartPictureToolAsync(asked.Tool, asked.Picture), "Picture tool");
            this.ComponentPane.FindOnOtherBoards = this.FindOnOtherBoardsAsync;
            this.ComponentPane.ApplyToOtherBoards = this.ApplyToOtherBoardsAsync;
        }

        // ###########################################################################################
        // "Copy from another board" (slice 5): the components with `partNumber` on the OTHER published
        // boards of this machine - read off the UI thread under CRT's "please wait" (fourteen board files
        // take a moment the first time; BoardDataReader keeps them after). A test hands each board's
        // published data in through PublishedBoardOverrideForTests, as for the pane.
        // ###########################################################################################
        private async Task<IReadOnlyList<ComponentCopySource>> FindOnOtherBoardsAsync(string partNumber, string region)
        {
            HardwareBoardEntry? shown = this.thisComponentEntry ?? this.thisCurrentEntry;

            List<HardwareBoardEntry> boards = (this.HardwareBoardsOverrideForTests ?? DataManager.HardwareBoards)
                .Where(entry => !entry.IsDraftOnly && (shown is null || !TabDrafts.SameBoard(entry, shown)))
                .ToList();

            Func<HardwareBoardEntry, BoardData?>? published = this.PublishedBoardOverrideForTests;

            return await BusyOverlay.RunLocalAsync(this, CrtWaitWording.LookingAtOtherBoards, async () =>
            {
                var read = new List<(string BoardName, string BoardIdentity, BoardData Board)>();

                foreach (HardwareBoardEntry entry in boards)
                {
                    if (published is not null)
                    {
                        if (published(entry) is { } given)
                        {
                            read.Add((entry.ToString(), entry.ExcelDataFile, given));
                        }

                        continue;
                    }

                    string path = DraftBoardSource.PublishedPathOf(DataManager.DataRoot, entry.ExcelDataFile);

                    if (path.Length > 0 && System.IO.File.Exists(path) && await BoardDataReader.LoadAsync(path, path) is { } board)
                    {
                        read.Add((entry.ToString(), entry.ExcelDataFile, board));
                    }
                }

                return (IReadOnlyList<ComponentCopySource>)ComponentCopy.Sources(read, partNumber, region);
            });
        }

        // ###########################################################################################
        // "Also on other boards" (slice 5): `parts` of `here` - the component as this draft has it - put
        // onto each target's component, in THAT board's own draft: made first when it has none (seeded
        // from the published file, as every first draft is), the change ONE step there named after this
        // board, saved at once - all under one "please wait". What came of it, in words. Nothing while
        // the Contribute tab is covered by "CRT has to be updated" (no draft is made then).
        // ###########################################################################################
        private async Task<string> ApplyToOtherBoardsAsync(ComponentCopySource here, IReadOnlyList<ComponentCopySource> targets, IReadOnlyList<ComponentCopyPart> parts)
        {
            if (this.IsUpdateRequiredShown)
            {
                return Handlers.Online.AppUpdateRequiredWording.NoDraftWhileDraftsTabCovered;
            }

            string fromName = (this.thisComponentEntry ?? this.thisCurrentEntry)?.ToString() ?? string.Empty;
            ComponentCopySource source = here with { BoardName = fromName };

            var changed = new List<string>();
            var notChanged = new List<(string Board, string Why)>();

            string draftsRoot = DraftManager.DraftsRoot;
            string dataRoot = DataManager.DataRoot;

            await BusyOverlay.RunLocalAsync(this, CrtWaitWording.ChangingOtherBoards, async () =>
            {
                foreach (ComponentCopySource target in targets.Where(target => target.BoardIdentity.Length > 0))
                {
                    if (!DraftBoardSource.HasDraft(draftsRoot, target.BoardIdentity) &&
                        !(await Task.Run(() => DraftSeeder.SeedFromPublishedFile(draftsRoot, dataRoot, target.BoardIdentity))).Created)
                    {
                        notChanged.Add((target.BoardName, ComponentPaneWording.AlsoNoDraft));
                        continue;
                    }

                    string label = target.Row.BoardLabel;
                    string region = target.Row.Region ?? string.Empty;
                    bool changesSomething = false;

                    // ###########################################################################################
                    // *** AS EVERY EDITOR DOES - DraftEdits (code review, 2026-10-26). *** The draft read first
                    // when its file changed, the change worked out off the UI thread, and taken back OUT when
                    // the save cannot be made: applied straight to the session, a failed save left the step in,
                    // and the timer that saves drafts no editor holds could save it after CRT said it had not.
                    // ###########################################################################################
                    DraftWorkbookEditOutcome outcome = await DraftEdits.ApplyAndSaveAsync(
                        target.BoardIdentity,
                        DraftStepWording.ComponentCopied(label, fromName),
                        (BoardData board) =>
                        {
                            BoardData copied = ComponentCopy.Apply(board, label, region, source, parts);
                            changesSomething |= !ReferenceEquals(copied, board);
                            return copied;
                        });

                    if (outcome == DraftWorkbookEditOutcome.NoDraft)
                        notChanged.Add((target.BoardName, ComponentPaneWording.AlsoNoDraft));
                    else if (outcome != DraftWorkbookEditOutcome.Saved)
                        notChanged.Add((target.BoardName, ComponentPaneWording.AlsoNotSaved));
                    else if (!changesSomething)
                        notChanged.Add((target.BoardName, ComponentPaneWording.AlsoAlreadySo));
                    else
                        changed.Add(target.BoardName);
                }

                return true;
            });

            // "My drafts" lists the drafts just made, and the tab's title follows.
            if (this.thisMainWindow is not null)
            {
                this.thisMainWindow.ApplyContributeTabState();
            }
            else
            {
                this.RefreshDrafts();
            }

            return ComponentPaneWording.AlsoDone(changed, notChanged);
        }

        // For tests, in place of Main: showing a component on the board (whether the board on screen
        // has to be read again first), and opening its popup on the draft's board.
        internal Action<string, bool>? ShowOnBoardOverrideForTests { get; set; }

        internal Action<BoardData, string, string>? PreviewOverrideForTests { get; set; }

        // ###########################################################################################
        // The oscilloscope, for the pane's "Capture from my oscilloscope" (slice 3b): the capture, from
        // Main's Oscilloscope tab (TabOscilloscope.CaptureOscilloscopePngAsync), and whether the tab has
        // a session now - Main says so as it changes (UpdateComponentInfoWindowsOscilloscopeSessionState).
        // ###########################################################################################
        internal void UseScopeCapture(Func<System.Threading.CancellationToken, Task<byte[]?>> capture, bool connected)
        {
            this.ComponentPane.CaptureScopeImage = capture;
            this.ComponentPane.ShowScopeConnected(connected);
        }

        internal void ShowScopeConnected(bool connected) => this.ComponentPane.ShowScopeConnected(connected);

        // ###########################################################################################
        // Opens `label` of the board shown in the pane - its draft made first when it has none
        // (OnCurrentDraftAsync), the table closed first when it is open (what it holds saved).
        // ###########################################################################################
        internal Task OpenComponentAsync(string label) =>
            this.OnCurrentDraftAsync(entry => this.OpenPaneAsync(
                entry,
                (session, published, publishedBoardFile) => this.ComponentPane.Show(session, label, this.thisCurrentRegion, published, publishedBoardFile)));

        // ###########################################################################################
        // "Add component" (slice 3c): the pane's form for a new one - named there and added to the draft
        // as one step - on the board shown, its draft made first when it has none.
        // ###########################################################################################
        internal Task AddComponentAsync() =>
            this.OnCurrentDraftAsync(entry => this.OpenPaneAsync(
                entry,
                (session, published, publishedBoardFile) => this.ComponentPane.ShowNew(session, this.thisCurrentRegion, published, publishedBoardFile)));

        // ###########################################################################################
        // One of the board's own lists in the pane (slice 4a) - its files and links, credits or
        // important signals - on the board shown, its draft made first when it has none. For the
        // signals, the board's KiCad nets are read off the UI thread and handed over once read.
        // ###########################################################################################
        internal Task OpenBoardPartAsync(PaneBoardPart part) =>
            this.OnCurrentDraftAsync(entry => this.OpenPaneAsync(
                entry,
                (session, published, publishedBoardFile) =>
                {
                    this.ComponentPane.ShowBoardPart(session, part, published, publishedBoardFile, this.thisCurrentRegion);

                    if (part == PaneBoardPart.ImportantSignals)
                    {
                        this.RunGuarded(() => this.LoadNetChoicesAsync(entry), "Reading the KiCad nets");
                    }
                }));

        // For tests, in place of reading a board's KiCad nets.
        internal Func<HardwareBoardEntry, Task<KiCadNetChoices>>? NetChoicesOverrideForTests { get; set; }

        private async Task LoadNetChoicesAsync(HardwareBoardEntry entry)
        {
            KiCadNetChoices choices = this.NetChoicesOverrideForTests is { } read
                ? await read(entry)
                : await TabDrafts.ReadNetChoicesAsync(
                    DataManager.DataRoot,
                    DraftFolderLayout.GetBoardFolder(DraftManager.DraftsRoot, entry.ExcelDataFile),
                    entry.ExcelDataFile);

            // Only to the part they were read for - the pane may show something else by now.
            if (this.thisComponentEntry is { } open && TabDrafts.SameBoard(open, entry) &&
                this.ComponentPane.BoardPart == PaneBoardPart.ImportantSignals)
            {
                this.ComponentPane.UseNetChoices(choices);
            }
        }

        // ###########################################################################################
        // The board's KiCad nets, as the draft has its KiCad data - the files the Schematics tab draws
        // from (Main.KiCadRawPathsFor, copied on the pool), every PCB file's nets by the panel's rule.
        // None, with a line in the log, when the data cannot be read.
        // ###########################################################################################
        private static async Task<KiCadNetChoices> ReadNetChoicesAsync(string dataRoot, string draftFolder, string excelDataFile)
        {
            try
            {
                List<string> files = await Task.Run(() => Main.KiCadRawPathsFor(dataRoot, draftFolder, excelDataFile));

                if (files.Count == 0)
                {
                    return KiCadNetChoices.None;
                }

                var bundle = await KiCadProjectLoader.LoadRawAsync(files);

                if (bundle?.Root is not { } root)
                {
                    return KiCadNetChoices.None;
                }

                return KiCadNetNames.Choices(root.Pcb.SelectMany(pcb => pcb.Nets.List).Select(net => (net.Name, net.NormalizedName, net.IsNoNet)));
            }
            catch (Exception ex)
            {
                Logger.Warning($"Could not read the KiCad nets of [{excelDataFile}] for its important signals - [{ex.Message}]");
                return KiCadNetChoices.None;
            }
        }

        // ###########################################################################################
        // The pane opened on `entry`'s draft - `show` hands it the session, the board as published and
        // that board's file - the table closed first when it is open (what it holds saved).
        // ###########################################################################################
        private async Task OpenPaneAsync(HardwareBoardEntry entry, Action<DraftSession, BoardData?, string> show)
        {
            if (this.IsTableOpen && !await this.CloseTableAsync())
            {
                return;
            }

            DraftStatus? status = DraftStatusReader.Resolve(DataManager.DataRoot, DraftManager.DraftsRoot, entry.ExcelDataFile, entry.IsPublished);

            BoardData? published = this.PublishedBoardOverrideForTests is not null
                ? this.PublishedBoardOverrideForTests(entry)
                : await BusyOverlay.RunLocalAsync(this, CrtWaitWording.OpeningComponent, () => TabDrafts.LoadPublishedBoardAsync(entry, status));

            // The one session of this draft, as the table opens it - something may have written the
            // file meanwhile: read again first, when nothing of the session's waits to be saved.
            DraftSession? session = DraftManager.Sessions.Get(DraftManager.DraftsRoot, entry.ExcelDataFile);

            if (session is null)
            {
                Logger.Warning($"Could not open the component pane on the draft of [{entry.ExcelDataFile}] - its board file could not be read");
                return;
            }

            session.ReadFileIfChanged(DraftSaveWording.ReadAgainStep);

            // Its list of files too, so a picture replaced under its own name reads as a change.
            string publishedBoardFile = published is null ? string.Empty : DraftBoardSource.PublishedPathOf(DataManager.DataRoot, entry.ExcelDataFile);

            show(session, published, publishedBoardFile);
            this.thisComponentEntry = entry;

            // The draft's header follows the session now open.
            this.RefreshDrafts();
        }

        // ###########################################################################################
        // "Show on board": what the pane holds saved first, then the Schematics tab with only the
        // component selected - the board on screen read again first when the pane's saves left it behind
        // (it catches up only once the pane closes or the tab is left, and a reload drops the selection).
        // The board on screen is handed that reload here, so leaving the tab reads nothing again.
        // ###########################################################################################
        internal async Task ShowComponentOnBoardAsync(string label)
        {
            if (!this.IsComponentOpen)
            {
                return;
            }

            await this.SaveComponentNowAsync();

            bool readBoardFirst = this.thisBoardOnScreenBehind is not null;
            this.thisBoardOnScreenBehind = null;

            if (this.ShowOnBoardOverrideForTests is { } show)
            {
                show(label, readBoardFirst);
                return;
            }

            this.thisMainWindow?.ShowComponentOnBoard(label, readBoardFirst);
        }

        // ###########################################################################################
        // "Mark components" and "Align KiCad" on a picture (slice 6 - contributor mode went): what the
        // pane holds saved first, then the Schematics tab in that tool - on the board read again first
        // when the pane's saves left it behind, as "Show on board" does, and handed that reload.
        // ###########################################################################################
        internal async Task StartPictureToolAsync(PictureTool tool, string picture)
        {
            if (!this.IsComponentOpen)
            {
                return;
            }

            await this.SaveComponentNowAsync();

            bool readBoardFirst = this.thisBoardOnScreenBehind is not null;
            this.thisBoardOnScreenBehind = null;

            if (this.PictureToolOverrideForTests is { } start)
            {
                start(tool, picture, readBoardFirst);
                return;
            }

            this.thisMainWindow?.StartPictureTool(tool, picture, readBoardFirst);
        }

        // Stands in for Main's start of a picture tool - a test's tab has no Main.
        internal Action<PictureTool, string, bool>? PictureToolOverrideForTests { get; set; }

        // ###########################################################################################
        // "Preview popup": the component's popup on the draft's board as it is in memory - what will be
        // published - with the draft's folder for its pictures, whatever "view boards as officially
        // published" says (that is about the board on screen, not about a preview of the draft).
        // ###########################################################################################
        private void PreviewComponent(string label)
        {
            if (this.thisComponentEntry is not { } entry || this.ComponentPane.Session is not { } session)
            {
                return;
            }

            BoardData board = session.Content.Board;

            if (this.PreviewOverrideForTests is { } preview)
            {
                preview(board, label, entry.ExcelDataFile);
                return;
            }

            this.thisMainWindow?.OpenComponentPreview(
                board,
                label,
                DraftFolderLayout.GetBoardFolder(DraftManager.DraftsRoot, entry.ExcelDataFile),
                entry.ExcelDataFile);
        }

        // ###########################################################################################
        // Closes the pane - what it holds goes into the draft, saved as it closes. The draft's row and
        // the board on screen follow that save (OnComponentPaneSaved): read now, from the file the
        // save is replacing, they kept it from being replaced (found by a test that failed in some
        // orders only - "WriteFailed").
        // ###########################################################################################
        internal void CloseComponent()
        {
            if (this.thisComponentEntry is not { } entry)
            {
                return;
            }

            this.ComponentPane.Close();
            this.thisComponentEntry = null;
            this.ApplyTableMode();
            this.ShowWorkspace();

            if (this.ComponentPane.LastSave.IsCompleted)
            {
                this.CatchUpBoardOnScreen();
            }

            // A draft made only to open the pane goes again, nothing in it (TabDrafts.Workspace.cs).
            this.LetGoOfDraftMadeOnlyToLook(entry);
        }

        // The pane let go of a draft that is going - nothing handed over, nothing saved.
        private void ForgetComponent()
        {
            this.ComponentPane.Forget();
            this.thisComponentEntry = null;
        }

        // ###########################################################################################
        // Everything the pane holds saved now - under CRT's "please wait", since the contributor
        // pressed for it. True when nothing is left unsaved (or the pane is closed).
        // ###########################################################################################
        internal async Task<bool> SaveComponentNowAsync()
        {
            if (!this.IsComponentOpen)
            {
                return true;
            }

            this.ComponentPane.CommitHeld();

            if (!this.ComponentPane.HasUnsavedChanges)
            {
                return true;
            }

            return await BusyOverlay.RunLocalAsync(this, CrtWaitWording.SavingDraft, this.ComponentPane.SaveNowAsync);
        }

        // The pane's change: the left column reads the draft as it is now.
        private void OnComponentPaneChanged()
        {
            if (this.ComponentPane.Label.Length > 0)
            {
                this.thisChangedHere[this.ComponentPane.Label] = this.ComponentPane.IsComponentChanged;
            }

            this.ShowBoardParts();
            this.ShowComponents();
        }

        // ###########################################################################################
        // The pane's draft saved: its row follows at once; the board on screen once the pane closes or
        // the tab is left - at once when the tab is not on screen (a save the pane started as it went).
        // ###########################################################################################
        private void OnComponentPaneSaved(string excelDataFile)
        {
            this.thisBoardOnScreenBehind = excelDataFile;

            if (this.thisMainWindow is not null)
            {
                this.thisMainWindow.ApplyContributeTabState();
            }
            else
            {
                this.RefreshDrafts();
            }

            if (this.thisActivationWindow is null || !this.IsComponentOpen)
            {
                this.CatchUpBoardOnScreen();
            }
        }

        // The board the left column shows: the draft as the pane or the table left it, when one is open
        // on the board shown - else the board Main showed.
        private BoardData? BoardForList()
        {
            if (this.thisCurrentEntry is { } entry &&
                (this.IsComponentOpenFor(entry) || this.IsTableOpenFor(entry)) &&
                DraftManager.Sessions.Find(DraftManager.DraftsRoot, entry.ExcelDataFile) is { } session)
            {
                return session.Content.Board;
            }

            return this.thisCurrentBoard;
        }

        // Main's marks, with what the pane changed since the board was last shown.
        private IReadOnlySet<string>? DraftedLabelsForList()
        {
            if (this.thisChangedHere.Count == 0)
            {
                return this.thisDraftedLabels;
            }

            var labels = new HashSet<string>(this.thisDraftedLabels ?? new HashSet<string>(), StringComparer.OrdinalIgnoreCase);

            foreach ((string label, bool changed) in this.thisChangedHere)
            {
                if (changed)
                {
                    labels.Add(label);
                }
                else
                {
                    labels.Remove(label);
                }
            }

            return labels;
        }
    }
}
