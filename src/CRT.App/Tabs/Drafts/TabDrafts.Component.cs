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
            this.ComponentPane.MarkAutomaticallyRequested += (_, picture) => this.RunGuarded(() => this.MarkAutomaticallyAsync(picture), "Mark all automatically");
            this.ComponentPane.FindOnOtherBoards = this.FindOnOtherBoardsAsync;
            this.ComponentPane.FindAlsoTargets = this.FindAlsoTargetsAsync;
            this.ComponentPane.ApplyToOtherBoards = this.ApplyToOtherBoardsAsync;
            this.ComponentPane.ComponentShown += (_, label) => this.RunGuarded(() => this.LoadKiCadPinsAsync(label), "Reading the KiCad pins");
            this.ComponentPane.ShowNetRequested += (_, net) => this.RunGuarded(() => this.ShowNetOnBoardAsync(net), "Show on board");
        }

        // For tests, in place of reading a component's pins from its board's KiCad data: the board
        // (its identity) and the label.
        internal Func<string, string, Task<IReadOnlyList<KiCadPin>>>? KiCadPinsOverrideForTests { get; set; }

        // ###########################################################################################
        // The pins of `label` from its board's KiCad data - the files the Schematics tab draws from
        // (Main.KiCadRawPathsFor, copied on the pool), parsed once and kept by KiCadProjectLoader -
        // handed to the pane if it still shows that component. None, with a line in the log, when
        // the data cannot be read.
        // ###########################################################################################
        private async Task LoadKiCadPinsAsync(string label)
        {
            if (this.ComponentPane.Session is not { } session)
            {
                return;
            }

            // Which ask this read answers - one started later wins, whichever ends last (code review,
            // 2026-10-26: a read from before a KiCad import put the old pins back).
            int ask = this.ComponentPane.KiCadPinsAsk;
            string identity = session.ExcelDataFile;

            IReadOnlyList<KiCadPin> pins = this.KiCadPinsOverrideForTests is { } read
                ? await read(identity, label)
                : await TabDrafts.ReadBoardKiCadAsync(DataManager.DataRoot, DraftFolderLayout.GetBoardFolder(session.DraftsRoot, identity), identity, "a component's pins") is { } root
                    ? await Task.Run(() => KiCadPinTable.PinsOf(root, label))
                    : [];

            this.ComponentPane.UseKiCadPins(session, label, pins, ask);
        }

        // ###########################################################################################
        // The board's KiCad project as the Schematics tab draws it - the files Main.KiCadRawPathsFor
        // copies (on the pool), parsed once and kept by KiCadProjectLoader - for the important signals'
        // nets and a component's pins alike (code review, 2026-10-26: two copies of this read). Null
        // without KiCad data, or when it cannot be read - said in the log, for `what`.
        // ###########################################################################################
        private static async Task<KiCadProjectRoot?> ReadBoardKiCadAsync(string dataRoot, string draftFolder, string excelDataFile, string what)
        {
            try
            {
                List<string> files = await Task.Run(() => Main.KiCadRawPathsFor(dataRoot, draftFolder, excelDataFile));

                return files.Count == 0 ? null : (await KiCadProjectLoader.LoadRawAsync(files))?.Root;
            }
            catch (Exception ex)
            {
                Logger.Warning($"Could not read the KiCad data of [{excelDataFile}] for {what} - [{ex.Message}]");
                return null;
            }
        }

        // Stands in for seeding a board's first draft from its published file - a test has no data
        // folder. True when the draft was made.
        internal Func<string, bool>? SeedDraftOverrideForTests { get; set; }

        // Where a headless test's published board FILE lies - whose list of files says which bytes a
        // board has. null is the shipped path: the board's file in the data folder.
        internal Func<HardwareBoardEntry, string>? PublishedBoardFileOverrideForTests { get; set; }

        private string PublishedBoardFileOf(HardwareBoardEntry entry) =>
            this.PublishedBoardFileOverrideForTests?.Invoke(entry) ?? DraftBoardSource.PublishedPathOf(DataManager.DataRoot, entry.ExcelDataFile);

        // One other published board as "Copy from another board" and "Also on other boards" read it.
        private sealed record OtherBoard(HardwareBoardEntry Entry, BoardData Board, IReadOnlyDictionary<string, BoardDocument.StoredFileRow>? Files);

        // ###########################################################################################
        // The OTHER published boards of this machine - never the board the pane is on - read IN
        // PARALLEL off the UI thread: each one's data and, with `withFiles`, its list of files, both
        // from ONE read of its board file (code review, 2026-10-26: read one after another, and twice
        // each). Without files, BoardDataReader's cache answers. A test hands each board's published
        // data in through PublishedBoardOverrideForTests, as for the pane.
        // ###########################################################################################
        private async Task<List<OtherBoard>> ReadOtherPublishedBoardsAsync(bool withFiles)
        {
            HardwareBoardEntry? shown = this.thisComponentEntry ?? this.thisCurrentEntry;

            List<HardwareBoardEntry> boards = (this.HardwareBoardsOverrideForTests ?? DataManager.HardwareBoards)
                .Where(entry => !entry.IsDraftOnly && (shown is null || !TabDrafts.SameBoard(entry, shown)))
                .ToList();

            Func<HardwareBoardEntry, BoardData?>? published = this.PublishedBoardOverrideForTests;

            OtherBoard?[] read = await Task.WhenAll(boards.Select(entry => TabDrafts.ReadOtherBoardAsync(entry, this.PublishedBoardFileOf(entry), published, withFiles)));

            return read.OfType<OtherBoard>().ToList();
        }

        private static async Task<OtherBoard?> ReadOtherBoardAsync(HardwareBoardEntry entry, string path, Func<HardwareBoardEntry, BoardData?>? published, bool withFiles)
        {
            if (published is not null)
            {
                return published(entry) is { } given
                    ? new OtherBoard(entry, given, withFiles ? await Task.Run(() => DraftFiles.ListedIn(path)) : null)
                    : null;
            }

            if (path.Length == 0 || !System.IO.File.Exists(path))
            {
                return null;
            }

            if (!withFiles)
            {
                return await BoardDataReader.LoadAsync(path, path) is { } board ? new OtherBoard(entry, board, null) : null;
            }

            return await Task.Run(() =>
            {
                if (BoardFiles.ReadDocument(path) is not { } document)
                {
                    return null;
                }

                var files = new Dictionary<string, BoardDocument.StoredFileRow>(StringComparer.OrdinalIgnoreCase);

                foreach (BoardDocument.StoredFileRow file in document.Files)
                {
                    files.TryAdd(file.Name, file);
                }

                return new OtherBoard(entry, BoardDocumentMapping.ToContent(document).Board, files);
            });
        }

        // "Copy from another board" (slice 5): the components with `partNumber` on the OTHER published
        // boards, under CRT's "please wait".
        private async Task<IReadOnlyList<ComponentCopySource>> FindOnOtherBoardsAsync(string partNumber, string region) =>
            await BusyOverlay.RunLocalAsync(this, CrtWaitWording.LookingAtOtherBoards, async () =>
            {
                List<OtherBoard> others = await this.ReadOtherPublishedBoardsAsync(withFiles: false);

                return (IReadOnlyList<ComponentCopySource>)ComponentCopy.Sources(
                    others.Select(other => (other.Entry.ToString(), other.Entry.ExcelDataFile, other.Board)),
                    partNumber,
                    region);
            });

        // ###########################################################################################
        // "Also on other boards": the components with `partNumber` ("" for none to look for) on the
        // OTHER published boards, and the boards listing the bytes each of `files` replaced
        // (ChangedFiles.Elsewhere) - ONE read of the other boards under ONE "please wait" (code review,
        // 2026-10-26: two searches read every board twice, each under its own).
        // ###########################################################################################
        private async Task<AlsoTargets> FindAlsoTargetsAsync(string partNumber, string region, IReadOnlyList<ChangedFile> files)
        {
            bool part = partNumber.Trim().Length > 0;
            bool anyFile = files.Count > 0;

            if (!part && !anyFile)
            {
                return AlsoTargets.None;
            }

            string wait = part && anyFile ? CrtWaitWording.LookingForPartAndFilesOnOtherBoards
                : part ? CrtWaitWording.LookingAtOtherBoards
                : CrtWaitWording.LookingForFilesOnOtherBoards;

            return await BusyOverlay.RunLocalAsync(this, wait, async () =>
            {
                List<OtherBoard> others = await this.ReadOtherPublishedBoardsAsync(withFiles: anyFile);

                return new AlsoTargets(
                    part ? ComponentCopy.Sources(others.Select(other => (other.Entry.ToString(), other.Entry.ExcelDataFile, other.Board)), partNumber, region) : [],
                    anyFile ? ChangedFiles.Elsewhere(files, others.Select(other => (other.Entry.ToString(), other.Entry.ExcelDataFile, other.Board, other.Files))) : []);
            });
        }

        // ###########################################################################################
        // "Also on other boards" (slice 5, and its files): `parts` of `here` - the component as this
        // draft has it - put onto each of `targets`' components, and each of `files` - a new file in
        // place of the one it replaced - put onto its board, all in THAT board's own draft: made first
        // when it has none (seeded from the published file, as every first draft is), saved at once -
        // all under one "please wait". What came of it, in words (AlsoOnOtherBoardsOutcome). Nothing
        // while the Contribute tab is covered by "CRT has to be updated" (no draft is made then).
        // ###########################################################################################
        private async Task<string> ApplyToOtherBoardsAsync(
            ComponentCopySource? here,
            IReadOnlyList<ComponentCopySource> targets,
            IReadOnlyList<ComponentCopyPart> parts,
            IReadOnlyList<ChangedFileElsewhere> files)
        {
            if (this.IsUpdateRequiredShown)
            {
                return Handlers.Online.AppUpdateRequiredWording.NoDraftWhileDraftsTabCovered;
            }

            HardwareBoardEntry? from = this.thisComponentEntry ?? this.thisCurrentEntry;
            string fromName = from?.ToString() ?? string.Empty;
            string fromIdentity = from?.ExcelDataFile ?? string.Empty;

            var outcome = new AlsoOnOtherBoardsOutcome();

            string draftsRoot = DraftManager.DraftsRoot;
            string dataRoot = DataManager.DataRoot;

            // Every board once, in the order offered: the part's boards first, then those with a file.
            var boards = targets.Select(target => (Identity: target.BoardIdentity, Name: target.BoardName))
                .Concat(files.Select(file => (Identity: file.BoardIdentity, Name: file.BoardName)))
                .Where(board => board.Identity.Length > 0)
                .DistinctBy(board => board.Identity, StringComparer.OrdinalIgnoreCase)
                .ToList();

            await BusyOverlay.RunLocalAsync(this, CrtWaitWording.ChangingOtherBoards, async () =>
            {
                foreach ((string identity, string boardName) in boards)
                {
                    bool made = false;

                    if (!DraftBoardSource.HasDraft(draftsRoot, identity))
                    {
                        made = this.SeedDraftOverrideForTests?.Invoke(identity) ??
                               (await Task.Run(() => DraftSeeder.SeedFromPublishedFile(draftsRoot, dataRoot, identity))).Created;

                        if (!made)
                        {
                            outcome.NoDraft(boardName);
                            continue;
                        }
                    }

                    bool changedThere = false;
                    var partsNotCopied = new List<(string Label, string Why)>();

                    // EVERY ticked component of the board, each its own step (code review, 2026-10-26:
                    // the board's first only, and its other 4164s were left as they were).
                    if (here is not null && parts.Count > 0)
                    {
                        foreach (ComponentCopySource target in targets.Where(target => string.Equals(target.BoardIdentity, identity, StringComparison.OrdinalIgnoreCase)))
                        {
                            if (await TabDrafts.CopyPartsToAsync(target, here with { BoardName = fromName }, parts, fromName) is { } why)
                                partsNotCopied.Add((target.Row.BoardLabel, why));
                            else
                                changedThere = true;
                        }
                    }

                    List<ChangedFileElsewhere> boardFiles = files
                        .Where(file => string.Equals(file.BoardIdentity, identity, StringComparison.OrdinalIgnoreCase))
                        .ToList();

                    if (boardFiles.Count > 0)
                    {
                        changedThere |= await TabDrafts.PutFilesIntoAsync(identity, boardName, boardFiles, fromIdentity, fromName, outcome);
                    }

                    outcome.Board(boardName, changedThere, partsNotCopied);

                    // A draft made here for a board that then took nothing goes again (code review,
                    // 2026-10-26: each was left as an empty draft in "My drafts", to discard by hand).
                    if (made && !changedThere)
                    {
                        this.LetGoOfDraftMadeForNothing(identity);
                    }
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

            return outcome.Said;
        }

        // A draft "Also on other boards" made for a board that took nothing - deleted again.
        private void LetGoOfDraftMadeForNothing(string identity)
        {
            if (!(this.DiscardOverrideForTests?.Invoke(identity) ?? DraftManager.DiscardDraft(identity)))
            {
                Logger.Warning($"The empty draft made for [{identity}] could not be deleted again");
            }
        }

        // ###########################################################################################
        // `parts` of `source` onto the component of `target`, in its board's draft, as ONE step named
        // after this board - null when that changed it, else why not.
        //
        // *** AS EVERY EDITOR DOES - DraftEdits (code review, 2026-10-26). *** The draft read first when
        // its file changed, the change worked out off the UI thread, and taken back OUT when the save
        // cannot be made: applied straight to the session, a failed save left the step in, and the
        // timer that saves drafts no editor holds could save it after CRT said it had not.
        // ###########################################################################################
        private static async Task<string?> CopyPartsToAsync(ComponentCopySource target, ComponentCopySource source, IReadOnlyList<ComponentCopyPart> parts, string fromName)
        {
            string label = target.Row.BoardLabel;
            string region = target.Row.Region ?? string.Empty;
            bool changesSomething = false;

            DraftWorkbookEditOutcome outcome = await DraftEdits.ApplyAndSaveAsync(
                target.BoardIdentity,
                DraftStepWording.ComponentCopied(label, fromName),
                (BoardData board) =>
                {
                    BoardData copied = ComponentCopy.Apply(board, label, region, source, parts);
                    changesSomething |= !ReferenceEquals(copied, board);
                    return copied;
                });

            return outcome == DraftWorkbookEditOutcome.NoDraft ? ComponentPaneWording.AlsoNoDraft
                : outcome != DraftWorkbookEditOutcome.Saved ? ComponentPaneWording.AlsoNotSaved
                : !changesSomething ? ComponentPaneWording.AlsoAlreadySo
                : null;
        }

        // ###########################################################################################
        // Each of `files` - this draft's new file in place of the bytes it replaced - into the draft of
        // the board `identity`, where it still has the published bytes (ChangedFiles.CanPut - never over
        // a version of its own). Each goes in under a name of that board's own, of its own type
        // (ChangedFiles.NameFor), and every row there citing the old file is pointed at it
        // (ChangedFiles.Repointed) - ONE step of that draft named after this board, which its Undo takes
        // back (code review, 2026-10-26: written over the old bytes, it was no step, and a board file had
        // no way back). Never under a name climbing out of the draft's folder (DraftFiles.TryArrivedPath).
        // Each file not put is said in `outcome`; true when any went in.
        // ###########################################################################################
        private static async Task<bool> PutFilesIntoAsync(
            string identity,
            string boardName,
            IReadOnlyList<ChangedFileElsewhere> files,
            string fromIdentity,
            string fromName,
            AlsoOnOtherBoardsOutcome outcome)
        {
            string dataRoot = DataManager.DataRoot;

            if (await DraftEdits.OpenForChangeAsync(identity) is not { } session)
            {
                foreach (ChangedFileElsewhere file in files)
                    outcome.FileNotPut(file.File.DisplayText, boardName, ComponentPaneWording.AlsoNoDraft);

                return false;
            }

            string folder = DraftFolderLayout.GetBoardFolder(session.DraftsRoot, identity);
            string fromFolder = DraftFolderLayout.GetBoardFolder(session.DraftsRoot, fromIdentity);
            List<string> taken = session.ListedFiles.Keys.Concat(DraftFiles.CitedNames(session.Content.Board)).ToList();

            var copies = new List<(string Source, string Name)>();
            var going = new List<(ChangedFileElsewhere File, string Name)>();

            foreach (ChangedFileElsewhere file in files)
            {
                ChangedFilePut put = ChangedFiles.CanPut(session.Content.Board, session.ListedFiles, file);

                if (put != ChangedFilePut.Put)
                {
                    outcome.FileNotPut(file.File.DisplayText, boardName, put switch
                    {
                        ChangedFilePut.AlreadySo => ComponentPaneWording.AlsoAlreadySo,
                        ChangedFilePut.OwnVersion => ComponentPaneWording.AlsoFileOwnVersion,
                        _ => ComponentPaneWording.AlsoFileNoLongerThere,
                    });

                    continue;
                }

                string name = ChangedFiles.NameFor(file, taken);

                if (!DraftFiles.TryArrivedPath(folder, name, out _))
                {
                    outcome.FileNotPut(file.File.DisplayText, boardName, ComponentPaneWording.AlsoFileNotCopied);
                    continue;
                }

                if (await Task.Run(() => DraftFileResolver.Resolve(dataRoot, fromFolder, file.File.DraftName, fromIdentity)) is not { } source)
                {
                    outcome.FileNotPut(file.File.DisplayText, boardName, ComponentPaneWording.AlsoFileNotFound);
                    continue;
                }

                taken.Add(name);
                copies.Add((source, name));
                going.Add((file, name));
            }

            if (going.Count == 0)
            {
                return false;
            }

            DraftWorkbookEditOutcome? saved = await DraftEdits.CopyInAndApplyAsync(
                identity,
                copies,
                DraftStepWording.FilesCopied(going.Select(put => put.File.File.RowName).ToList(), fromName),
                board => going.Aggregate(board, (changed, put) => ChangedFiles.Repointed(changed, put.File.NameThere, put.Name)));

            if (saved != DraftWorkbookEditOutcome.Saved)
            {
                string why = saved is null ? ComponentPaneWording.AlsoFileNotCopied : ComponentPaneWording.AlsoNotSaved;

                foreach ((ChangedFileElsewhere file, _) in going)
                    outcome.FileNotPut(file.File.DisplayText, boardName, why);

                return false;
            }

            // Each one as the draft now lists it - a copy that failed on its own is said, not counted.
            bool any = false;

            foreach ((ChangedFileElsewhere file, string name) in going)
            {
                if (session.ListedFiles.TryGetValue(name, out BoardDocument.StoredFileRow? listed) &&
                    string.Equals(listed.Sha256, file.File.DraftSha256, StringComparison.OrdinalIgnoreCase))
                {
                    any = true;
                }
                else
                {
                    outcome.FileNotPut(file.File.DisplayText, boardName, ComponentPaneWording.AlsoFileNotCopied);
                }
            }

            return any;
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
        private static async Task<KiCadNetChoices> ReadNetChoicesAsync(string dataRoot, string draftFolder, string excelDataFile) =>
            await TabDrafts.ReadBoardKiCadAsync(dataRoot, draftFolder, excelDataFile, "its important signals") is { } root
                ? KiCadNetNames.Choices(root.Pcb.SelectMany(pcb => pcb.Nets.List).Select(net => (net.Name, net.NormalizedName, net.IsNoNet)))
                : KiCadNetChoices.None;

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
            string publishedBoardFile = published is null ? string.Empty : this.PublishedBoardFileOf(entry);

            show(session, published, publishedBoardFile);
            this.thisComponentEntry = entry;

            // The draft's header follows the session now open.
            this.RefreshDrafts();
        }

        // ###########################################################################################
        // What "Show on board", lighting a net up and the picture tools share (code review, 2026-10-26:
        // three copies of it): what the pane holds saved first, then whether the board on screen has to
        // be read again first - the pane's saves left it behind (it catches up only once the pane closes
        // or the tab is left, and a reload drops a selection). The board on screen is handed that reload
        // here, so leaving the tab reads nothing again. Null while no pane is open.
        // ###########################################################################################
        private async Task<bool?> SaveAndHandOverBoardOnScreenAsync()
        {
            if (!this.IsComponentOpen)
            {
                return null;
            }

            await this.SaveComponentNowAsync();

            bool readBoardFirst = this.thisBoardOnScreenBehind is not null;
            this.thisBoardOnScreenBehind = null;

            return readBoardFirst;
        }

        // "Show on board": the Schematics tab with only the component selected.
        internal async Task ShowComponentOnBoardAsync(string label)
        {
            if (await this.SaveAndHandOverBoardOnScreenAsync() is not bool readBoardFirst)
            {
                return;
            }

            if (this.ShowOnBoardOverrideForTests is { } show)
            {
                show(label, readBoardFirst);
                return;
            }

            this.thisMainWindow?.ShowComponentOnBoard(label, readBoardFirst);
        }

        // For tests, in place of Main: lighting up a net on the board (whether the board on screen
        // has to be read again first).
        internal Action<string, bool>? ShowNetOverrideForTests { get; set; }

        // "Show on board" on an important signal or a suggested one (click-to-light-up): the Schematics
        // tab with only that net lit up.
        internal async Task ShowNetOnBoardAsync(string netName)
        {
            if (await this.SaveAndHandOverBoardOnScreenAsync() is not bool readBoardFirst)
            {
                return;
            }

            if (this.ShowNetOverrideForTests is { } show)
            {
                show(netName, readBoardFirst);
                return;
            }

            this.thisMainWindow?.ShowNetOnBoard(netName, readBoardFirst);
        }

        // "Mark components" and "Align KiCad" on a picture (slice 6 - contributor mode went): the
        // Schematics tab in that tool.
        internal async Task StartPictureToolAsync(PictureTool tool, string picture)
        {
            if (await this.SaveAndHandOverBoardOnScreenAsync() is not bool readBoardFirst)
            {
                return;
            }

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
