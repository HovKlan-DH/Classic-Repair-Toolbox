using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Platform.Storage;
using Handlers.DataHandling;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace CRT
{
    // ###########################################################################################
    // THE "KiCad data" HALF OF THE PANEL: the board's KiCad files listed with their sizes, a project
    // folder imported (picked or dropped here), a file removed, the
    // report of what the KiCad data lines up with, and the summary of what an import found, left out
    // and leaves to do (CRT 4.0 step 3, slice 3e - KiCadImportSummary). Moved out of
    // KiCadDataPanel.axaml.cs as it passed the size the project keeps a file under (code review,
    // 2026-10-28).
    //
    // Part of the KiCadDataPanel partial class - see KiCadDataPanel.axaml.cs for the file map.
    // ###########################################################################################
    public partial class KiCadDataPanel
    {
        // The match report currently being built, if any. Removing a KiCad file waits for it first:
        // the report READS these files (a 45 MB .kicad_pcb takes a moment), and Windows refuses to
        // delete a file another handle has open - so a Remove clicked while the panel was still
        // parsing failed with a sharing violation. See RefreshReportAsync.
        private Task thisPendingReport = Task.CompletedTask;

        // The last KiCad import made in this panel that copied something: the files taken, every
        // other file of the folder, and which an older KiCad wrote - for the summary's "Left out"
        // (KiCadImportSummary). Not stored anywhere else: the folder is a path on the contributor's
        // machine, said of once, as it was imported.
        private KiCadFolderScan? thisLastImport;

        // ###########################################################################################
        // Says what KiCad data this board actually has - a count, and then the files themselves.
        //
        // *** IT NAMES THE FILES, because a count alone told a returning contributor nothing
        // (owner report, 2026-09-24). *** Reopening this panel after an import showed
        // "25 KiCad files imported" and not one word about WHICH project that was, so there was no
        // way to tell an import of the right folder from an import of the wrong one - or to notice
        // that the pages sub-folder had come along. The files are the only durable record: the
        // folder the contributor picked is not stored anywhere, deliberately, since it is a path on
        // their machine that may not exist tomorrow.
        //
        // Paths are shown RELATIVE to the "KiCad data" folder, so a multi-sheet project reads as
        // "Pages/vic.kicad_sch" rather than an absolute path nobody can scan down.
        //
        // Shows what ReadKiCadAsync read off the UI thread - nothing here touches the disk.
        // ###########################################################################################
        private void ShowKiCadState(KiCadSnapshot snapshot)
        {
            this.KiCadStateText.Text = KiCadImportedFiles.StateLine(snapshot.Imported.Count, snapshot.PublishedCount);

            this.KiCadFiles.Clear();

            foreach ((string file, string size) in snapshot.Imported)
            {
                this.KiCadFiles.Add(new BoardKiCadFileRow
                {
                    RelativePath = KiCadDataPanel.RelativeToFolder(snapshot.Copy, file),
                    SizeText = size,
                    FullPath = file,
                });
            }

            this.KiCadFilesPanel.IsVisible = this.KiCadFiles.Count > 0;
        }

        // ###########################################################################################
        // WHAT THE KiCad SECTION SHOWS, READ OFF THE UI THREAD (code review, 2026-10-25). Making the
        // board's KiCad copy (BoardKiCadFolder - tens of megabytes on a large board, copied again
        // whenever an import changes it) ran on the UI thread, twice per refresh, against "the copy is
        // NEVER made on the UI thread": the panel froze, its "please wait" unable to paint. One read
        // now brings everything the list and the report need: the copy, the draft's own files in it
        // with their sizes, every file in it, and how many the published board has.
        // ###########################################################################################
        private sealed record KiCadSnapshot(
            string Copy,
            IReadOnlyList<(string Path, string Size)> Imported,
            IReadOnlyList<string> Files,
            int PublishedCount);

        private Task<KiCadSnapshot> ReadKiCadAsync()
        {
            string dataRoot = this.DataRoot;
            string draftFolder = this.thisDraftFolder;
            string excelDataFile = this.thisExcelDataFile;
            string draftBoardFile = this.DraftBoardFile;
            Action? reading = this.ReadingKiCadForTests;

            return Task.Run(() =>
            {
                reading?.Invoke();
                return KiCadDataPanel.ReadKiCad(dataRoot, draftFolder, excelDataFile, draftBoardFile);
            });
        }

        // Called where the KiCad read runs, just before it - for a test proving that is not the UI thread.
        internal Action? ReadingKiCadForTests { get; set; }

        // Called where a picked or dropped folder is walked, just before it - for a test holding it there.
        internal Action? ScanningKiCadForTests { get; set; }

        internal bool HasLastImportForTests => this.thisLastImport is not null;

        private static KiCadSnapshot ReadKiCad(string dataRoot, string draftFolder, string excelDataFile, string draftBoardFile)
        {
            string publishedBoardFile = DraftBoardSource.PublishedPathOf(dataRoot, excelDataFile);

            // How many KiCad files the published board has - for the report's source line, counted from
            // its list without copying them (code review, 2026-10-25: a second full copy, for a count).
            int publishedCount = BoardKiCadFolder.ListedNames(publishedBoardFile).Count;

            string copy = BoardKiCadFolder.ForBoard(dataRoot, draftFolder, excelDataFile, AppConfig.KiCadCacheFolder);

            if (copy.Length == 0)
            {
                return new KiCadSnapshot(string.Empty, [], [], publishedCount);
            }

            List<string> files = KiCadRawFileScanner.Scan(copy);

            // The KiCad files the draft brought itself - imported, or replaced - as files under their
            // names in the copy, whichever of the two places their bytes are in now.
            //
            // *** IN CRT 4.0 A DRAFT NAMES ITS KiCad FILES IN ITS BOARD FILE (2026-10-24). *** An import
            // lies in "<draft>/KiCad data" only until the draft's next write takes it into the draft's
            // store, and a seeded draft lists the published board's KiCad files too - so the files
            // listed are the draft's own by name (BoardKiCadFolder.ImportedNames). Sub-folders
            // included - see KiCadRawFileScanner: a multi-sheet project keeps most of its sheets under
            // Pages/, and counting only the top level reported "3 KiCad files imported" for 25.
            string copyRoot = Path.GetDirectoryName(copy)!;
            var importedNames = BoardKiCadFolder
                .ImportedNames(draftBoardFile, publishedBoardFile)
                .Select(name => Path.GetFullPath(Path.Combine(copyRoot, name.Replace('/', Path.DirectorySeparatorChar))))
                .ToHashSet(StringComparer.OrdinalIgnoreCase);

            List<(string Path, string Size)> imported = files
                .Where(path => importedNames.Contains(Path.GetFullPath(path)))
                .Select(path => (path, KiCadDataPanel.DescribeSize(path)))
                .ToList();

            return new KiCadSnapshot(copy, imported, files, publishedCount);
        }

        // The file path as it reads inside the "KiCad data" folder, with forward slashes so a
        // sub-folder looks the way the Wiki and the KiCad project itself write it.
        private static string RelativeToFolder(string root, string fullPath)
        {
            if (string.IsNullOrWhiteSpace(root))
            {
                return Path.GetFileName(fullPath);
            }

            try
            {
                return Path.GetRelativePath(root, fullPath).Replace(Path.DirectorySeparatorChar, '/');
            }
            catch (Exception)
            {
                return Path.GetFileName(fullPath);
            }
        }

        // ###########################################################################################
        // One file's size, in the unit a person would say it in.
        //
        // Worth showing because size is the one cue that separates a real board from a stub: a
        // .kicad_pcb is tens of megabytes, and a 2 KB one means the export went wrong. Returns an
        // empty string rather than throwing when the file cannot be measured - it has just been
        // listed from disk, so that is close to impossible, but a preview line is not worth an
        // exception.
        // ###########################################################################################
        private static string DescribeSize(string fullPath)
        {
            try
            {
                long bytes = new FileInfo(fullPath).Length;

                if (bytes >= 1024L * 1024L)
                {
                    return $"{bytes / (1024.0 * 1024.0):0.#} MB";
                }

                return bytes >= 1024L
                    ? $"{bytes / 1024.0:0.#} KB"
                    : $"{bytes} bytes";
            }
            catch (Exception)
            {
                return string.Empty;
            }
        }

        // ###########################################################################################
        // The draft's own "KiCad data" folder - at the draft folder's ROOT, exactly where a
        // published board kept it (Phase 6, 2026-09-23). Where an import is copied to; the draft's
        // next write takes what lies there into its store (DraftFiles).
        // ###########################################################################################
        private string KiCadImportFolder =>
            string.IsNullOrWhiteSpace(this.thisDraftFolder)
                ? string.Empty
                : Path.Combine(this.thisDraftFolder, AppConfig.KiCadDataFolderName);

        // ------------------------------------------------------------------ KiCad import

        private async void OnImportKiCadClick(object? sender, RoutedEventArgs e)
        {
            try
            {
                var topLevel = TopLevel.GetTopLevel(this);
                if (topLevel == null)
                {
                    return;
                }

                var folders = await topLevel.StorageProvider.OpenFolderPickerAsync(new FolderPickerOpenOptions
                {
                    Title = "Select the KiCad project folder",
                    AllowMultiple = false,
                });

                if (folders == null || folders.Count == 0)
                {
                    return;
                }

                await this.ImportKiCadFolderAsync(folders[0].Path.LocalPath);
            }
            catch (Exception ex)
            {
                // Nothing may leave an async void handler: App lets such an exception end CRT
                // (code review, 2026-10-13 - AsyncVoidHandlersGuardTests).
                Logger.Critical($"OnImportKiCadClick failed - [{ex}]");
            }
        }

        private async void OnKiCadDropped(object? sender, DragEventArgs e)
        {
            try
            {
                string? folder = GetDroppedPaths(e).FirstOrDefault(Directory.Exists);

                if (!string.IsNullOrWhiteSpace(folder))
                {
                    await this.ImportKiCadFolderAsync(folder);
                }
            }
            catch (Exception ex)
            {
                // Nothing may leave an async void handler: App lets such an exception end CRT
                // (code review, 2026-10-13 - AsyncVoidHandlersGuardTests).
                Logger.Critical($"OnKiCadDropped failed - [{ex}]");
            }
        }

        // ###########################################################################################
        // Removes one imported KiCad file from the draft (owner request, 2026-09-24), then
        // refreshes the listing and the match report so both describe what is left.
        //
        // The FILE is deleted, not just hidden - see KiCadImportedFiles for why a KiCad file,
        // unlike a schematic image, has no row to drop instead. No confirmation: only the draft's
        // copy goes, and re-importing the KiCad folder brings it straight back, so a prompt per file
        // would cost more than a mis-click does. The status line says so.
        //
        // The board on screen is read again as the removal ends (DraftChanged - TabDrafts.KiCad.cs), the
        // same as after an import, so the removed file's traces disappear then.
        // ###########################################################################################
        private async void OnRemoveKiCadFileClick(object? sender, RoutedEventArgs e)
        {
            try
            {
                if (sender is Button { Tag: BoardKiCadFileRow row })
                {
                    await this.RemoveKiCadFileAsync(row);
                }
            }
            catch (Exception ex)
            {
                // Nothing may leave an async void handler: App lets such an exception end CRT
                // (code review, 2026-10-13 - AsyncVoidHandlersGuardTests).
                Logger.Critical($"OnRemoveKiCadFileClick failed - [{ex}]");
            }
        }

        // ---- the KiCad board's components added to the board ---------------------------------------

        // ###########################################################################################
        // *** AN IMPORT ADDS THE KiCad BOARD'S COMPONENTS (owner request, 2026-11-01: "I instead do
        // think it should add all components when importing the KiCad data. Then you can delete this
        // option from 'Board status' ... Have checkbox if it should add all data automatically also?").
        // *** While "Add the KiCad components automatically to the board" is ticked - remembered, ticked
        // until unticked - an import adds every component of the KiCad board the board does not have, by the
        // Contribute tab's one rule (TabDrafts.AddFromKiCadToAsync: ONE step, pins and datasheet links,
        // saved at once), handed in as AddKiCadComponents. What came of it is said in the summary, and
        // the components still missing - imported unticked, or deleted since - are one button away
        // there ("Add N components from KiCad", Board status's until then).
        // ###########################################################################################
        internal Func<Task<(string? Note, bool Changed)>>? AddKiCadComponents { get; set; }

        // The check box, as the panel opens with it - set before Initialize.
        internal bool AddsComponentsOnImport
        {
            get => this.AddComponentsCheckBox.IsChecked == true;
            set
            {
                this.thisShowingAddsComponents = true;

                try
                {
                    this.AddComponentsCheckBox.IsChecked = value;
                }
                finally
                {
                    this.thisShowingAddsComponents = false;
                }
            }
        }

        // Told when the contributor ticks or unticks it - remembered for the next import.
        internal Action<bool>? AddsComponentsOnImportChanged { get; set; }

        private bool thisShowingAddsComponents;

        // What the last add said (BoardReadiness.AddedFromKiCad) - shown in the summary.
        private string? thisComponentsAddedNote;

        internal string? ComponentsAddedNoteForTests => this.KiCadComponentsAddedText.IsVisible ? Handlers.Theming.BoldRuns.TextOf(this.KiCadComponentsAddedText) : null;

        private void OnAddComponentsCheckedChanged(object? sender, RoutedEventArgs e)
        {
            if (!this.thisShowingAddsComponents)
            {
                this.AddsComponentsOnImportChanged?.Invoke(this.AddComponentsCheckBox.IsChecked == true);
            }
        }

        private async void OnAddKiCadComponentsClick(object? sender, RoutedEventArgs e)
        {
            try
            {
                await this.AddKiCadComponentsNowAsync();
            }
            catch (Exception ex)
            {
                // Nothing may leave an async void handler: App lets such an exception end CRT.
                Logger.Critical($"OnAddKiCadComponentsClick failed - [{ex}]");
            }
        }

        // The summary's button: one of the panel's draft changes, the report read again after it.
        internal Task AddKiCadComponentsNowAsync() =>
            this.ChangeDraftAsync(() => BusyOverlay.HoldAsync(this, CrtWaitWording.AddingKiCadComponents, async () =>
            {
                await this.AddComponentsFromKiCadAsync();
                await BusyOverlay.RunLocalAsync(this, CrtWaitWording.CheckingKiCadMatches, this.RefreshReportAsync);
            }));

        // The components added - inside a change of the panel's own - and the board's labels the report
        // judges the KiCad data against following them.
        private async Task AddComponentsFromKiCadAsync()
        {
            if (this.AddKiCadComponents is not { } add)
            {
                return;
            }

            (string? note, bool changed) = await add();
            this.thisComponentsAddedNote = note;

            if (!changed)
            {
                return;
            }

            this.ChangedTheDraft = true;
            this.ReadBoardFromDraft();
        }

        // The board's labels and "Important signals" rows the report judges the KiCad data against, as
        // the draft held in memory has them now.
        private void ReadBoardFromDraft()
        {
            if (DraftManager.Sessions.Find(DraftManager.DraftsRoot, this.thisExcelDataFile) is not { } session)
            {
                return;
            }

            BoardData board = session.Content.Board;

            this.thisBoardLabels = board.Components
                .Select(component => component.BoardLabel)
                .Where(label => !string.IsNullOrWhiteSpace(label))
                .ToList();
            this.thisImportantSignals = board.KiCadImportantSignals;
        }

        // ###########################################################################################
        // *** THE DRAFT MOVED UNDER THE PANEL (code review, 2026-11-01). *** It was a modal window until
        // 2026-10-30; a panel stays on screen beside the header's Undo and Redo and "Review changes"'s
        // "Undo changes to", and went on saying the components were added, its "Add N components from
        // KiCad" hidden and its report judging the KiCad data against the labels from before. What it
        // shows is read again - queued with its own changes, none of which it is: nothing is told
        // (DraftChanged), the tab having caught up already. What the last add said goes with it.
        // ###########################################################################################
        internal void FollowDraft() => _ = this.FollowDraftAsync();

        private async Task FollowDraftAsync()
        {
            try
            {
                await this.ChangeDraftAsync(() => BusyOverlay.RunLocalAsync(this, CrtWaitWording.CheckingKiCadMatches, () =>
                {
                    this.thisComponentsAddedNote = null;
                    this.ReadBoardFromDraft();

                    return this.RefreshReportAsync();
                }));
            }
            catch (Exception ex)
            {
                // Started without an awaiting caller: nothing may escape it.
                Logger.Warning($"The KiCad data panel could not follow the draft - [{ex.Message}]");
            }
        }

        // One of the panel's draft changes, so closing waits for it (KiCadDataPanel.DraftChanges.cs).
        internal Task RemoveKiCadFileAsync(BoardKiCadFileRow row) =>
            this.ChangeDraftAsync(async () =>
            {
                // A report still parsing these files would hold this one open - see thisPendingReport.
                await this.WaitForPendingReportAsync();

                // Taken off the draft's list by its name - the bytes go with it, from the draft's
                // store or from where the import left them (DraftWorkbookStore.RemoveFiles). A write of
                // the file list, which is no part of the draft's content, so it runs under the draft
                // held in memory - saved first, read back after (DraftEdits.WriteFileDirectly).
                string name = $"{AppConfig.KiCadDataFolderName}/{row.RelativePath}";
                string draftsRoot = DraftManager.DraftsRoot;
                string excelDataFile = this.thisExcelDataFile;

                // A file the published board has goes back to the published one (DraftWorkbookStore.RemoveFiles).
                string publishedBoardFile = DraftBoardSource.PublishedPathOf(this.DataRoot, excelDataFile);
                bool backToPublished = false;
                bool removed;

                try
                {
                    removed = await BusyOverlay.RunLocalAsync(
                        this,
                        CrtWaitWording.SavingDraft,
                        () => DraftEdits.WriteFileDirectlyAsync(
                            excelDataFile,
                            () =>
                            {
                                backToPublished = DraftFiles.ListedIn(publishedBoardFile)?.ContainsKey(name) == true;
                                return DraftWorkbookStore.RemoveFiles(draftsRoot, excelDataFile, [name], publishedBoardFile);
                            },
                            succeeded => succeeded,
                            DraftStepWording.KiCadFileRemoved,
                            clearHistory: false));
                }
                catch (DraftNotSavedException ex)
                {
                    this.ShowStatus($"Could not remove [{row.RelativePath}]. " + DraftSaveWording.EditNotSaved(ex.Failure), isError: true);
                    return;
                }

                if (!removed)
                {
                    this.ShowStatus($"Could not remove [{row.RelativePath}] - see the log for details.", isError: true);
                    return;
                }

                this.ChangedTheDraft = true;

                this.ShowStatus(KiCadImportedFiles.RemovedLine(row.RelativePath, backToPublished));

                // The list and the report, read again together (ReadKiCadAsync).
                await BusyOverlay.RunLocalAsync(this, CrtWaitWording.CheckingKiCadMatches, this.RefreshReportAsync);
            });

        // Builds the match report and records it as the one in flight, so a Remove can wait for it.
        private Task RefreshReportAsync()
        {
            this.thisPendingReport = this.BuildAndShowReportAsync();
            return this.thisPendingReport;
        }

        // A report that failed has already said so on the status line; its failure is no reason to
        // refuse the Remove that is waiting on it.
        private async Task WaitForPendingReportAsync()
        {
            try
            {
                await this.thisPendingReport;
            }
            catch (Exception ex)
            {
                Logger.Warning($"The KiCad match report failed before a file was removed - [{ex.Message}]");
            }
        }

        // ###########################################################################################
        // Copies the KiCad project's own files into the draft, then reports what they line up with.
        //
        // *** A CONTRIBUTOR MAY PICK THE WHOLE KiCad PROJECT FOLDER, and only the few relevant
        // files are taken (owner request, 2026-09-24). *** That is the EXTENSION filter's
        // doing: .kicad_pcb / .kicad_pro / .kicad_sch and nothing else, so footprint libraries
        // (.kicad_mod), 3D models (.step/.wrl), gerbers, netlists and backups are all left where
        // they are - at any depth.
        //
        // SUB-FOLDERS ARE INCLUDED, because a multi-sheet project puts its pages in one and the
        // import would otherwise take the root sheet and none of the circuitry. The sub-folder
        // structure is PRESERVED rather than flattened, since a root sheet references its pages by
        // relative path. The result is the same layout the shipped C128 board already has by hand.
        //
        // The rule is KiCadRawFileScanner's, shared with the reader that later loads the folder, so
        // what is imported is exactly what will be read.
        // ###########################################################################################
        internal Task ImportKiCadFolderAsync(string sourceFolder)
        {
            if (string.IsNullOrWhiteSpace(this.thisDraftFolder))
            {
                this.ShowStatus("Could not resolve where to save - no drafts folder for this board.", isError: true);
                return Task.CompletedTask;
            }

            string destinationFolder = this.KiCadImportFolder;
            string excelDataFile = this.thisExcelDataFile;
            Action? scanning = this.ScanningKiCadForTests;

            // ###########################################################################################
            // *** THE WALK, THE COPY AND THE REPORT ARE ONE OF THE PANEL'S CHANGES, UNDER ONE WAIT (code
            // review, 2026-10-28). *** The folder is walked off the UI thread - a folder dropped can be
            // large - and walked outside the queue, the panel could be closed meanwhile: closing waits
            // only for what is queued, and the import went into the draft after it, the board on screen
            // never read again. The window is held across all of it, so it does not brighten between
            // them (2026-09-28); closing waits for it (KiCadDataPanel.DraftChanges.cs).
            // ###########################################################################################
            return this.ChangeDraftAsync(() => BusyOverlay.HoldAsync(this, CrtWaitWording.ImportingKiCad, async () =>
            {
                KiCadFolderScan scan;

                try
                {
                    // What is taken, and - for the summary - what is left out, in ONE walk (code review,
                    // 2026-10-28: it was walked twice).
                    scan = await Task.Run(() =>
                    {
                        scanning?.Invoke();
                        return KiCadRawFileScanner.ScanSplit(sourceFolder);
                    });
                }
                catch (Exception ex)
                {
                    Logger.Warning($"Could not read the KiCad folder [{sourceFolder}] - [{ex.Message}]");
                    this.ShowStatus("Could not read that folder.", isError: true);
                    return;
                }

                IReadOnlyList<string> sourceFiles = scan.Taken;

                if (sourceFiles.Count == 0)
                {
                    // Said for what it is - an older KiCad's or another program's project is told how
                    // to bring it in, and KiCad files that hold no bytes are named (2026-11-02).
                    this.ShowStatus(
                        scan.Empty.Count > 0 ? KiCadImportSummary.EmptyFiles(scan.Empty) : KiCadImportSummary.NothingToImport(scan.LeftOut, scan.OlderKiCad),
                        isError: true);

                    return;
                }

                // Re-importing OVERWRITES files a report may still be reading - the same sharing
                // violation a Remove would hit, so the same wait (see thisPendingReport).
                await this.WaitForPendingReportAsync();

                // A draft gone from under the panel: the copy would make its folder again, with no board
                // file or marker, for nothing to take in (code review, 2026-11-02).
                if (!DraftBoardSource.HasDraft(DraftManager.DraftsRoot, excelDataFile))
                {
                    this.ShowStatus(KiCadImportSummary.NoDraftAnyMore, isError: true);
                    return;
                }

                (int copied, bool complete) = await this.WriteDraftAsync(
                    CrtWaitWording.ImportingKiCad,
                    () => KiCadImportedFiles.CopyIn(sourceFolder, sourceFiles, destinationFolder));

                if (copied == 0)
                {
                    this.ShowStatus("Could not copy those KiCad files - see the log for details.", isError: true);
                    return;
                }

                // Said of by the summary only now - an import that copied nothing imported nothing
                // (code review, 2026-10-28: recorded before the copy, a later report summed it up).
                this.thisLastImport = scan;

                // ###########################################################################################
                // *** AND LISTED IN THE DRAFT AT ONCE (code review, 2026-10-25). *** A write with nothing
                // else changed takes what was copied into the draft's list (DraftFiles), so everything
                // that reads the list - Submit, the Drafts row - agrees with the Schematics tab. Through
                // the draft held in memory when one is open (DraftEdits.TakeInNewFilesAsync). A write that
                // fails loses nothing while there is a draft: the files stay where they were copied, found
                // and taken in by the next write, and Submit carries them meanwhile
                // (SubmissionKiCadFiles.Collect). But it is said (code review, 2026-11-02): "Imported N KiCad
                // files" stood over a draft that had not taken them, or was no longer there at all.
                // ###########################################################################################
                bool takenIn = await BusyOverlay.RunLocalAsync(this, CrtWaitWording.ImportingKiCad, () => DraftEdits.TakeInNewFilesAsync(excelDataFile));

                if (!takenIn && !DraftBoardSource.HasDraft(DraftManager.DraftsRoot, excelDataFile))
                {
                    this.ShowStatus(KiCadImportSummary.NoDraftAnyMore, isError: true);
                    return;
                }

                // ###########################################################################################
                // *** EVEN A COPY THAT STOPPED PART-WAY CHANGED THE DRAFT (code review, 2026-10-10). ***
                // The files before the one that failed have been overwritten, so the board on screen
                // is read again as this change ends (DraftChanged) - it used to go on drawing the old project.
                // ###########################################################################################
                this.ChangedTheDraft = true;

                // And the KiCad board's components with them, while the check box says so (2026-11-01).
                if (this.AddsComponentsOnImport)
                {
                    await this.AddComponentsFromKiCadAsync();
                }

                string imported = sourceFiles.Count == 1 ? "Imported 1 KiCad file" : $"Imported {sourceFiles.Count} KiCad files";
                string checking = sourceFiles.Count == 1 ? $"{imported}. Checking what it lines up with..." : $"{imported}. Checking what they line up with...";

                if (!takenIn)
                {
                    this.ShowStatus(KiCadImportSummary.NotTakenIn(copied), isError: true);
                }
                else if (complete)
                {
                    this.ShowStatus(checking);
                }
                else
                {
                    this.ShowStatus(
                        $"Only {copied} of the {sourceFiles.Count} KiCad files could be copied - see the log for details. Import the folder again once the rest can be read.",
                        isError: true);
                }

                await BusyOverlay.RunLocalAsync(this, CrtWaitWording.CheckingKiCadMatches, this.RefreshReportAsync);

                // The check done: no longer "Checking..." under the finished summary (2026-10-28) -
                // unless the report said something of its own meanwhile.
                if (string.Equals(this.StatusText.Text, checking, StringComparison.Ordinal))
                {
                    this.ShowStatus($"{imported} - the summary above says what {(sourceFiles.Count == 1 ? "it holds" : "they hold")}.");
                }
            }));
        }

        // ###########################################################################################
        // Parses the board's KiCad data and shows the matched/unmatched report.
        //
        // The comparison is built from root.Pcb[0]'s footprint references, which is EXACTLY what
        // TabSchematics.KiCad.cs's BuildKiCadNormalizedNetNamesForReferences uses to turn a board
        // label into copper. Reporting anything else - schematic symbols, or every PCB file unioned
        // together - would claim matches that light up nothing, which is worse than saying nothing.
        //
        // *** THE FILES ARE THE ONES THE SCHEMATICS TAB READS (code review, 2026-10-10): the
        // board's KiCad project as the draft has it, the published files with the draft's own
        // (BoardKiCadCopy). *** Read from the draft's imports alone, a net or a footprint only the
        // published files have was reported as missing - "will NOT light up" for an important
        // signal the Schematics tab draws - and the view names on offer differed from those the
        // pictures could be matched to.
        // ###########################################################################################
        private async Task BuildAndShowReportAsync()
        {
            // The list of files and the report from one read off the UI thread (ReadKiCadAsync).
            KiCadSnapshot snapshot = await this.ReadKiCadAsync();

            if (this.thisClosed)
            {
                return;
            }

            this.ShowKiCadState(snapshot);

            IReadOnlyList<string> files = snapshot.Files;

            string source = KiCadImportedFiles.ReportSourceLine(snapshot.PublishedCount, snapshot.Imported.Count);
            this.ReportSourceText.Text = source;
            this.ReportSourceText.IsVisible = source.Length > 0;

            if (files.Count == 0)
            {
                this.ReportPanel.IsVisible = false;
                this.KiCadSummaryPanel.IsVisible = false;
                return;
            }

            var bundle = await KiCadProjectLoader.LoadRawAsync(files);
            var root = bundle?.Root;

            if (root == null)
            {
                this.ShowStatus("Those KiCad files could not be read.", isError: true);
                this.ReportPanel.IsVisible = false;
                this.KiCadSummaryPanel.IsVisible = false;
                return;
            }

            var footprintReferences = root.Pcb.Count == 0
                ? Enumerable.Empty<string>()
                : root.Pcb[0].Footprints
                    .Select(footprint => footprint.Reference ?? string.Empty);

            var report = KiCadReferenceMatcher.Build(
                this.thisBoardLabels,
                footprintReferences,
                root.Pcb.Count,
                root.Project.Views.Select(view => view.DisplayName),
                KiCadComponentTable.Of(root).Select(component => component.Label));

            this.ShowReport(report);

            // The net names the "Important signals" sheet can use, from EVERY PCB file - the
            // Schematics tab's panel reads them all (KiCadDataPanel.NetNames.cs).
            this.ShowNetNameReport(KiCadNetNames.Build(
                root.Pcb.SelectMany(pcb => pcb.Nets.List).Select(net => (net.Name, net.NormalizedName, net.IsNoNet)),
                this.thisImportantSignals));

            await this.ShowSummaryAsync(root);
        }

        // ###########################################################################################
        // THE SUMMARY (CRT 4.0 step 3, slice 3e - KiCadImportSummary): what the KiCad data holds, and what
        // the import made in this panel left out - and why, for a project from an older KiCad. What is
        // left to do on the board is "Board status"'s to say (owner request, 2026-11-03). The counts are
        // worked out off the UI thread.
        // ###########################################################################################
        private async Task ShowSummaryAsync(KiCadProjectRoot root)
        {
            string draftsRoot = DraftManager.DraftsRoot;
            string identity = this.thisExcelDataFile;

            string draftFolder = this.thisDraftFolder;

            (string found, string files, int notOnBoard) = await Task.Run(() =>
            {
                string foundLine = KiCadImportSummary.Found(root);
                DraftSession? session = DraftManager.Sessions.Find(draftsRoot, identity) ?? DraftManager.Sessions.Get(draftsRoot, identity);
                BoardContent? content = session?.Content;

                // The KiCad components the board still lacks - the summary's button adds them, by the
                // very rule it counts them with.
                int missing = content is null ? 0 : KiCadBoardComponents.NotOn(content.Board, KiCadComponentTable.Of(root)).Count;

                // The names the left column counts: the draft's list, and a file copied in and not
                // listed yet (TabDrafts.KiCadFileNamesOf's rule).
                IEnumerable<string> names = (session?.ListedFiles.Keys ?? Enumerable.Empty<string>())
                    .Concat(string.IsNullOrWhiteSpace(draftFolder) ? [] : DraftFiles.ArrivedKiCadNames(draftFolder));

                return (foundLine, KiCadImportSummary.Files(names), missing);
            });

            if (this.thisClosed)
            {
                return;
            }

            Handlers.Theming.BoldRuns.ShowMarked(this.KiCadFilesText, files);
            this.KiCadFilesText.IsVisible = files.Length > 0;
            this.KiCadFoundText.Text = found;

            string leftOut = this.thisLastImport is { } import
                ? string.Join(" ", new[] { KiCadImportSummary.LeftOut(import.LeftOut, import.OlderKiCad), KiCadImportSummary.EmptyFiles(import.Empty) }.Where(line => line.Length > 0))
                : string.Empty;

            this.KiCadLeftOutText.Text = leftOut;
            this.KiCadLeftOutText.IsVisible = leftOut.Length > 0;

            // What of the project CRT cannot read - an older KiCad's files, another program's - and what to do.
            string cannotRead = this.thisLastImport is { } last
                ? string.Join(" ", KiCadImportSummary.CannotRead(last.LeftOut, last.OlderKiCad, last.Taken))
                : string.Empty;
            this.KiCadOlderText.Text = cannotRead;
            this.KiCadOlderText.IsVisible = cannotRead.Length > 0;

            // What the import's "Add the KiCad components" did, and the rest one button away (2026-11-01).
            Handlers.Theming.BoldRuns.ShowMarked(this.KiCadComponentsAddedText, this.thisComponentsAddedNote);
            this.KiCadComponentsAddedText.IsVisible = !string.IsNullOrWhiteSpace(this.thisComponentsAddedNote);
            this.AddKiCadComponentsButton.Content = KiCadImportSummary.AddComponentsButton(notOnBoard);
            this.AddKiCadComponentsButton.IsVisible = notOnBoard > 0 && this.AddKiCadComponents is not null;

            this.KiCadSummaryPanel.IsVisible = true;
        }

        // ###########################################################################################
        // Renders one report. Each section is hidden entirely when it has nothing in it rather than
        // shown with a zero - an empty heading reads as a fault, and the unmatched section in
        // particular must only appear when there is genuinely something wrong.
        //
        // *** EVERY COMPONENT IS LISTED, EACH AS ITS OWN BADGE (owner request, 2026-09-24). ***
        // The lists used to be comma-separated prose, and the unlabelled one stopped after 50 with
        // "and 315 more" - on a new board that hid most of the to-do list the section exists to
        // show. The cap was there so a long list could not push the unmatched section off screen,
        // but that section sits ABOVE this one, and the panel's own scroller now carries a long
        // list the way it carries everything else.
        // ###########################################################################################
        internal void ShowReport(KiCadReferenceMatchReport report)
        {
            this.ReportPanel.IsVisible = true;
            this.ReportSummaryText.Text = report.Summary;

            this.UnmatchedPanel.IsVisible = report.UnmatchedLabels.Count > 0;
            this.UnmatchedHeaderText.Text = report.UnmatchedLabels.Count == 1
                ? "1 component's copper can NOT be highlighted"
                : $"{report.UnmatchedLabels.Count} components' copper can NOT be highlighted";
            this.UnmatchedList.ItemsSource = report.UnmatchedLabels;

            this.MatchedPanel.IsVisible = report.MatchedLabels.Count > 0;
            this.MatchedHeaderText.Text = report.MatchedLabels.Count == 1
                ? "1 component matches"
                : $"{report.MatchedLabels.Count} components match";
            this.MatchedList.ItemsSource = report.MatchedLabels;

            this.UnusedPanel.IsVisible = report.UnusedReferences.Count > 0;
            this.UnusedHeaderText.Text = report.UnusedReferences.Count == 1
                ? "1 component is not on the board yet"
                : $"{report.UnusedReferences.Count} components are not on the board yet";
            this.UnusedList.ItemsSource = report.UnusedReferences;

            this.ViewNamesPanel.IsVisible = report.ViewDisplayNames.Count > 0;
            this.ViewNamesHintText.Text = ComponentPaneWording.KiCadViewsHint;
            this.ViewNamesListText.Text = string.Join("\n", report.ViewDisplayNames);
        }
    }
}
