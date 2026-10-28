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
    // THE "KiCad data" HALF OF THE WINDOW: the board's KiCad files listed with their sizes, a project
    // folder imported (picked, dropped here, or dropped on the Contribute tab), a file removed, the
    // report of what the KiCad data lines up with, and the summary of what an import found, left out
    // and leaves to do (CRT 4.0 step 3, slice 3e - KiCadImportSummary). Moved out of
    // BoardFilesWindow.axaml.cs as it passed the size the project keeps a file under (code review,
    // 2026-10-28).
    //
    // Part of the BoardFilesWindow partial class - see BoardFilesWindow.axaml.cs for the file map.
    // ###########################################################################################
    public partial class BoardFilesWindow
    {
        // The match report currently being built, if any. Removing a KiCad file waits for it first:
        // the report READS these files (a 45 MB .kicad_pcb takes a moment), and Windows refuses to
        // delete a file another handle has open - so a Remove clicked while the window was still
        // parsing failed with a sharing violation. See RefreshReportAsync.
        private Task thisPendingReport = Task.CompletedTask;

        // The last KiCad import made in this window that copied something: the files taken, every
        // other file of the folder, and which an older KiCad wrote - for the summary's "Left out"
        // (KiCadImportSummary). Not stored anywhere else: the folder is a path on the contributor's
        // machine, said of once, as it was imported.
        private KiCadFolderScan? thisLastImport;

        // ###########################################################################################
        // Says what KiCad data this board actually has - a count, and then the files themselves.
        //
        // *** IT NAMES THE FILES, because a count alone told a returning contributor nothing
        // (owner report, 2026-09-24). *** Reopening this window after an import showed
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
            int count = snapshot.Imported.Count;

            this.KiCadStateText.Text = count == 0
                ? "No KiCad data imported yet."
                : count == 1
                    ? "1 KiCad file imported:"
                    : $"{count} KiCad files imported:";

            this.KiCadFiles.Clear();

            foreach ((string file, string size) in snapshot.Imported)
            {
                this.KiCadFiles.Add(new BoardKiCadFileRow
                {
                    RelativePath = BoardFilesWindow.RelativeToFolder(snapshot.Copy, file),
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
        // NEVER made on the UI thread": the window froze, its "please wait" unable to paint. One read
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
                return BoardFilesWindow.ReadKiCad(dataRoot, draftFolder, excelDataFile, draftBoardFile);
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
                .Select(path => (path, BoardFilesWindow.DescribeSize(path)))
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
                var topLevel = GetTopLevel(this);
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
        // The board on screen is reloaded when this window closes (TabDrafts.ManageFilesAsync), the
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

        // One of the window's draft changes, so closing waits for it (BoardFilesWindow.DraftChanges.cs).
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
            // *** THE WALK, THE COPY AND THE REPORT ARE ONE OF THE WINDOW'S CHANGES, UNDER ONE WAIT (code
            // review, 2026-10-28). *** The folder is walked off the UI thread - a folder dropped can be
            // large - and walked outside the queue, the window could be closed meanwhile: closing waits
            // only for what is queued, and the import went into the draft after it, the board on screen
            // never read again. The window is held across all of it, so it does not brighten between
            // them (2026-09-28); closing waits for it (BoardFilesWindow.DraftChanges.cs).
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
                    // to bring it in.
                    this.ShowStatus(KiCadImportSummary.NothingToImport(scan.LeftOut, scan.OlderKiCad), isError: true);
                    return;
                }

                // Re-importing OVERWRITES files a report may still be reading - the same sharing
                // violation a Remove would hit, so the same wait (see thisPendingReport).
                await this.WaitForPendingReportAsync();

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
                // fails costs nothing: the files stay where they were copied, found and taken in by the
                // next write, and Submit carries them meanwhile (SubmissionKiCadFiles.Collect).
                // ###########################################################################################
                await BusyOverlay.RunLocalAsync(this, CrtWaitWording.ImportingKiCad, () => DraftEdits.TakeInNewFilesAsync(excelDataFile));

                // ###########################################################################################
                // *** EVEN A COPY THAT STOPPED PART-WAY CHANGED THE DRAFT (code review, 2026-10-10). ***
                // The files before the one that failed have been overwritten, so the board on screen
                // is read again when the window closes - it used to go on drawing the old project.
                // ###########################################################################################
                this.ChangedTheDraft = true;

                string imported = sourceFiles.Count == 1 ? "Imported 1 KiCad file" : $"Imported {sourceFiles.Count} KiCad files";
                string checking = sourceFiles.Count == 1 ? $"{imported}. Checking what it lines up with..." : $"{imported}. Checking what they line up with...";

                if (complete)
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
        // signal the Schematics tab draws - and the CAD names on offer differed from the
        // "Schematic images" window's.
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
                root.Project.Views.Select(view => view.DisplayName));

            this.ShowReport(report);

            // The net names the "Important signals" sheet can use, from EVERY PCB file - the
            // Schematics tab's panel reads them all (BoardFilesWindow.NetNames.cs).
            this.ShowNetNameReport(KiCadNetNames.Build(
                root.Pcb.SelectMany(pcb => pcb.Nets.List).Select(net => (net.Name, net.NormalizedName, net.IsNoNet)),
                this.thisImportantSignals));

            await this.ShowSummaryAsync(root);
        }

        // ###########################################################################################
        // THE SUMMARY (CRT 4.0 step 3, slice 3e - KiCadImportSummary): what the KiCad data holds, what
        // the import made in this window left out - and why, for a project from an older KiCad - and
        // what is left to do, by the getting-ready checklist's rules on the draft held in memory -
        // opened as the checklist opens it, none said for a draft that cannot be. The counts are
        // worked out off the UI thread.
        // ###########################################################################################
        private async Task ShowSummaryAsync(KiCadProjectRoot root)
        {
            string draftsRoot = DraftManager.DraftsRoot;
            string identity = this.thisExcelDataFile;

            (string found, List<string> toDo, bool known) = await Task.Run(() =>
            {
                string foundLine = KiCadImportSummary.Found(root);
                BoardContent? content = (DraftManager.Sessions.Find(draftsRoot, identity) ?? DraftManager.Sessions.Get(draftsRoot, identity))?.Content;

                List<string> lines = content is null
                    ? []
                    : KiCadImportSummary.LeftToDo(BoardReadiness.Steps(string.Empty, content, hasKiCadFiles: true, BoardKiCadFacts.Of(root)));

                return (foundLine, lines, content is not null);
            });

            if (this.thisClosed)
            {
                return;
            }

            this.KiCadFoundText.Text = found;

            string leftOut = this.thisLastImport is { } import ? KiCadImportSummary.LeftOut(import.LeftOut, import.OlderKiCad) : string.Empty;
            this.KiCadLeftOutText.Text = leftOut;
            this.KiCadLeftOutText.IsVisible = leftOut.Length > 0;

            // What of the project CRT cannot read - an older KiCad's files, another program's - and what to do.
            string cannotRead = this.thisLastImport is { } last
                ? string.Join(" ", KiCadImportSummary.CannotRead(last.LeftOut, last.OlderKiCad, last.Taken))
                : string.Empty;
            this.KiCadOlderText.Text = cannotRead;
            this.KiCadOlderText.IsVisible = cannotRead.Length > 0;

            this.KiCadLeftToDoPanel.IsVisible = known;
            this.KiCadLeftToDoText.Text = toDo.Count == 0
                ? KiCadImportSummary.NothingLeftToDo
                : string.Join("\n", toDo.Select(line => "- " + line));

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
        // but that section sits ABOVE this one, and the window's own scroller now carries a long
        // list the way it carries everything else.
        // ###########################################################################################
        internal void ShowReport(KiCadReferenceMatchReport report)
        {
            this.ReportPanel.IsVisible = true;
            this.ReportSummaryText.Text = report.Summary;

            this.UnmatchedPanel.IsVisible = report.UnmatchedLabels.Count > 0;
            this.UnmatchedHeaderText.Text = report.UnmatchedLabels.Count == 1
                ? "1 component will NOT light up"
                : $"{report.UnmatchedLabels.Count} components will NOT light up";
            this.UnmatchedList.ItemsSource = report.UnmatchedLabels;

            this.MatchedPanel.IsVisible = report.MatchedLabels.Count > 0;
            this.MatchedHeaderText.Text = report.MatchedLabels.Count == 1
                ? "1 component matches"
                : $"{report.MatchedLabels.Count} components match";
            this.MatchedList.ItemsSource = report.MatchedLabels;

            this.UnusedPanel.IsVisible = report.UnusedReferences.Count > 0;
            this.UnusedHeaderText.Text = report.UnusedReferences.Count == 1
                ? "1 component is not labelled yet"
                : $"{report.UnusedReferences.Count} components are not labelled yet";
            this.UnusedList.ItemsSource = report.UnusedReferences;

            this.ViewNamesPanel.IsVisible = report.ViewDisplayNames.Count > 0;
            this.ViewNamesListText.Text = string.Join("\n", report.ViewDisplayNames);
        }
    }
}
