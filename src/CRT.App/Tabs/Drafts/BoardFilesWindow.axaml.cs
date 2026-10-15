using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Media.Imaging;
using Avalonia.Platform.Storage;
using Avalonia.VisualTree;
using Handlers.DataHandling;
using Handlers.Theming;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace CRT
{
    // ###########################################################################################
    // Importing a board's schematic images and KiCad data, and reporting what the KiCad data actually
    // lines up with (NewContributeStrategy.md Phase 2, session 2c, task 9).
    //
    // This replaces three manual steps from Assets/Wiki/Add-new-board-with-KiCad-data.md: copying
    // image files into the board folder by hand, creating a folder named exactly "KiCad data" and
    // filling it, and reading the generated CAD view names out of the LOGFILE. The report is the
    // piece with no equivalent at all today - a board label that does not match a KiCad reference
    // designator silently lights up no copper, which the Wiki itself calls "the number one cause of
    // 'I did everything and no traces appear'".
    //
    // Everything is copied, never moved (the same import-by-copy rule ComponentDraftWriter and
    // WorklogAttachmentWriter follow), into the board's own draft folder. Nothing is written into
    // "Data/", and no .xlsx is created - the schematic rows go into draft.json like every other
    // edit in this phase.
    //
    // FILE MAP:
    //   BoardFilesWindow.axaml.cs          - opening, the image and KiCad imports, the report
    //   BoardFilesWindow.DraftChanges.cs   - every change to the draft: queued, off the UI thread,
    //                                        under the "please wait"; and closing after them
    //   BoardFilesWindow.SchematicOrder.cs - dragging the schematic images into a new order
    //   BoardFilesWindow.CadNames.cs       - picking each image's CAD name from the KiCad views
    //   BoardFilesWindow.SchematicNames.cs - renaming an image (its "Schematic name")
    //   BoardFilesWindow.NetNames.cs       - the report's KiCad net names and Important signals
    // ###########################################################################################
    public partial class BoardFilesWindow : Window
    {
        // The width each row preview is decoded to. Small on purpose: this is a "does that look
        // right?" glance, not a viewer, and the source images are routinely 4000px+ across.
        private const int ThumbnailWidth = 160;

        private string thisExcelDataFile = string.Empty;
        private string thisDraftFolder = string.Empty;

        // The match report currently being built, if any. Removing a KiCad file waits for it first:
        // the report READS these files (a 45 MB .kicad_pcb takes a moment), and Windows refuses to
        // delete a file another handle has open - so a Remove clicked while the window was still
        // parsing failed with a sharing violation. See RefreshReportAsync.
        private Task thisPendingReport = Task.CompletedTask;

        // The board labels the KiCad report compares against. Supplied by the caller from the
        // MERGED BoardData (official plus draft), so this works the same on a brand-new board
        // (where every label is drafted) and on an existing board.
        private IReadOnlyList<string> thisBoardLabels = Array.Empty<string>();

        // The board's "Important signals" rows, from the same merged BoardData as the labels - the
        // report says which of them name a net the KiCad data does not have (2026-10-10).
        private IReadOnlyList<KiCadImportantSignalEntry> thisImportantSignals = Array.Empty<KiCadImportantSignalEntry>();

        public ObservableCollection<BoardSchematicRow> Schematics { get; } = new();

        // The imported KiCad files, listed so a returning contributor can see WHICH project is in
        // this board rather than only how many files it has.
        public ObservableCollection<BoardKiCadFileRow> KiCadFiles { get; } = new();

        public BoardFilesWindow()
        {
            this.InitializeComponent();

            this.SchematicsItemsControl.ItemsSource = this.Schematics;
            this.KiCadFilesItemsControl.ItemsSource = this.KiCadFiles;

            // Drag and drop as well as the pickers: dragging a folder of board scans in is the
            // natural gesture, and the Wiki's manual procedure was itself a series of file copies.
            DragDrop.SetAllowDrop(this.SchematicDropTarget, true);
            DragDrop.SetAllowDrop(this.KiCadDropTarget, true);

            this.SchematicDropTarget.AddHandler(DragDrop.DragOverEvent, OnDragOver);
            this.SchematicDropTarget.AddHandler(DragDrop.DropEvent, this.OnSchematicsDropped);
            this.KiCadDropTarget.AddHandler(DragDrop.DragOverEvent, OnDragOver);
            this.KiCadDropTarget.AddHandler(DragDrop.DropEvent, this.OnKiCadDropped);

            this.AddHandler(KeyDownEvent, this.OnWindowKeyDown, RoutingStrategies.Tunnel);

            this.WireSchematicRowDrag();
        }

        // ###########################################################################################
        // Points the window at one board, showing ONE of its two sections. boardLabels comes from
        // the merged BoardData the caller already has, rather than being re-derived here - the
        // caller is the one place that knows whether it is looking at a drafted-only board or an
        // official one with an overlay.
        //
        // *** ONE SECTION PER OPENING (owner request, 2026-09-24). *** The Drafts tab used to
        // open this window from a single "Schematic images and KiCad data" button showing both;
        // it now has a button for each. One window with a section rather than two windows, because
        // the two halves share every piece of plumbing here (the draft folder, the status line, the
        // drop handling) - but only the chosen half is shown AND loaded: opening "Schematic images"
        // must not parse a 45 MB .kicad_pcb for a report that is not on screen, and opening "KiCad
        // data" must not decode every board image for previews nobody sees.
        // ###########################################################################################
        public void Initialize(
            string displayName,
            string excelDataFile,
            IReadOnlyList<string> boardLabels,
            BoardFilesSection section,
            IReadOnlyList<KiCadImportantSignalEntry>? importantSignals = null)
        {
            this.thisExcelDataFile = excelDataFile;
            this.thisDraftFolder = DraftManager.GetBoardFolder(excelDataFile);
            this.thisBoardLabels = boardLabels;
            this.thisImportantSignals = importantSignals ?? Array.Empty<KiCadImportantSignalEntry>();

            this.HeaderText.Text = displayName;

            bool isKiCad = section == BoardFilesSection.KiCadData;

            this.Title = isKiCad ? "KiCad data" : "Schematic images";
            this.IntroText.Text = isKiCad
                ? "Add the KiCad project that backs this board's schematic images. The files are copied into your own local draft - your original project is left where it is."
                : "Add the schematic images you want to work on. Each one is copied into your own local draft - the original files are left where they are.";

            this.SchematicsSection.IsVisible = !isKiCad;
            this.KiCadSection.IsVisible = isKiCad;

            // ###########################################################################################
            // *** WHAT IS READ ON OPENING IS READ UNDER THE "PLEASE WAIT" (owner request, 2026-10-10:
            // "it can take a few secs to display, so please do show that new standard wait
            // method"). *** The draft workbook, every image's preview and the KiCad data used to be
            // read on the UI thread, so the window froze - and, opened from the Drafts tab, appeared
            // seconds after the click with nothing to say why. Now each is read off the UI thread
            // while the window's BusyOverlay holds it; Opening completes when all of it is in.
            // ###########################################################################################
            if (!isKiCad)
            {
                this.Opening = BusyOverlay.HoldAsync(this, CrtWaitWording.OpeningSchematicImages, async () =>
                {
                    // The CAD names come from the KiCad data, read without parsing the PCB - see
                    // BoardFilesWindow.CadNames.cs - alongside the images.
                    this.CadNamesLoading = this.LoadCadViewNamesAsync();

                    await this.ReloadSchematicsFromDraftAsync();
                    await this.CadNamesLoading;
                });

                return;
            }

            this.RefreshKiCadState();

            // ###########################################################################################
            // *** AND REBUILD THE REPORT FOR DATA THAT IS ALREADY THERE (owner report,
            // 2026-09-24). ***
            //
            // The match report used to be built only at the end of an import, so reopening the
            // window on a board that already had KiCad data showed an empty panel - the one piece
            // of information that says whether the board labels line up with the KiCad references
            // was available exactly once, and only to whoever performed the import. Parsing a
            // 45 MB .kicad_pcb takes a moment, so it is under the "please wait" too.
            // ###########################################################################################
            this.Opening = BusyOverlay.HoldAsync(this, CrtWaitWording.CheckingKiCadMatches, this.RefreshReportAsync);
        }

        // Completes when everything the window reads on opening is in - for tests, which wait for it
        // as the user does.
        internal Task Opening { get; private set; } = Task.CompletedTask;

        // ###########################################################################################
        // Whether anything in this window changed the draft - an image added, removed or moved, a
        // CAD name picked, KiCad files imported or removed. The Drafts tab reloads the board on
        // screen after the window closes ONLY then (owner report, 2026-10-10: closing the window
        // refreshed the detached schematic thumbnails for no reason - it reloaded the board every
        // time, changed or not). Set only once the change is on disk.
        // ###########################################################################################
        internal bool ChangedTheDraft { get; private set; }

        // Set by OnClosed: a read still in flight then has nowhere to show what it read.
        private bool thisClosed;

        // The draft's schematic rows read again, under the "please wait" - for a change made in
        // this window that has to show what the draft now holds.
        private Task ReloadSchematicsUnderOverlayAsync() =>
            BusyOverlay.RunLocalAsync(this, CrtWaitWording.ReadingSchematicImages, this.ReloadSchematicsFromDraftAsync);

        // ###########################################################################################
        // Rebuilds the schematic list from the draft on disk, so what is shown is always what was
        // actually saved rather than an in-memory copy that could drift from it.
        //
        // The workbook and every preview are read OFF the UI thread (2026-10-10) - a couple of
        // dozen board scans took seconds to decode, with the window frozen - and the rows on screen
        // are swapped for the new ones only once all of it is in. Callers run it under the window's
        // "please wait" (Opening, ReloadSchematicsUnderOverlayAsync).
        // ###########################################################################################
        private async Task ReloadSchematicsFromDraftAsync()
        {
            string draftsRoot = DraftManager.DraftsRoot;
            string dataRoot = this.DataRoot;
            string draftFolder = this.thisDraftFolder;
            string excelDataFile = this.thisExcelDataFile;

            // Reads the draft's own BOARD rather than a delta list (Phase 6): the schematics ARE
            // the workbook's rows, so there is no Deleted state to skip and no loosely-typed Entry
            // to type-check - a row that is gone is simply not there.
            List<(BoardSchematicEntry Entry, Bitmap? Thumbnail)> read = await Task.Run(() =>
                (DraftWorkbookStore.LoadDraftBoard(draftsRoot, excelDataFile)?.Schematics ?? [])
                    .Select(entry => (entry, BoardFilesWindow.LoadThumbnail(dataRoot, draftFolder, entry.SchematicImageFile)))
                    .ToList());

            if (this.thisClosed)
            {
                // Nothing will show these, and OnClosed has already released what it could see.
                foreach ((_, Bitmap? thumbnail) in read)
                {
                    thumbnail?.Dispose();
                }

                return;
            }

            // A drag cannot survive its rows being replaced - it would hold a row no longer listed.
            this.ResetSchematicRowDrag();

            // ###########################################################################################
            // *** A NAME STILL BEING TYPED STAYS IN ITS BOX (code review, 2026-10-10). *** Images
            // dropped in while a Name box held an unsaved name replaced every row, and the typed name
            // went with its old row, without a word. It is handed to the new row of the same image
            // (SchematicRows: name, file, place), still unsaved - Enter, leaving the box or closing
            // the window saves it as before.
            // ###########################################################################################
            List<(SchematicRowRef Row, string Typed)> typed = this.Schematics
                .Where(row => row.HasNameEdit)
                .Select(row => (this.RefOf(row), row.NameDraft))
                .ToList();

            // Before the Clear, or the rows holding them are gone before they can be released.
            this.DisposeThumbnails();

            this.Schematics.Clear();

            var rows = read
                .Select(item => new BoardSchematicRow
                {
                    SchematicName = item.Entry.SchematicName,
                    ImageFile = item.Entry.SchematicImageFile,
                    CadName = item.Entry.CadName,
                    Thumbnail = item.Thumbnail,
                    CadNameChosen = this.OnCadNameChosen,
                })
                .ToList();

            foreach ((SchematicRowRef row, string text) in typed)
            {
                int index = SchematicRows.IndexOf(read.Select(item => item.Entry).ToList(), row);

                if (index >= 0)
                {
                    rows[index].NameDraft = text;
                }
            }

            // Each row's CAD name choices are in place BEFORE it is listed, so its drop-down never
            // starts out empty (BoardFilesWindow.CadNames.cs).
            this.ApplyCadNameChoices(rows);

            foreach (BoardSchematicRow row in rows)
            {
                this.Schematics.Add(row);
            }

            this.NoSchematicsText.IsVisible = this.Schematics.Count == 0;
        }

        // ###########################################################################################
        // Releases the previews when the window goes away.
        //
        // OnClosed rather than OnClosing: closing can still be cancelled, and disposing a bitmap the
        // window then goes on rendering is an ObjectDisposedException on the render thread, which is
        // fatal in Avalonia. The Workbooks board pane carries the same warning for the same reason.
        // ###########################################################################################
        protected override void OnClosed(EventArgs e)
        {
            this.thisClosed = true;

            // A drag still in flight (released outside the window) stops its auto-scroll timer here.
            this.ResetSchematicRowDrag();

            this.DisposeThumbnails();

            base.OnClosed(e);
        }

        // ###########################################################################################
        // Decodes one row preview, or null when the file cannot be read.
        //
        // *** DECODED AT THUMBNAIL WIDTH, NEVER FULL SIZE. *** A board scan here is routinely
        // 4000px+ across, which is ~47 MB of BGRA once decoded - a dozen of those loaded to draw a
        // 160px preview would cost half a gigabyte for nothing. Bitmap.DecodeToWidth scales during
        // decode, so the full image is never materialised. The same reasoning and the same API the
        // worklog photo rows already use.
        //
        // Every failure returns null rather than throwing: this window is the one place a
        // contributor looks to find out that an import went wrong, so it has to survive a corrupt
        // or half-copied file and SHOW that rather than refusing to open.
        // ###########################################################################################
        // Static, and handed its two roots: it runs on a pool thread (ReloadSchematicsFromDraftAsync),
        // where nothing of the window may be touched.
        private static Bitmap? LoadThumbnail(string dataRoot, string draftFolder, string imageFile)
        {
            if (string.IsNullOrWhiteSpace(imageFile))
            {
                return null;
            }

            try
            {
                // ###########################################################################################
                // *** RESOLVED THROUGH DraftFileResolver, NEVER BY COMBINING THE TWO BY HAND. ***
                //
                // A stored SchematicImageFile is relative to the DATA ROOT
                // ("Test Manu3/Test HW3/Test Board3/6510.jpg"), while the draft folder already IS
                // those same segments - so Path.Combine(draftFolder, imageFile) asks for
                // "<draft>/Test Manu3/Test HW3/Test Board3/Test Manu3/Test HW3/Test Board3/6510.jpg"
                // and every preview came back empty. That is exactly the doubled path
                // DraftFileResolver documents and strips, and writing the combine by hand here
                // walked straight into it.
                //
                // It also gives the right answer for an image that lives only in the published tree
                // (a shared file is not copied into a draft), which a draft-only lookup never could.
                // ###########################################################################################
                string? path = DraftFileResolver.Resolve(
                    dataRoot,
                    draftFolder,
                    imageFile);

                if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
                {
                    return null;
                }

                using FileStream stream = File.OpenRead(path);

                return Bitmap.DecodeToWidth(stream, BoardFilesWindow.ThumbnailWidth);
            }
            catch (Exception ex)
            {
                CrtLog.Warning($"Could not build a preview for [{imageFile}] - [{ex.Message}]");
                return null;
            }
        }

        // ###########################################################################################
        // Releases every decoded preview.
        //
        // Called before each rebuild and on close. The rows are replaced wholesale by
        // ReloadSchematicsFromDraft, so without this each import leaks the previous pass worth of
        // bitmaps - and unlike a cache there is nothing to reuse them.
        // ###########################################################################################
        private void DisposeThumbnails()
        {
            foreach (BoardSchematicRow row in this.Schematics)
            {
                row.Thumbnail?.Dispose();
            }
        }

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
        // ###########################################################################################
        private void RefreshKiCadState()
        {
            var files = this.EnumerateImportedKiCadFiles();

            this.KiCadStateText.Text = files.Count == 0
                ? "No KiCad data imported yet."
                : files.Count == 1
                    ? "1 KiCad file imported:"
                    : $"{files.Count} KiCad files imported:";

            this.KiCadFiles.Clear();

            string root = this.KiCadImportFolder;

            foreach (string file in files)
            {
                this.KiCadFiles.Add(new BoardKiCadFileRow
                {
                    RelativePath = BoardFilesWindow.RelativeToFolder(root, file),
                    SizeText = BoardFilesWindow.DescribeSize(file),
                    FullPath = file,
                });
            }

            this.KiCadFilesPanel.IsVisible = this.KiCadFiles.Count > 0;
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

        private List<string> EnumerateImportedKiCadFiles()
        {
            string folder = this.KiCadImportFolder;

            if (string.IsNullOrWhiteSpace(folder) || !Directory.Exists(folder))
            {
                return new List<string>();
            }

            // Sub-folders included - see KiCadRawFileScanner. A multi-sheet project keeps most of
            // its sheets under Pages/, and counting only the top level reported "3 KiCad files
            // imported" for a 25-file project.
            return KiCadRawFileScanner.Scan(folder);
        }

        // ###########################################################################################
        // The draft's own "KiCad data" folder - at the draft folder's ROOT, exactly where a
        // published board keeps it (Phase 6, 2026-09-23).
        //
        // It used to sit under a Files/ subfolder, which was the draft-only layout. Now that a
        // draft folder IS a board folder the extra level is gone, so publishing a drafted board
        // needs no path rewriting at all and Main.GetCurrentBoardKiCadRawPaths finds it with the
        // identical rule it applies to an official board.
        // ###########################################################################################
        private string KiCadImportFolder =>
            string.IsNullOrWhiteSpace(this.thisDraftFolder)
                ? string.Empty
                : Path.Combine(this.thisDraftFolder, AppConfig.KiCadDataFolderName);

        // The board's published "KiCad data" folder (DraftBoardSource.PublishedKiCadFolderOf) - read
        // with the draft's, as the Schematics tab reads both (KiCadRawFileScanner.ScanBoard).
        private string PublishedKiCadFolder =>
            DraftBoardSource.PublishedKiCadFolderOf(this.DataRoot, this.thisExcelDataFile);

        // The published data tree - DataManager's, or a test's (set before Initialize: a test may not
        // point DataManager anywhere from the HeadlessUi collection).
        internal string? DataRootOverrideForTests { get; set; }

        private string DataRoot => this.DataRootOverrideForTests ?? DataManager.DataRoot;

        // ------------------------------------------------------------------ Schematic import

        private async void OnAddSchematicsClick(object? sender, RoutedEventArgs e)
        {
            try
            {
                var topLevel = GetTopLevel(this);
                if (topLevel == null)
                {
                    return;
                }

                var files = await topLevel.StorageProvider.OpenFilePickerAsync(new FilePickerOpenOptions
                {
                    Title = "Select schematic images",
                    AllowMultiple = true,
                    FileTypeFilter = new[]
                    {
                        new FilePickerFileType("Images")
                        {
                            Patterns = ContributionPackaging.DisplayableImageExtensions
                                .Select(extension => "*" + extension)
                                .ToArray(),
                        },
                    },
                });

                if (files == null || files.Count == 0)
                {
                    return;
                }

                await this.ImportSchematicsAsync(files.Select(file => file.Path.LocalPath));
            }
            catch (Exception ex)
            {
                // Nothing may leave an async void handler: App lets such an exception end CRT
                // (code review, 2026-10-13 - AsyncVoidHandlersGuardTests).
                Logger.Critical($"OnAddSchematicsClick failed - [{ex}]");
            }
        }

        private async void OnSchematicsDropped(object? sender, DragEventArgs e)
        {
            try
            {
                var paths = GetDroppedPaths(e);
                if (paths.Count == 0)
                {
                    return;
                }

                await this.ImportSchematicsAsync(paths.Where(File.Exists));
            }
            catch (Exception ex)
            {
                // Nothing may leave an async void handler: App lets such an exception end CRT
                // (code review, 2026-10-13 - AsyncVoidHandlersGuardTests).
                Logger.Critical($"OnSchematicsDropped failed - [{ex}]");
            }
        }

        // ###########################################################################################
        // Copies each image into the draft and writes an Added Schematics row for it.
        //
        // The schematic NAME defaults to the file name without its extension, which is what a
        // contributor has already named the view when they saved the scan. The stored image path is
        // the bare file name, matching how a published board stores it (the reference C64 board
        // keeps its images directly in the board folder), so publishing later needs no rewriting.
        //
        // *** THE COPYING AND THE DRAFT WRITE RUN OFF THE UI THREAD, under the "please wait"
        // overlay (2026-09-28). *** A handful of large scans and a whole-workbook rewrite used to
        // freeze the window with nothing on screen. What is decided from the window (which names are
        // taken) is read here first; the pool thread sees only plain values. One of the window's
        // draft changes, queued behind any before it (BoardFilesWindow.DraftChanges.cs).
        // ###########################################################################################
        internal Task ImportSchematicsAsync(IEnumerable<string> sourcePaths)
        {
            if (string.IsNullOrWhiteSpace(this.thisDraftFolder))
            {
                this.ShowStatus("Could not resolve where to save - no drafts folder for this board.", isError: true);
                return Task.CompletedTask;
            }

            var accepted = sourcePaths
                .Where(path => !string.IsNullOrWhiteSpace(path))
                .Where(ContributionPackaging.IsDisplayableImageFile)
                .ToList();

            if (accepted.Count == 0)
            {
                this.ShowStatus("No usable image files - pick PNG, JPG, GIF, BMP or WEBP.", isError: true);
                return Task.CompletedTask;
            }

            string draftFolder = this.thisDraftFolder;
            string excelDataFile = this.thisExcelDataFile;

            // The copy AND the list read again are one wait, so the window does not brighten for a
            // moment between them.
            return this.ChangeDraftAsync(() => BusyOverlay.HoldAsync(this, CrtWaitWording.AddingSchematics, async () =>
            {
                // Names taken so far, seeded from what is already drafted and added to as this batch
                // goes. The list on screen is only reloaded from disk once the whole batch is saved, so
                // checking it alone would let two files imported together collide with each other - and
                // a colliding schematic name REPLACES the earlier row (the name is the section's
                // natural key), silently leaving one schematic where two were imported.
                var takenNames = new HashSet<string>(
                    this.Schematics.Select(row => row.SchematicName),
                    StringComparer.OrdinalIgnoreCase);

                (int importedCount, bool saved) = await this.WriteDraftAsync(CrtWaitWording.AddingSchematics, () =>
                    BoardFilesWindow.CopySchematicsIntoDraft(accepted, draftFolder, excelDataFile, takenNames));

                if (importedCount == 0)
                {
                    this.ShowStatus("Could not import those images - see the log for details.", isError: true);
                    return;
                }

                if (!saved)
                {
                    // The copies were taken back out (CopySchematicsIntoDraft) - nothing changed.
                    this.ShowStatus(
                        "The images could not be added to your draft. If this board's Excel file is open in Excel, close it there and try again.",
                        isError: true);
                    return;
                }

                this.ChangedTheDraft = true;

                await this.ReloadSchematicsUnderOverlayAsync();
                this.ShowStatus(importedCount == 1 ? "Added 1 schematic image." : $"Added {importedCount} schematic images.");
            }));
        }

        // ###########################################################################################
        // The copying and the one draft edit, on the pool thread: how many images were copied, and
        // whether the draft took them.
        //
        // *** A DRAFT THAT DID NOT TAKE THEM GETS NO COPIES EITHER (code review, 2026-10-10). *** The
        // edit's answer was ignored, so a workbook Excel held open still read "Added 2 schematic
        // images." with nothing new in the list - and the copies stayed in the draft folder, rows
        // nothing would ever name. They are deleted again: each was copied to a file name nothing had
        // (ResolveFreeImageFileName), so nothing else is lost with them.
        // ###########################################################################################
        private static (int Imported, bool Saved) CopySchematicsIntoDraft(
            IReadOnlyList<string> accepted,
            string draftFolder,
            string excelDataFile,
            HashSet<string> takenNames)
        {
            int importedCount = 0;

            // Collected as the batch goes and written in ONE draft edit at the end - a save per
            // image would re-read and re-write the whole workbook per file.
            var importedSchematics = new List<(string SchematicName, string StoredFileName)>();

            foreach (string sourcePath in accepted)
            {
                try
                {
                    // ###########################################################################################
                    // *** THE IMAGE LANDS AT THE FOLDER ROOT, NOT IN A Files/ SUBFOLDER (Phase 6).
                    // *** A draft folder is a board folder, and a published board keeps its
                    // schematic images beside its workbook. The stored row therefore carries the
                    // ordinary DATA-ROOT-RELATIVE path a published row does
                    // ("Commodore/C64/250407/top.png"), so nothing about the row is draft-aware.
                    // ###########################################################################################
                    string storedFileName = BoardFilesWindow.ResolveFreeImageFileName(draftFolder, Path.GetFileName(sourcePath));
                    string destination = Path.Combine(draftFolder, storedFileName);

                    Directory.CreateDirectory(Path.GetDirectoryName(destination)!);
                    File.Copy(sourcePath, destination, overwrite: false);

                    string schematicName = ResolveFreeSchematicName(
                        Path.GetFileNameWithoutExtension(storedFileName),
                        takenNames);

                    takenNames.Add(schematicName);
                    importedSchematics.Add((schematicName, storedFileName));

                    importedCount++;
                }
                catch (Exception ex)
                {
                    Logger.Warning($"Failed to import board image [{sourcePath}] - [{ex.Message}]");
                }
            }

            if (importedCount == 0)
            {
                return (0, false);
            }

            bool saved = DraftWorkbookStore.Edit(
                DraftManager.DraftsRoot,
                excelDataFile,
                board => BoardFilesWindow.WithSchematics(board, excelDataFile, importedSchematics));

            if (!saved)
            {
                foreach ((_, string storedFileName) in importedSchematics)
                {
                    try
                    {
                        File.Delete(Path.Combine(draftFolder, storedFileName));
                    }
                    catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
                    {
                        Logger.Warning($"Could not take back the copied image [{storedFileName}] - [{ex.Message}]");
                    }
                }
            }

            return (importedCount, saved);
        }

        // ###########################################################################################
        // A file name not already taken inside the draft's Files/ folder. Never overwrites: two
        // different scans legitimately share a name ("top.png" from two folders), and silently
        // replacing the first with the second would lose a file the contributor still has a row
        // pointing at.
        // ###########################################################################################
        private static string ResolveFreeImageFileName(string draftFolder, string fileName)
        {
            string baseName = Path.GetFileNameWithoutExtension(fileName);
            string extension = Path.GetExtension(fileName);
            string candidate = fileName;
            int suffix = 2;

            while (File.Exists(Path.Combine(draftFolder, candidate)))
            {
                candidate = $"{baseName} ({suffix}){extension}";
                suffix++;
            }

            return candidate;
        }

        // ###########################################################################################
        // The board with imported schematics added - the transform DraftWorkbookStore.Edit applies.
        //
        // *** THE STORED PATH IS DATA-ROOT-RELATIVE, exactly as a published row's is. *** The file
        // sits at the draft folder's root, and that folder IS the board's
        // Manufacturer/Hardware/Board path - so the row reads "Commodore/C64/250407/top.png" and
        // resolves correctly whether the board is being read from the draft tree or, once
        // published, from Data/. Storing a bare file name would make the row meaningless anywhere
        // but inside the draft.
        //
        // A name that already exists REPLACES its row rather than adding a second: schematic name
        // is the section's natural key, so two rows sharing one would be a duplicate the differ has
        // to guess about. The caller's ResolveFreeSchematicName already avoids this within a batch;
        // this is the guard for a name that arrived some other way.
        // ###########################################################################################
        private static BoardData WithSchematics(
            BoardData board,
            string excelDataFile,
            IReadOnlyList<(string SchematicName, string StoredFileName)> imported)
        {
            string boardFolder = string.Join(
                '/',
                excelDataFile.Split('/', StringSplitOptions.RemoveEmptyEntries).SkipLast(1));

            var schematics = new List<BoardSchematicEntry>(board.Schematics);

            foreach ((string schematicName, string storedFileName) in imported)
            {
                schematics.RemoveAll(existing => string.Equals(
                    existing.SchematicName?.Trim(),
                    schematicName,
                    StringComparison.OrdinalIgnoreCase));

                schematics.Add(new BoardSchematicEntry
                {
                    SchematicName = schematicName,
                    SchematicImageFile = boardFolder.Length > 0
                        ? boardFolder + "/" + storedFileName
                        : storedFileName,
                });
            }

            return BoardFilesWindow.WithSchematicList(board, schematics);
        }

        // ###########################################################################################
        // Copies the board with its Schematics section replaced. Every OTHER field is carried
        // across - one left out of this copy is erased from the workbook, silently.
        //
        // *** THROUGH BoardData.WithSchematics, NOT A HAND-WRITTEN COPY (2026-09-27). *** The copy
        // that was here listed the sections and forgot HardwareName and BoardName, so every image
        // import or removal wrote the draft back without its "# Hardware:" / "# Board:" caption.
        // ###########################################################################################
        internal static BoardData WithSchematicList(BoardData board, List<BoardSchematicEntry> schematics) =>
            board.WithSchematics(schematics);

        // ###########################################################################################
        // A schematic name not in takenNames. The name is the section's natural key, so reusing one
        // would REPLACE the existing row (see NewBoardDraftWriter.AddSchematic) and the contributor
        // would silently end up with one schematic where they imported two.
        //
        // internal and static so it can be tested directly: the collision rule is the whole point,
        // and the copy path around it needs a real filesystem that rule 6 keeps out of the suite.
        // ###########################################################################################
        internal static string ResolveFreeSchematicName(string desiredName, ISet<string> takenNames)
        {
            string baseName = string.IsNullOrWhiteSpace(desiredName) ? "Board image" : desiredName.Trim();
            string candidate = baseName;
            int suffix = 2;

            while (takenNames.Contains(candidate))
            {
                candidate = $"{baseName} ({suffix})";
                suffix++;
            }

            return candidate;
        }

        // ###########################################################################################
        // Removes a schematic row from the draft. The copied image file is deliberately LEFT in the
        // draft folder: it costs a few megabytes that the eventual discard reclaims anyway
        // (DiscardDraft deletes the whole folder), and deleting it here would destroy the only copy
        // if the contributor removed the row by mistake.
        //
        // *** IT ASKS FIRST WHEN SOMETHING IS LOST WITH IT (code review, 2026-10-10). *** The
        // highlights drawn on the image and its KiCad calibration go with the row (every draft
        // write follows what hangs on a schematic's name), and adding the image again brings none
        // of them back - SchematicRemoval has why. An image with nothing on it goes at once.
        //
        // Only the row pressed on goes - found by its name AND its image file (SchematicRows). By
        // name alone, two rows sharing a name both went.
        // ###########################################################################################
        private async void OnRemoveSchematicClick(object? sender, RoutedEventArgs e)
        {
            try
            {
                if (sender is Button { Tag: BoardSchematicRow row })
                {
                    await this.RemoveSchematicAsync(row);
                }
            }
            catch (Exception ex)
            {
                // Nothing may leave an async void handler: App lets such an exception end CRT
                // (code review, 2026-10-13 - AsyncVoidHandlersGuardTests).
                Logger.Critical($"OnRemoveSchematicClick failed - [{ex}]");
            }
        }

        internal Task RemoveSchematicAsync(BoardSchematicRow row) =>
            this.ChangeDraftAsync(async () =>
            {
                if (!this.Schematics.Contains(row))
                {
                    return;
                }

                string name = row.SchematicName;
                string draftsRoot = DraftManager.DraftsRoot;
                string excelDataFile = this.thisExcelDataFile;
                string workbook = DraftFolderLayout.GetWorkbookPath(draftsRoot, excelDataFile);
                SchematicRowRef target = this.RefOf(row);

                // Another image keeping the name keeps what hangs on it, too.
                bool nameStays = this.Schematics.Any(other => !ReferenceEquals(other, row) && SchematicRename.SameName(other.SchematicName, name));

                SchematicRemovalLoss loss = nameStays
                    ? SchematicRemovalLoss.Nothing
                    : await this.WriteDraftAsync(CrtWaitWording.CheckingSchematicRemoval, () => SchematicRemoval.Read(workbook, name));

                if (!loss.IsNothing && !await this.ConfirmRemovalAsync(name, loss))
                {
                    return;
                }

                await BusyOverlay.HoldAsync(this, CrtWaitWording.RemovingSchematic, async () =>
                {
                    bool found = false;

                    // Removing the ROW, not the file - the image bytes stay in the draft folder, which
                    // is what the status message below promises. Dropping a row from a board is simply
                    // absence now; there is no tombstone to write.
                    bool saved = await this.WriteDraftAsync(CrtWaitWording.RemovingSchematic, () => DraftWorkbookStore.Edit(
                        draftsRoot,
                        excelDataFile,
                        board =>
                        {
                            (BoardData without, bool wasFound) = SchematicRows.Without(board, target);
                            found = wasFound;
                            return without;
                        }));

                    if (saved && found)
                    {
                        this.ChangedTheDraft = true;
                    }

                    await this.ReloadSchematicsUnderOverlayAsync();

                    if (!saved)
                    {
                        this.ShowStatus(
                            $"[{name}] could not be removed. If this board's Excel file is open in Excel, close it there and try again.",
                            isError: true);
                    }
                    else if (!found)
                    {
                        this.ShowStatus(
                            $"[{name}] is no longer in the draft - it changed outside this window, so the list has been read again.",
                            isError: true);
                    }
                    else
                    {
                        this.ShowStatus($"Removed [{name}]. The image file is kept in your draft folder.");
                    }
                });
            });

        // The question RemoveSchematicWindow asks - or the test's stand-in for it.
        private async Task<bool> ConfirmRemovalAsync(string schematicName, SchematicRemovalLoss loss)
        {
            if (this.ConfirmRemovalOverrideForTests is { } confirm)
            {
                return await confirm(schematicName, loss);
            }

            var dialog = new RemoveSchematicWindow();
            dialog.Initialize(schematicName, loss);

            return await dialog.ShowDialog<bool?>(this) == true;
        }

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

                if (!KiCadImportedFiles.TryRemove(this.KiCadImportFolder, row.FullPath))
                {
                    this.ShowStatus($"Could not remove [{row.RelativePath}] - see the log for details.", isError: true);
                    return;
                }

                this.ChangedTheDraft = true;

                this.RefreshKiCadState();
                this.ShowStatus($"Removed [{row.RelativePath}] from your draft. Your own KiCad project is not touched - import it again to get the file back.");

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
        internal async Task ImportKiCadFolderAsync(string sourceFolder)
        {
            if (string.IsNullOrWhiteSpace(this.thisDraftFolder))
            {
                this.ShowStatus("Could not resolve where to save - no drafts folder for this board.", isError: true);
                return;
            }

            List<string> sourceFiles;

            try
            {
                sourceFiles = KiCadRawFileScanner.Scan(sourceFolder);
            }
            catch (Exception ex)
            {
                Logger.Warning($"Could not read the KiCad folder [{sourceFolder}] - [{ex.Message}]");
                this.ShowStatus("Could not read that folder.", isError: true);
                return;
            }

            if (sourceFiles.Count == 0)
            {
                this.ShowStatus(
                    "No KiCad files in that folder. Only modern .kicad_pcb / .kicad_pro / .kicad_sch files are imported - sub-folders are searched too.",
                    isError: true);
                return;
            }

            string destinationFolder = this.KiCadImportFolder;

            // Re-importing OVERWRITES files a report may still be reading - the same sharing
            // violation a Remove would hit, so the same wait (see thisPendingReport).
            await this.WaitForPendingReportAsync();

            // The copy AND the report that follows are one wait (2026-09-28): the window is held
            // across both, so it does not brighten for a moment between them. One of the window's
            // draft changes, so closing waits for it (BoardFilesWindow.DraftChanges.cs).
            await this.ChangeDraftAsync(() => BusyOverlay.HoldAsync(this, CrtWaitWording.ImportingKiCad, async () =>
            {
                (int copied, bool complete) = await this.WriteDraftAsync(
                    CrtWaitWording.ImportingKiCad,
                    () => KiCadImportedFiles.CopyIn(sourceFolder, sourceFiles, destinationFolder));

                if (copied == 0)
                {
                    this.ShowStatus("Could not copy those KiCad files - see the log for details.", isError: true);
                    return;
                }

                // ###########################################################################################
                // *** EVEN A COPY THAT STOPPED PART-WAY CHANGED THE DRAFT (code review, 2026-10-10). ***
                // The files before the one that failed have been overwritten, so the board on screen
                // is read again when the window closes - it used to go on drawing the old project.
                // ###########################################################################################
                this.ChangedTheDraft = true;

                this.RefreshKiCadState();

                if (complete)
                {
                    this.ShowStatus(sourceFiles.Count == 1
                        ? "Imported 1 KiCad file. Checking what it lines up with..."
                        : $"Imported {sourceFiles.Count} KiCad files. Checking what they line up with...");
                }
                else
                {
                    this.ShowStatus(
                        $"Only {copied} of the {sourceFiles.Count} KiCad files could be copied - see the log for details. Import the folder again once the rest can be read.",
                        isError: true);
                }

                await BusyOverlay.RunLocalAsync(this, CrtWaitWording.CheckingKiCadMatches, this.RefreshReportAsync);
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
        // published board's "KiCad data" AND the draft's (KiCadRawFileScanner.ScanBoard). *** Read
        // from the draft's alone, a net or a footprint only the published files have was reported
        // as missing - "will NOT light up" for an important signal the Schematics tab draws - and
        // the CAD names on offer differed from the "Schematic images" window's.
        // ###########################################################################################
        private async Task BuildAndShowReportAsync()
        {
            List<string> publishedFiles = KiCadRawFileScanner.Scan(this.PublishedKiCadFolder);
            List<string> draftFiles = this.EnumerateImportedKiCadFiles();
            List<string> files = KiCadRawFileScanner.ScanBoard(this.PublishedKiCadFolder, this.KiCadImportFolder);

            string source = KiCadImportedFiles.ReportSourceLine(publishedFiles.Count, draftFiles.Count);
            this.ReportSourceText.Text = source;
            this.ReportSourceText.IsVisible = source.Length > 0;

            if (files.Count == 0)
            {
                this.ReportPanel.IsVisible = false;
                return;
            }

            var bundle = await KiCadProjectLoader.LoadRawAsync(files);
            var root = bundle?.Root;

            if (root == null)
            {
                this.ShowStatus("Those KiCad files could not be read.", isError: true);
                this.ReportPanel.IsVisible = false;
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

        // ------------------------------------------------------------------ Shared plumbing

        // Only a copy is ever offered, never a move - the source files stay where the contributor
        // put them, the same import-by-copy rule every other attachment path in the app follows.
        // Without the format check the box appears to accept dragged text and then silently does
        // nothing, the same trap WorklogAddPhotoWindow's own drop target documents.
        private static void OnDragOver(object? sender, DragEventArgs e)
        {
            e.DragEffects = e.DataTransfer.Contains(DataFormat.File) ? DragDropEffects.Copy : DragDropEffects.None;
            e.Handled = true;
        }

        private static List<string> GetDroppedPaths(DragEventArgs e)
        {
            e.Handled = true;

            var items = e.DataTransfer.TryGetFiles();

            return items == null
                ? new List<string>()
                : items
                    .Select(item => item.TryGetLocalPath())
                    .Where(path => !string.IsNullOrWhiteSpace(path))
                    .Select(path => path!)
                    .ToList();
        }

        internal void ShowStatus(string message, bool isError = false)
        {
            this.StatusText.Text = message;
            this.StatusText.Foreground = ThemeResources.ResolveBrush(isError ? "Text_Fail_Fg" : "Text_Success_Fg");
        }

        internal string StatusTextForTests => this.StatusText.Text ?? string.Empty;

        private void OnWindowKeyDown(object? sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
            {
                // In a Name box holding an unsaved name, Escape puts the name back - it does not
                // close the window (BoardFilesWindow.SchematicNames.cs).
                if (this.RevertNameEdit(e.Source))
                {
                    e.Handled = true;
                    return;
                }

                // ###########################################################################################
                // *** AN OPEN DROP-DOWN IS CLOSED BY ITS OWN ESCAPE (code review, 2026-10-10). *** This
                // handler runs on the TUNNEL route, before the ComboBox sees the key - so Escape pressed
                // to put away an image's CAD name list closed the whole window instead. Left alone, the
                // ComboBox closes its list and keeps its pick.
                // ###########################################################################################
                if (this.SchematicsItemsControl.GetVisualDescendants().OfType<ComboBox>().Any(combo => combo.IsDropDownOpen))
                {
                    return;
                }

                this.Close();
                e.Handled = true;
            }
        }

        private void OnCloseClick(object? sender, RoutedEventArgs e) => this.Close();
    }

    // ###########################################################################################
    // Which half of BoardFilesWindow to open - one per button on the Drafts tab.
    // ###########################################################################################
    public enum BoardFilesSection
    {
        SchematicImages,
        KiCadData,
    }

    // ###########################################################################################
    // One imported KiCad file, as the window lists it: where it sits inside the "KiCad data" folder
    // and how big it is.
    // ###########################################################################################
    public sealed class BoardKiCadFileRow
    {
        public string RelativePath { get; init; } = string.Empty;

        public string SizeText { get; init; } = string.Empty;

        // The file on disk, exactly as the scan found it - what "Remove" deletes. Kept separately
        // rather than rebuilt from RelativePath, which is display text (forward-slashed).
        public string FullPath { get; init; } = string.Empty;
    }

    // ###########################################################################################
    // One board image row in the list - a view model over a drafted Schematics row.
    // ###########################################################################################
    public sealed class BoardSchematicRow : System.ComponentModel.INotifyPropertyChanged, IDraggableRow
    {
        public event System.ComponentModel.PropertyChangedEventHandler? PropertyChanged;

        // ###########################################################################################
        // True while this row is the one being dragged, which draws it as an empty dashed slot
        // where a drop would land - the worklog photo rows' IsDropPlaceholder, for the same reason:
        // the dragged row moves through the list and renders as the gap itself, so the rows around
        // it already stand in the order the drop will produce. See BoardFilesWindow.SchematicOrder.cs.
        // ###########################################################################################
        public bool IsDropPlaceholder
        {
            get => this.thisIsDropPlaceholder;
            set
            {
                if (this.thisIsDropPlaceholder == value)
                {
                    return;
                }

                this.thisIsDropPlaceholder = value;
                this.PropertyChanged?.Invoke(this, new System.ComponentModel.PropertyChangedEventArgs(nameof(this.IsDropPlaceholder)));
                this.PropertyChanged?.Invoke(this, new System.ComponentModel.PropertyChangedEventArgs(nameof(this.IsNotDropPlaceholder)));
            }
        }

        public bool IsNotDropPlaceholder => !this.thisIsDropPlaceholder;

        private bool thisIsDropPlaceholder;

        // ###########################################################################################
        // The row's own measured height while it is the placeholder, so the gap is exactly the row
        // being moved. It MUST notify: it is assigned just before IsDropPlaceholder, and a binding
        // that never heard the change drew every gap at the starting value (the worklog's Files
        // list found this out). Stored always; only the notification is gated, against sub-pixel
        // churn mid-drag.
        // ###########################################################################################
        public double PlaceholderHeight
        {
            get => this.thisPlaceholderHeight;
            set
            {
                bool isMeaningfulChange = Math.Abs(this.thisPlaceholderHeight - value) >= 0.5;

                this.thisPlaceholderHeight = value;

                if (isMeaningfulChange)
                {
                    this.PropertyChanged?.Invoke(this, new System.ComponentModel.PropertyChangedEventArgs(nameof(this.PlaceholderHeight)));
                }
            }
        }

        private double thisPlaceholderHeight = 72.0;

        // ###########################################################################################
        // The image's name as the draft holds it, and what the Name box shows (owner request,
        // 2026-10-10). The box edits NameDraft; only a rename the draft took sets SchematicName
        // (BoardFilesWindow.SchematicNames.cs) - every save in this window finds its row by it.
        // ###########################################################################################
        public string SchematicName
        {
            get => this.thisSchematicName;
            set
            {
                if (string.Equals(this.thisSchematicName, value, StringComparison.Ordinal))
                {
                    return;
                }

                this.thisSchematicName = value ?? string.Empty;
                this.RaiseChanged(nameof(this.SchematicName));
            }
        }

        private string thisSchematicName = string.Empty;

        public string NameDraft
        {
            get => this.thisNameDraft ?? this.thisSchematicName;
            set
            {
                if (string.Equals(this.NameDraft, value, StringComparison.Ordinal))
                {
                    return;
                }

                this.thisNameDraft = value;
                this.RaiseChanged(nameof(this.NameDraft));
            }
        }

        private string? thisNameDraft;

        // Whether the Name box holds something the draft does not.
        internal bool HasNameEdit => !string.Equals(this.NameDraft, this.thisSchematicName, StringComparison.Ordinal);

        public string ImageFile { get; init; } = string.Empty;

        // What the draft holds in this image's "CAD name" - set again once a pick is saved.
        public string CadName { get; set; } = string.Empty;

        // ###########################################################################################
        // A small preview of the image this row names (owner request, 2026-09-24), so a
        // contributor can SEE what they imported rather than trusting a file name.
        //
        // Null when the file is missing or unreadable, which the template shows as a placeholder -
        // a draft can legitimately name an image whose file has been moved or deleted by hand, and
        // that is worth seeing rather than hiding.
        //
        // OWNED BY THE WINDOW, not by this row: BoardFilesWindow disposes every thumbnail when it
        // rebuilds the list and when it closes. See DisposeThumbnails.
        // ###########################################################################################
        public Bitmap? Thumbnail { get; init; }

        public bool HasThumbnail => this.Thumbnail != null;

        public bool IsMissing => this.Thumbnail == null;

        // The image's file. The CAD name used to follow it here; it has its own drop-down now.
        public string DetailText => this.ImageFile;

        // ###########################################################################################
        // The CAD name drop-down (owner request, 2026-10-10) - see BoardFilesWindow.CadNames.cs.
        //
        // ONE list for the row's lifetime, changed in place by ApplyCadNameChoices: a ComboBox
        // whose items are replaced lets go of its selection, and this row would hear "nothing
        // chosen" for every image.
        // ###########################################################################################
        public ObservableCollection<CadNameChoiceItem> CadNameChoices { get; } = new();

        // Told when the user picks another CAD name - the window saves it. Never for the drop-down
        // letting go of its items (null).
        internal Action<BoardSchematicRow, string>? CadNameChosen { get; init; }

        public CadNameChoiceItem? SelectedCadNameChoice
        {
            get => this.thisSelectedCadNameChoice;
            set
            {
                if (value is null || ReferenceEquals(value, this.thisSelectedCadNameChoice))
                {
                    return;
                }

                this.thisSelectedCadNameChoice = value;
                this.RaiseChanged(nameof(this.SelectedCadNameChoice));

                if (!string.Equals(value.CadName, this.CadName?.Trim() ?? string.Empty, StringComparison.OrdinalIgnoreCase))
                {
                    this.CadNameChosen?.Invoke(this, value.CadName);
                }
            }
        }

        private CadNameChoiceItem? thisSelectedCadNameChoice;

        // ###########################################################################################
        // Makes CadNameChoices read the given choices, in order, changing the list IN PLACE
        // (InPlaceList.Sync): an entry already there keeps its object, a new one is inserted, one no
        // longer wanted (a name another image has just taken) is removed - never the selected one,
        // which the choices always hold (SchematicCadNameChoices lists the image's own name). Then
        // selects the entry SchematicCadNameChoices.Selected names for CadName.
        // ###########################################################################################
        internal void ApplyCadNameChoices(IReadOnlyList<SchematicCadNameChoice> choices)
        {
            InPlaceList.Sync(
                this.CadNameChoices,
                choices,
                (item, choice) => string.Equals(item.CadName, choice.CadName, StringComparison.OrdinalIgnoreCase),
                choice => new CadNameChoiceItem(choice),
                keep: this.thisSelectedCadNameChoice);

            this.thisSelectedCadNameChoice = choices.Count == 0
                ? null
                : this.CadNameChoices.FirstOrDefault(item => string.Equals(
                    item.CadName,
                    SchematicCadNameChoices.Selected(choices, this.CadName).CadName,
                    StringComparison.OrdinalIgnoreCase));

            // Always said, so a drop-down that lost its selection while its items moved shows it again.
            this.RaiseChanged(nameof(this.SelectedCadNameChoice));
        }

        private void RaiseChanged(string propertyName) =>
            this.PropertyChanged?.Invoke(this, new System.ComponentModel.PropertyChangedEventArgs(propertyName));
    }

    // ###########################################################################################
    // One entry in a schematic image's CAD name drop-down. An object of its own rather than the
    // CRT.Data record, so the entry a drop-down shows keeps its identity while the entries around
    // it come and go (BoardSchematicRow.ApplyCadNameChoices).
    // ###########################################################################################
    public sealed class CadNameChoiceItem
    {
        public CadNameChoiceItem(SchematicCadNameChoice choice)
        {
            this.CadName = choice.CadName;
            this.Label = choice.Label;
        }

        public string CadName { get; }

        public string Label { get; }

        public override string ToString() => this.Label;
    }

    // ###########################################################################################
    // One KiCad net name in the report: copied by a click (BoardFilesWindow.NetNames.cs), and
    // marked when an "Important signals" row already uses it.
    // ###########################################################################################
    public sealed class KiCadNetNameBadge
    {
        public string Name { get; init; } = string.Empty;

        public bool IsUsed { get; init; }

        public string ToolTipText => this.IsUsed
            ? "Click to copy - already used in Important signals"
            : "Click to copy";
    }
}
