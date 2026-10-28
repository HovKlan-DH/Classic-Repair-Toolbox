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
    //   BoardFilesWindow.axaml.cs          - opening, files dropped on the Contribute tab, the images
    //   BoardFilesWindow.KiCad.cs          - the KiCad files, their import and removal, the report and
    //                                        the import's summary
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

            // ###########################################################################################
            // *** AND REBUILD THE REPORT FOR DATA THAT IS ALREADY THERE (owner report,
            // 2026-09-24). ***
            //
            // The match report used to be built only at the end of an import, so reopening the
            // window on a board that already had KiCad data showed an empty panel - the one piece
            // of information that says whether the board labels line up with the KiCad references
            // was available exactly once, and only to whoever performed the import. Parsing a
            // 45 MB .kicad_pcb takes a moment, so it is under the "please wait" too - and so is the
            // list of files, which the same read brings (ReadKiCadAsync).
            // ###########################################################################################
            this.Opening = BusyOverlay.HoldAsync(this, CrtWaitWording.CheckingKiCadMatches, this.RefreshReportAsync);
        }

        // Completes when everything the window reads on opening is in - for tests, which wait for it
        // as the user does.
        internal Task Opening { get; private set; } = Task.CompletedTask;

        // ###########################################################################################
        // Files dropped on the Contribute tab (CRT 4.0 step 3, slice 3e - TabDrafts.Drop.cs): added
        // once the window has read what it opens on, exactly as its own drop target adds them - the
        // pictures, or the KiCad project folder with its summary. DroppedImport completes when done.
        // ###########################################################################################
        internal void ImportOnceOpen(IReadOnlyList<string> paths) => this.DroppedImport = this.ImportOnceOpenAsync(paths);

        internal Task DroppedImport { get; private set; } = Task.CompletedTask;

        private async Task ImportOnceOpenAsync(IReadOnlyList<string> paths)
        {
            try
            {
                await this.Opening;

                if (this.thisClosed || paths.Count == 0)
                {
                    return;
                }

                if (this.KiCadSection.IsVisible)
                {
                    await this.ImportKiCadFolderAsync(paths[0]);
                }
                else
                {
                    await this.ImportSchematicsAsync(paths.Where(File.Exists));
                }
            }
            catch (Exception ex)
            {
                // Started without an awaiting caller: nothing may escape it.
                Logger.Critical($"Adding the files dropped on the Contribute tab failed - [{ex}]");
            }
        }

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

        private string DraftBoardFile =>
            string.IsNullOrWhiteSpace(this.thisDraftFolder)
                ? string.Empty
                : Path.Combine(this.thisDraftFolder, DataTreeLayout.BoardFileName);

        // The board's KiCad project as the draft has it - the published files and the draft's own -
        // copied under their names for the loaders (BoardKiCadFolder): what the Schematics tab
        // draws the traces from. "" when the board has no KiCad files. Copies files: NEVER call it
        // on the UI thread (the CAD names read it on the pool; the KiCad section through ReadKiCadAsync).
        private string BoardKiCadCopy() =>
            BoardKiCadFolder.ForBoard(this.DataRoot, this.thisDraftFolder, this.thisExcelDataFile, AppConfig.KiCadCacheFolder);

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

            // Never a hidden file - macOS's "._IMG_0001.JPG" beside each picture on a USB stick (code
            // review, 2026-10-28): its name is refused at Submit.
            var accepted = sourcePaths
                .Where(path => !string.IsNullOrWhiteSpace(path))
                .Where(ContributionPackaging.IsAddablePictureFile)
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

                List<(string SchematicName, string StoredFileName)> imported = await this.WriteDraftAsync(CrtWaitWording.AddingSchematics, () =>
                    BoardFilesWindow.CopySchematicsIntoDraftFolder(accepted, draftFolder, takenNames));

                int importedCount = imported.Count;

                if (importedCount == 0)
                {
                    this.ShowStatus("Could not import those images - see the log for details.", isError: true);
                    return;
                }

                // ONE change of the draft for the whole batch - one undo step, one save.
                DraftWorkbookEditOutcome outcome = await this.ApplyToDraftAsync(
                    CrtWaitWording.AddingSchematics,
                    DraftStepWording.SchematicsAdded(importedCount),
                    board => BoardFilesWindow.WithSchematics(board, excelDataFile, imported));

                if (outcome != DraftWorkbookEditOutcome.Saved)
                {
                    // The change was taken back out (DraftEdits), and so are the copies - nothing changed.
                    await BusyOverlay.RunLocalAsync(this, CrtWaitWording.AddingSchematics, () => Task.Run(() => BoardFilesWindow.TakeBackCopies(draftFolder, imported)));

                    this.ShowStatus("The images could not be added to your draft. " + DraftSaveWording.EditNotSaved(outcome), isError: true);
                    return;
                }

                this.ChangedTheDraft = true;

                await this.ReloadSchematicsUnderOverlayAsync();
                this.ShowStatus(importedCount == 1 ? "Added 1 schematic image." : $"Added {importedCount} schematic images.");
            }));
        }

        // ###########################################################################################
        // The copying, on the pool thread: each image copied into the draft folder, with the name its
        // row gets - the batch then goes into the draft as ONE change (ImportSchematicsAsync).
        //
        // *** A DRAFT THAT DID NOT TAKE THEM GETS NO COPIES EITHER (code review, 2026-10-10). *** The
        // edit's answer was ignored, so a workbook Excel held open still read "Added 2 schematic
        // images." with nothing new in the list - and the copies stayed in the draft folder, rows
        // nothing would ever name. They are deleted again (TakeBackCopies): each was copied to a file
        // name nothing had (ResolveFreeImageFileName), so nothing else is lost with them.
        // ###########################################################################################
        private static List<(string SchematicName, string StoredFileName)> CopySchematicsIntoDraftFolder(
            IReadOnlyList<string> accepted,
            string draftFolder,
            HashSet<string> takenNames)
        {
            // Collected as the batch goes and written in ONE draft edit at the end - a save per
            // image would re-write the whole board file per file.
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
                }
                catch (Exception ex)
                {
                    Logger.Warning($"Failed to import board image [{sourcePath}] - [{ex.Message}]");
                }
            }

            return importedSchematics;
        }

        // The copies of a batch the draft did not take, deleted again - see CopySchematicsIntoDraftFolder.
        private static void TakeBackCopies(string draftFolder, IReadOnlyList<(string SchematicName, string StoredFileName)> imported)
        {
            foreach ((_, string storedFileName) in imported)
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
        // The board with imported schematics added - the change ImportSchematicsAsync applies.
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
                    DraftWorkbookEditOutcome outcome = await this.ApplyToDraftAsync(
                        CrtWaitWording.RemovingSchematic,
                        DraftStepWording.SchematicRemoved(name),
                        board =>
                        {
                            (BoardData without, bool wasFound) = SchematicRows.Without(board, target);
                            found = wasFound;
                            return wasFound ? without : null;
                        });

                    bool saved = outcome == DraftWorkbookEditOutcome.Saved;

                    if (saved && found)
                    {
                        this.ChangedTheDraft = true;
                    }

                    await this.ReloadSchematicsUnderOverlayAsync();

                    if (!saved)
                    {
                        this.ShowStatus($"[{name}] could not be removed. " + DraftSaveWording.EditNotSaved(outcome), isError: true);
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

            return StorageItemPaths.LocalPaths(e.DataTransfer.TryGetFiles());
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
