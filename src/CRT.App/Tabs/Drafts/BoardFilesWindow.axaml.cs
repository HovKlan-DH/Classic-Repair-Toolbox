using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Handlers.DataHandling;
using Handlers.Theming;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Threading.Tasks;

namespace CRT
{
    // ###########################################################################################
    // "KiCad DATA": importing a board's KiCad project, and reporting what it holds and what it lines
    // up with (NewContributeStrategy.md Phase 2, session 2c, task 9).
    //
    // This replaces manual steps from Assets/Wiki/Add-new-board-with-KiCad-data.md: creating a folder
    // named exactly "KiCad data" and filling it, and reading the generated CAD view names out of the
    // LOGFILE. The report is the piece with no equivalent at all before - a board label that does not
    // match a KiCad reference designator silently lights up no copper, which the Wiki itself calls
    // "the number one cause of 'I did everything and no traces appear'".
    //
    // Everything is copied, never moved (the same import-by-copy rule every other attachment path
    // follows), into the board's own draft folder.
    //
    // *** THE WINDOW'S OTHER HALF, "SCHEMATIC IMAGES", IS GONE (CRT 4.0 step 3 - the plan: "Pictures:
    // today's Schematic images content, moved inline"). *** Pictures are added, named, matched to
    // their KiCad view, ordered and removed on their cards in the Contribute tab's Pictures part
    // (ComponentPane.BoardPictureList.cs), its rules - SchematicRename, SchematicCadNameChoices,
    // SchematicRemoval - unchanged.
    //
    // FILE MAP:
    //   BoardFilesWindow.axaml.cs          - opening, files dropped on the Contribute tab, closing
    //   BoardFilesWindow.KiCad.cs          - the KiCad files, their import and removal, the report and
    //                                        the import's summary
    //   BoardFilesWindow.DraftChanges.cs   - every change to the draft: queued, off the UI thread,
    //                                        under the "please wait"; and closing after them
    //   BoardFilesWindow.NetNames.cs       - the report's KiCad net names and Important signals
    // ###########################################################################################
    public partial class BoardFilesWindow : Window
    {
        private string thisExcelDataFile = string.Empty;
        private string thisDraftFolder = string.Empty;

        // The board labels the KiCad report compares against. Supplied by the caller from the
        // MERGED BoardData (official plus draft), so this works the same on a brand-new board
        // (where every label is drafted) and on an existing board.
        private IReadOnlyList<string> thisBoardLabels = Array.Empty<string>();

        // The board's "Important signals" rows, from the same merged BoardData as the labels - the
        // report says which of them name a net the KiCad data does not have (2026-10-10).
        private IReadOnlyList<KiCadImportantSignalEntry> thisImportantSignals = Array.Empty<KiCadImportantSignalEntry>();

        // The imported KiCad files, listed so a returning contributor can see WHICH project is in
        // this board rather than only how many files it has.
        public ObservableCollection<BoardKiCadFileRow> KiCadFiles { get; } = new();

        public BoardFilesWindow()
        {
            this.InitializeComponent();

            this.KiCadFilesItemsControl.ItemsSource = this.KiCadFiles;

            // Drag and drop as well as the picker: dragging the project folder in is the natural
            // gesture, and the Wiki's manual procedure was itself a series of file copies.
            DragDrop.SetAllowDrop(this.KiCadDropTarget, true);

            this.KiCadDropTarget.AddHandler(DragDrop.DragOverEvent, OnDragOver);
            this.KiCadDropTarget.AddHandler(DragDrop.DropEvent, this.OnKiCadDropped);

            this.AddHandler(KeyDownEvent, this.OnWindowKeyDown, RoutingStrategies.Tunnel);
        }

        // ###########################################################################################
        // Points the window at one board. boardLabels comes from the merged BoardData the caller
        // already has, rather than being re-derived here - the caller is the one place that knows
        // whether it is looking at a drafted-only board or an official one with an overlay.
        //
        // *** THE REPORT IS BUILT FOR DATA THAT IS ALREADY THERE (owner report, 2026-09-24). *** It used
        // to be built only at the end of an import, so reopening the window on a board that already had
        // KiCad data showed an empty panel. Parsing a 45 MB .kicad_pcb takes a moment, so it is read
        // under the "please wait" (owner request, 2026-10-10) - and so is the list of files, which the
        // same read brings (ReadKiCadAsync). Opening completes when all of it is in.
        // ###########################################################################################
        public void Initialize(
            string displayName,
            string excelDataFile,
            IReadOnlyList<string> boardLabels,
            IReadOnlyList<KiCadImportantSignalEntry>? importantSignals = null)
        {
            this.thisExcelDataFile = excelDataFile;
            this.thisDraftFolder = DraftManager.GetBoardFolder(excelDataFile);
            this.thisBoardLabels = boardLabels;
            this.thisImportantSignals = importantSignals ?? Array.Empty<KiCadImportantSignalEntry>();

            this.HeaderText.Text = displayName;

            this.Opening = BusyOverlay.HoldAsync(this, CrtWaitWording.CheckingKiCadMatches, this.RefreshReportAsync);
        }

        // Completes when everything the window reads on opening is in - for tests, which wait for it
        // as the user does.
        internal Task Opening { get; private set; } = Task.CompletedTask;

        // ###########################################################################################
        // A KiCad project folder dropped on the Contribute tab (CRT 4.0 step 3, slice 3e -
        // TabDrafts.Drop.cs): imported once the window has read what it opens on, exactly as its own
        // drop target imports it - its summary and all. DroppedImport completes when done.
        // ###########################################################################################
        internal void ImportOnceOpen(IReadOnlyList<string> paths) => this.DroppedImport = this.ImportOnceOpenAsync(paths);

        internal Task DroppedImport { get; private set; } = Task.CompletedTask;

        private async Task ImportOnceOpenAsync(IReadOnlyList<string> paths)
        {
            try
            {
                await this.Opening;

                if (!this.thisClosed && paths.Count > 0)
                {
                    await this.ImportKiCadFolderAsync(paths[0]);
                }
            }
            catch (Exception ex)
            {
                // Started without an awaiting caller: nothing may escape it.
                Logger.Critical($"Importing the KiCad folder dropped on the Contribute tab failed - [{ex}]");
            }
        }

        // ###########################################################################################
        // Whether anything in this window changed the draft - KiCad files imported or removed. The
        // Contribute tab reads the board on screen again after the window closes ONLY then (owner
        // report, 2026-10-10: it was read again every time, changed or not). Set once the change is in.
        // ###########################################################################################
        internal bool ChangedTheDraft { get; private set; }

        // Set by OnClosed: a read still in flight then has nowhere to show what it read.
        private bool thisClosed;

        protected override void OnClosed(EventArgs e)
        {
            this.thisClosed = true;
            base.OnClosed(e);
        }

        private string DraftBoardFile =>
            string.IsNullOrWhiteSpace(this.thisDraftFolder)
                ? string.Empty
                : Path.Combine(this.thisDraftFolder, DataTreeLayout.BoardFileName);

        // The published data tree - DataManager's, or a test's (set before Initialize: a test may not
        // point DataManager anywhere from the HeadlessUi collection).
        internal string? DataRootOverrideForTests { get; set; }

        private string DataRoot => this.DataRootOverrideForTests ?? DataManager.DataRoot;

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

        // Escape closes the window.
        private void OnWindowKeyDown(object? sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
            {
                this.Close();
                e.Handled = true;
            }
        }

        private void OnCloseClick(object? sender, RoutedEventArgs e) => this.Close();
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
