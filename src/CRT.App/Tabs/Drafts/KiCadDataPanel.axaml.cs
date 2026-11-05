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
    // *** ITS OTHER HALF, "SCHEMATIC IMAGES", IS GONE (CRT 4.0 step 3 - the plan: "Pictures: today's
    // Schematic images content, moved inline"). *** Pictures are added, named, matched to their KiCad
    // view, ordered and removed on their cards in the Contribute tab's Pictures part
    // (ComponentPane.BoardPictureList.cs), its rules - SchematicRename, SchematicCadNameChoices,
    // SchematicRemoval - unchanged.
    //
    // *** A PANEL ON THE RIGHT OF THE CONTRIBUTE TAB, NOT A WINDOW (owner request, 2026-10-30: "I do
    // not understand why the 'KiCad data' opens in a new window? ... it should just show up in the
    // 'normal' content area"). *** It was the BoardFilesWindow, a dialog. The tab makes one for each
    // opening (TabDrafts.KiCad.cs) and shows it where a component or a board part shows; its waits run
    // under CRT's own "please wait" (BusyOverlay.For finds Main's). Each change it makes is caught up
    // with as it ends (DraftChanged) - the board on screen, the draft's row, the left column's count -
    // not once it closes, since it can stay open while the contributor looks elsewhere.
    //
    // FILE MAP:
    //   KiCadDataPanel.axaml.cs          - opening, an import asked for as it opens, closing
    //   KiCadDataPanel.KiCad.cs          - the KiCad files, their import and removal, the report and
    //                                        the import's summary
    //   KiCadDataPanel.DraftChanges.cs   - every change to the draft: queued, off the UI thread,
    //                                        under the "please wait"; told as each ends; closing after them
    //   KiCadDataPanel.NetNames.cs       - the report's KiCad net names and Important signals
    // ###########################################################################################
    public partial class KiCadDataPanel : UserControl
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

        public KiCadDataPanel()
        {
            this.InitializeComponent();

            this.KiCadFilesItemsControl.ItemsSource = this.KiCadFiles;

            // Drag and drop as well as the picker: dragging the project folder in is the natural
            // gesture, and the Wiki's manual procedure was itself a series of file copies.
            DragDrop.SetAllowDrop(this.KiCadDropTarget, true);

            this.KiCadDropTarget.AddHandler(DragDrop.DragOverEvent, OnDragOver);
            this.KiCadDropTarget.AddHandler(DragDrop.DropEvent, this.OnKiCadDropped);
        }

        // ###########################################################################################
        // Points the panel at one board. boardLabels comes from the merged BoardData the caller
        // already has, rather than being re-derived here - the caller is the one place that knows
        // whether it is looking at a drafted-only board or an official one with an overlay.
        //
        // *** THE REPORT IS BUILT FOR DATA THAT IS ALREADY THERE (owner report, 2026-09-24). *** It used
        // to be built only at the end of an import, so opening it again on a board that already had
        // KiCad data showed an empty panel. Parsing a 45 MB .kicad_pcb takes a moment, so it is read
        // under the "please wait" (owner request, 2026-10-10) - and so is the list of files, which the
        // same read brings (ReadKiCadAsync). Opening completes when all of it is in. The board is
        // named by the draft's header above the panel, not again here.
        // ###########################################################################################
        public void Initialize(
            string excelDataFile,
            IReadOnlyList<string> boardLabels,
            IReadOnlyList<KiCadImportantSignalEntry>? importantSignals = null)
        {
            this.thisExcelDataFile = excelDataFile;
            this.thisDraftFolder = DraftManager.GetBoardFolder(excelDataFile);
            this.thisBoardLabels = boardLabels;
            this.thisImportantSignals = importantSignals ?? Array.Empty<KiCadImportantSignalEntry>();

            this.Opening = BusyOverlay.HoldAsync(this, CrtWaitWording.CheckingKiCadMatches, this.RefreshReportAsync);
        }

        // The board the panel is open on - its board file's identity.
        internal string ExcelDataFile => this.thisExcelDataFile;

        // Completes when everything the panel reads on opening is in - for tests, which wait for it
        // as the user does.
        internal Task Opening { get; private set; } = Task.CompletedTask;

        // ###########################################################################################
        // Whether anything in this panel changed the draft - KiCad files imported or removed. The
        // Contribute tab reads the board on screen again ONLY then (owner report, 2026-10-10: it was
        // read again every time, changed or not) - as each such change ends (DraftChanged). Set once
        // the change is in.
        // ###########################################################################################
        internal bool ChangedTheDraft
        {
            get => this.thisChangesMade > 0;
            private set
            {
                if (value)
                {
                    this.thisChangesMade++;
                }
            }
        }

        // How many changes this panel made to the draft - a queued change that moved it tells
        // DraftChanged as it ends (KiCadDataPanel.DraftChanges.cs).
        private int thisChangesMade;

        // Told on the UI thread when a change to the draft has ended - the tab catches up with it.
        internal event EventHandler? DraftChanged;

        // Set by Close: a read still in flight then has nowhere to show what it read.
        private bool thisClosed;

        internal bool IsClosed => this.thisClosed;

        // ###########################################################################################
        // The panel closed - what it reads from now on is shown nowhere. Its queued changes go on to
        // the end (DraftChanging), each still told as it ends; the tab waits for them before taking the
        // panel away when the contributor closed it, so their "please wait" is seen.
        // ###########################################################################################
        internal void Close() => this.thisClosed = true;

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
    }

    // ###########################################################################################
    // One imported KiCad file, as the panel lists it: where it sits inside the "KiCad data" folder
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
    // One KiCad net in the report (KiCadDataPanel.NetNames.cs), marked when an important KiCad signal
    // already uses it.
    // ###########################################################################################
    public sealed class KiCadNetNameBadge
    {
        public string Name { get; init; } = string.Empty;

        public bool IsUsed { get; init; }

        public string? ToolTipText => this.IsUsed ? "Already used in \"KiCad: Important signals\"" : null;
    }
}
