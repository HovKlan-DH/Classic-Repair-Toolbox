using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Markup.Xaml;
using Handlers.MaintainerHandling;
using Handlers.DataHandling;

namespace CRT
{
    // ###########################################################################################
    // Placing a new board in CRT's drop-down lists - see the markup's header. This file lays out
    // the server's listing, lets the one row move (Controls/ListRowDrag, the Drafts tab's schematic
    // images' drag) and sends the result; BoardPlacementDisplay decides and words everything.
    //
    // *** THE QUEUE's MINUTE CHECK MUST NOT TAKE THE PANEL FROM UNDER THE MAINTAINER. *** Show is
    // called on every check. While the names or the position have been changed and not saved, it
    // leaves the panel alone; otherwise it rebuilds only when the listing for this board changed -
    // the rule the other screens keep ("nothing open is reloaded unless its row changed").
    // ###########################################################################################
    public partial class BoardPlacementView : UserControl
    {
        private readonly ListRowDrag<PlacementListRow> thisDrag;

        private ReviewApiClient? thisClient;
        private ReviewSession? thisSession;

        private UnlistedBoardEntry? thisEntry;
        private IReadOnlyList<BoardListingRow> thisListed = [];
        private PlacementListRow? thisNewRow;

        // The names or the position changed since the panel was built or saved.
        private bool thisIsDirty;

        private bool thisIsSaving;

        // Filling the boxes from code is not an edit.
        private bool thisIsFilling;

        // The message on show is the names' problem (UpdateState's own), not a save's answer.
        private bool thisShowsNameProblem;

        public BoardPlacementView()
        {
            this.InitializeComponent();

            ItemsControl list = this.GetControl<ItemsControl>("PlacementItemsControl");
            list.ItemsSource = this.Rows;

            // Only the board being placed moves, and only for an account that may place it.
            this.thisDrag = new ListRowDrag<PlacementListRow>(
                this,
                list,
                this.GetControl<ScrollViewer>("PlacementScrollViewer"),
                this.Rows,
                (_, _, _) => this.OnDropped(),
                row => row.CanDrag);

            foreach (string name in new[] { "HardwareNameBox", "BoardNameBox", "NotesBox" })
                this.GetControl<TextBox>(name).TextChanged += (_, _) => this.OnFieldsChanged();

            this.IsVisible = false;
        }

        private void InitializeComponent() => AvaloniaXamlLoader.Load(this);

        public void Initialize(ReviewApiClient? client, ReviewSession? session)
        {
            this.thisClient = client;
            this.thisSession = session;
        }

        // Told after a placement was saved, so the Boards screen can read the lists again.
        public Func<Task>? AfterSaved { get; set; }

        // BETA's list with the board being placed in it, top to bottom.
        public ObservableCollection<PlacementListRow> Rows { get; } = new();

        // The board on show, or null when the panel is hidden.
        public string? ShownBoardId => this.thisEntry?.BoardId;

        // ###########################################################################################
        // Shows the placement of `boardId` when the listing says it needs one; hides otherwise (a
        // listed board, or no listing known). See the header for what a repeat call keeps.
        // ###########################################################################################
        public void Show(BoardListingAnswer? listing, string? boardId)
        {
            UnlistedBoardEntry? entry = BoardPlacementDisplay.UnlistedEntry(listing, boardId);

            if (listing is null || entry is null)
            {
                this.Hide();
                return;
            }

            bool sameBoard = string.Equals(this.thisEntry?.BoardId, entry.BoardId, StringComparison.OrdinalIgnoreCase);

            if (sameBoard && (this.thisIsDirty || this.thisIsSaving || this.thisDrag.IsDragging))
                return;

            if (sameBoard && entry == this.thisEntry && this.thisListed.SequenceEqual(listing.Rows))
                return;

            this.Build(entry, listing.Rows, keepMessage: sameBoard);
        }

        public void Hide()
        {
            this.thisDrag.Reset();
            this.thisEntry = null;
            this.thisNewRow = null;
            this.thisListed = [];
            this.thisIsDirty = false;
            this.Rows.Clear();
            this.ShowMessage(null, isError: false);
            this.IsVisible = false;
        }

        private void Build(UnlistedBoardEntry entry, IReadOnlyList<BoardListingRow> listed, bool keepMessage)
        {
            this.thisDrag.Reset();

            this.thisEntry = entry;
            this.thisListed = listed.ToList();
            this.thisIsDirty = false;

            BoardPlacement placement = BoardPlacementDisplay.Starting(entry);
            int start = BoardPlacementDisplay.StartIndex(this.thisListed, placement.AfterExcelDataFile, out bool afterIsGone);

            this.Rows.Clear();

            foreach (BoardListingRow row in this.thisListed)
                this.Rows.Add(new PlacementListRow(row.HardwareName, row.BoardName, row.ExcelDataFile, isNewBoard: false, canDrag: false));

            this.thisNewRow = new PlacementListRow(placement.HardwareName, placement.BoardName, string.Empty, isNewBoard: true, canDrag: entry.CanPlace);
            this.Rows.Insert(start, this.thisNewRow);

            this.thisIsFilling = true;

            try
            {
                this.SetBox("HardwareNameBox", placement.HardwareName, entry.CanPlace);
                this.SetBox("BoardNameBox", placement.BoardName, entry.CanPlace);
                this.SetBox("NotesBox", placement.Notes, entry.CanPlace);
            }
            finally
            {
                this.thisIsFilling = false;
            }

            this.SetText("ExplanationText", BoardPlacementDisplay.Explanation(entry));
            this.SetText("SavedStateText", entry.CanPlace
                ? BoardPlacementDisplay.SavedState(entry, afterIsGone)
                : $"{BoardPlacementDisplay.SavedState(entry, afterIsGone)} {BoardPlacementDisplay.CannotPlaceMessage}");

            this.GetControl<TextBlock>("DragHintText").IsVisible = entry.CanPlace;
            this.GetControl<Button>("SaveButton").IsVisible = entry.CanPlace;

            if (!keepMessage)
                this.ShowMessage(null, isError: false);

            this.IsVisible = true;
            this.UpdateState();
        }

        // The row's handle - the new board's panel, the only one there is.
        private void OnRowPointerPressed(object? sender, PointerPressedEventArgs e) => this.thisDrag.Press(sender, e);

        private void OnDropped()
        {
            this.thisIsDirty = true;
            this.UpdateState();
        }

        private void OnFieldsChanged()
        {
            if (this.thisIsFilling || this.thisNewRow is null)
                return;

            this.thisIsDirty = true;
            this.thisNewRow.HardwareName = this.BoxText("HardwareNameBox");
            this.thisNewRow.BoardName = this.BoxText("BoardNameBox");
            this.UpdateState();
        }

        // Where CRT will show it, why it cannot be saved (if it cannot), and the button.
        private void UpdateState()
        {
            if (this.thisEntry is null || this.thisNewRow is null)
                return;

            int index = this.Rows.IndexOf(this.thisNewRow);

            this.SetText("WhereInCrtText", BoardPlacementDisplay.WhereInCrt(
                this.Rows.Select(row => (row.HardwareName, row.BoardName)).ToList(),
                index));

            string? problem = this.NameProblem();

            // Only the name problem this shows itself is cleared again - a refused save's words stay
            // until the maintainer does something about them.
            if (problem is not null && this.thisIsDirty)
            {
                this.ShowMessage(problem, isError: true);
                this.thisShowsNameProblem = true;
            }
            else if (problem is null && this.thisShowsNameProblem)
            {
                this.ShowMessage(null, isError: false);
                this.thisShowsNameProblem = false;
            }

            this.GetControl<Button>("SaveButton").IsEnabled =
                this.thisEntry.CanPlace && !this.thisIsSaving && problem is null;
        }

        private string? NameProblem() =>
            this.thisEntry is null
                ? null
                : BoardPlacementDisplay.NameProblem(
                    this.thisListed,
                    this.thisEntry.BoardId,
                    this.BoxText("HardwareNameBox"),
                    this.BoxText("BoardNameBox"),
                    this.BoxText("NotesBox"));

        // What Save sends: the names in the boxes, and the row above the panel as "after".
        internal SetPlacementRequest? BuildRequest()
        {
            if (this.thisEntry is null || this.thisNewRow is null)
                return null;

            List<string> listedInOrder = this.Rows
                .Where(row => !ReferenceEquals(row, this.thisNewRow))
                .Select(row => row.ExcelDataFile)
                .ToList();

            return new SetPlacementRequest(
                this.thisEntry.BoardId,
                this.BoxText("HardwareNameBox"),
                this.BoxText("BoardNameBox"),
                this.BoxText("NotesBox"),
                BoardPlacementDisplay.AfterAt(listedInOrder, this.Rows.IndexOf(this.thisNewRow)));
        }

        private async void OnSaveClick(object? sender, RoutedEventArgs e) => await this.SaveAsync();

        internal async Task SaveAsync()
        {
            SetPlacementRequest? request = this.BuildRequest();

            if (request is null || this.thisIsSaving || this.NameProblem() is not null)
                return;

            ReviewApiClient? client = this.thisClient;
            ReviewSession? session = this.thisSession;

            Func<SetPlacementRequest, CancellationToken, Task<ReviewApiResult<SetPlacementAnswer>>>? send =
                this.SaveOverrideForTests is { } answer
                    ? (placement, _) => answer(placement)
                    : client is not null && session is not null
                        ? (placement, token) => client.SetPlacementAsync(session, placement, token)
                        : null;

            if (send is null)
                return;

            this.thisIsSaving = true;
            this.UpdateState();
            this.ShowMessage(null, isError: false);

            string waiting = MaintainerWaitWording.SavingPlacement(request.BoardId ?? string.Empty);
            bool saved = false;

            try
            {
                await BusyOverlay.HoldAsync(this, waiting, async () =>
                {
                    ReviewApiResult<SetPlacementAnswer> result = await ServerWait.CallAsync(this, waiting, token => send(request, token));

                    // No answer in two minutes: the listing, read again, says whether the place was kept.
                    if (result.Failure == ReviewApiFailure.TimedOut)
                    {
                        bool? landed = null;

                        if (this.ListingOverrideForTests is { } listing)
                        {
                            landed = BoardPlacementDisplay.HasLanded(request, await listing());
                        }
                        else if (client is not null && session is not null)
                        {
                            ReviewApiResult<BoardListingAnswer> now = await ServerWait.CallAsync(
                                this, WaitWording.Checking, token => client.GetBoardListingAsync(session, token));

                            landed = now.IsOk ? BoardPlacementDisplay.HasLanded(request, now.Value!) : null;
                        }

                        saved = landed == true;
                        this.ShowMessage(MaintainerWaitWording.PlacementAfterTimeout(request.BoardId ?? string.Empty, landed), isError: !saved);
                        return;
                    }

                    if (!result.IsOk)
                    {
                        this.ShowMessage(result.Message, isError: true);
                        return;
                    }

                    saved = true;
                    this.ShowMessage(result.Value!.Message, isError: false);
                });
            }
            finally
            {
                this.thisIsSaving = false;
            }

            // Saved: what is on screen IS the placement now, so the next check may rebuild it.
            if (saved)
                this.thisIsDirty = false;

            this.UpdateState();

            if (saved && this.AfterSaved is not null)
                await ServerWait.RunAsync(this, MaintainerWaitWording.ReadingListing, this.AfterSaved);
        }

        // Answers Save without a server - for tests.
        internal Func<SetPlacementRequest, Task<ReviewApiResult<SetPlacementAnswer>>>? SaveOverrideForTests { get; set; }

        // Answers the look at the listing after a timeout, instead of the server - for tests.
        internal Func<Task<BoardListingAnswer>>? ListingOverrideForTests { get; set; }

        internal void SetNamesForTests(string hardware, string board, string notes)
        {
            this.GetControl<TextBox>("HardwareNameBox").Text = hardware;
            this.GetControl<TextBox>("BoardNameBox").Text = board;
            this.GetControl<TextBox>("NotesBox").Text = notes;
        }

        internal string TextOfForTests(string name) =>
            this.GetControl<TextBlock>(name) is { IsVisible: true } block ? block.Text ?? string.Empty : string.Empty;

        internal bool IsSaveEnabledForTests =>
            this.GetControl<Button>("SaveButton") is { IsVisible: true, IsEnabled: true };

        internal bool AreNamesEditableForTests => !this.GetControl<TextBox>("HardwareNameBox").IsReadOnly;

        private void SetBox(string name, string text, bool editable)
        {
            TextBox box = this.GetControl<TextBox>(name);
            box.Text = text;
            box.IsReadOnly = !editable;
        }

        private string BoxText(string name) => this.GetControl<TextBox>(name).Text?.Trim() ?? string.Empty;

        private void SetText(string name, string text) => this.GetControl<TextBlock>(name).Text = text;

            private void ShowMessage(string? message, bool isError)
        {
            this.thisShowsNameProblem = false;
            WindowMessage.Show(this.FindControl<TextBlock>("MessageText"), message, isError);
        }
    }

    // ###########################################################################################
    // One row of the placement list: a board already listed, or the one being placed - whose names
    // follow the boxes as they are typed, and which alone can be dragged.
    // ###########################################################################################
    public sealed class PlacementListRow : INotifyPropertyChanged, IDraggableRow
    {
        public PlacementListRow(string hardwareName, string boardName, string excelDataFile, bool isNewBoard, bool canDrag)
        {
            this.thisHardwareName = hardwareName;
            this.thisBoardName = boardName;
            this.ExcelDataFile = excelDataFile;
            this.IsNewBoard = isNewBoard;
            this.CanDrag = isNewBoard && canDrag;
        }

        public event PropertyChangedEventHandler? PropertyChanged;

        public string HardwareName
        {
            get => this.thisHardwareName;
            set => this.Set(ref this.thisHardwareName, value, nameof(this.HardwareName));
        }

        private string thisHardwareName;

        public string BoardName
        {
            get => this.thisBoardName;
            set => this.Set(ref this.thisBoardName, value, nameof(this.BoardName));
        }

        private string thisBoardName;

        // Blank for the board being placed - it has no row in the file yet.
        public string ExcelDataFile { get; }

        public bool IsNewBoard { get; }

        public bool CanDrag { get; }

        public bool ShowsAsListed => !this.IsNewBoard && !this.thisIsDropPlaceholder;

        public bool ShowsAsNewBoard => this.IsNewBoard && !this.thisIsDropPlaceholder;

        // ###########################################################################################
        // The placeholder pair ListRowDrag drives. Both notify - a binding that never heard the
        // change draws the gap at the starting height (the worklog's Files list found this out).
        // ###########################################################################################
        public bool IsDropPlaceholder
        {
            get => this.thisIsDropPlaceholder;
            set
            {
                if (this.thisIsDropPlaceholder == value)
                    return;

                this.thisIsDropPlaceholder = value;
                this.Raise(nameof(this.IsDropPlaceholder));
                this.Raise(nameof(this.ShowsAsListed));
                this.Raise(nameof(this.ShowsAsNewBoard));
            }
        }

        private bool thisIsDropPlaceholder;

        public double PlaceholderHeight
        {
            get => this.thisPlaceholderHeight;
            set
            {
                bool isMeaningfulChange = Math.Abs(this.thisPlaceholderHeight - value) >= 0.5;

                this.thisPlaceholderHeight = value;

                if (isMeaningfulChange)
                    this.Raise(nameof(this.PlaceholderHeight));
            }
        }

        private double thisPlaceholderHeight = 32.0;

        private void Set(ref string field, string value, string name)
        {
            if (string.Equals(field, value, StringComparison.Ordinal))
                return;

            field = value;
            this.Raise(name);
        }

        private void Raise(string name) => this.PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }
}
