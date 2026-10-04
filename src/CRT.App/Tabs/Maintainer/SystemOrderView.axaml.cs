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
using Handlers.DataHandling;
using Handlers.MaintainerHandling;

namespace CRT
{
    // ###########################################################################################
    // Account > "Order of systems" - see the markup for what it is; SystemOrderDisplay words it and
    // CRT.Server's SystemOrderFlow writes it.
    //
    // *** READ WHEN CHOSEN, AND NEVER UNDER A MOVE NOT SAVED. *** BETA's list (the Systems screen's
    // own GET /api/review/systems/listing) is read when this item is chosen on the Account screen and
    // after a save. Coming back to the item with moves not saved keeps them - reading again would
    // throw the work away.
    //
    // *** THE WHOLE LIST IS SENT. *** Save sends every system on screen in its new order; the server
    // refuses it (409) when BETA's list changed meanwhile, and the message says to open it again.
    // ###########################################################################################
    public partial class SystemOrderView : UserControl
    {
        private readonly ListRowDrag<OrderListRow> thisDrag;

        private ReviewApiClient? thisClient;
        private ReviewSession? thisSession;

        // The system ids as last read, top to bottom - what "moved" is measured against.
        private IReadOnlyList<SystemListingRow> thisAsRead = [];

        private bool thisIsSaving;

        public SystemOrderView()
        {
            this.InitializeComponent();

            ItemsControl list = this.GetControl<ItemsControl>("OrderItemsControl");
            list.ItemsSource = this.Rows;

            this.thisDrag = new ListRowDrag<OrderListRow>(
                this,
                list,
                this.GetControl<ScrollViewer>("OrderScrollViewer"),
                this.Rows,
                (_, _, _) => this.UpdateState());

            this.GetControl<TextBlock>("HeadingText").Text = SystemOrderDisplay.Heading;
            this.GetControl<TextBlock>("ExplanationText").Text = SystemOrderDisplay.Explanation;
            this.GetControl<Button>("SaveButton").Content = SystemOrderDisplay.SaveButton;
            this.GetControl<Button>("UndoButton").Content = SystemOrderDisplay.UndoButton;
        }

        private void InitializeComponent() => AvaloniaXamlLoader.Load(this);

        public void Initialize(ReviewApiClient? client, ReviewSession? session)
        {
            this.thisClient = client;
            this.thisSession = session;
        }

        // BETA's list as it stands on screen, top to bottom.
        public ObservableCollection<OrderListRow> Rows { get; } = new();

        // Whether the list on screen has been moved and not saved.
        internal bool IsReordered =>
            SystemOrderDisplay.IsReordered(
                this.thisAsRead.Select(row => row.SystemId).ToList(),
                this.Rows.Select(row => row.SystemId).ToList());

        // Signed out: nothing of the previous account's list left.
        public void Clear()
        {
            this.thisDrag.Reset();
            this.thisAsRead = [];
            this.Rows.Clear();
            this.ShowMessage(null, isError: false);
            this.UpdateState();
        }

        // ###########################################################################################
        // Reads BETA's list under CRT's overlay - unless moves are waiting to be saved (see the
        // header). A list that cannot be read leaves what is shown, with the reason.
        // ###########################################################################################
        public async Task LoadAsync()
        {
            if (this.IsReordered || this.thisIsSaving)
                return;

            if (this.thisClient is not ReviewApiClient client || this.thisSession is not ReviewSession session)
                return;

            ReviewApiResult<SystemListingAnswer> listing = await ServerWait.CallAsync(
                this, MaintainerWaitWording.ReadingListing, token => client.GetSystemListingAsync(session, token));

            if (!listing.IsOk)
            {
                this.ShowMessage(listing.Message, isError: true);
                return;
            }

            this.ShowMessage(null, isError: false);
            this.UseListing(listing.Value!);
        }

        // ###########################################################################################
        // The list as read: every row a panel that moves. With no list at all, says so - never an
        // empty box that reads as "there are no systems".
        // ###########################################################################################
        internal void UseListing(SystemListingAnswer listing)
        {
            ArgumentNullException.ThrowIfNull(listing);

            this.thisDrag.Reset();
            this.thisAsRead = listing.HasList ? listing.Rows.ToList() : [];

            this.Rows.Clear();

            foreach (SystemListingRow row in this.thisAsRead)
                this.Rows.Add(new OrderListRow(row.SystemId, row.HardwareName, row.BoardName));

            if (!listing.HasList)
                this.ShowMessage(SystemOrderDisplay.NoList, isError: true);

            this.UpdateState();
        }

        private void OnRowPointerPressed(object? sender, PointerPressedEventArgs e) => this.thisDrag.Press(sender, e);

        // Save and "Cancel" only once something moved; the line under the list says so.
        private void UpdateState()
        {
            bool moved = this.IsReordered;

            this.GetControl<Button>("SaveButton").IsEnabled = moved && !this.thisIsSaving;
            this.GetControl<Button>("UndoButton").IsEnabled = moved && !this.thisIsSaving;

            if (moved)
                this.ShowMessage(SystemOrderDisplay.Unsaved, isError: false);
            else if (this.FindControl<TextBlock>("MessageText")?.Text == SystemOrderDisplay.Unsaved)
                this.ShowMessage(null, isError: false);
        }

        private void OnUndoClick(object? sender, RoutedEventArgs e) => this.PutBack();

        // The list back as it was read - the moves dropped.
        internal void PutBack()
        {
            this.thisDrag.Reset();

            IReadOnlyList<SystemListingRow> asRead = this.thisAsRead;
            this.Rows.Clear();

            foreach (SystemListingRow row in asRead)
                this.Rows.Add(new OrderListRow(row.SystemId, row.HardwareName, row.BoardName));

            this.UpdateState();
        }

        // Moves one row, as a drag would - for tests.
        internal void MoveForTests(int from, int to)
        {
            this.Rows.Move(from, to);
            this.UpdateState();
        }

        private async void OnSaveClick(object? sender, RoutedEventArgs e) => await this.SaveAsync();

        // ###########################################################################################
        // Sends the order, says what the server did, and reads the list again - so what is on screen
        // is the list as BETA now holds it. A refusal keeps the moves on screen with the server's
        // words; no answer in two minutes is said, and the next look at the item reads the truth.
        // ###########################################################################################
        internal async Task SaveAsync()
        {
            if (!this.IsReordered || this.thisIsSaving)
                return;

            // Each system once - see SystemOrderDisplay.OrderToSend.
            IReadOnlyList<string> order = SystemOrderDisplay.OrderToSend(this.Rows.Select(row => row.SystemId));

            Func<IReadOnlyList<string>, CancellationToken, Task<ReviewApiResult<SystemOrderAnswer>>>? send =
                this.SaveOverrideForTests is { } answer
                    ? (ids, _) => answer(ids)
                    : this.thisClient is ReviewApiClient client && this.thisSession is ReviewSession session
                        ? (ids, token) => client.SetSystemOrderAsync(session, ids, token)
                        : null;

            if (send is null)
                return;

            this.thisIsSaving = true;
            this.UpdateState();

            (string Text, bool IsError)? said = null;

            try
            {
                ReviewApiResult<SystemOrderAnswer> result = await ServerWait.CallAsync(
                    this, MaintainerWaitWording.SavingSystemOrder, token => send(order, token));

                if (result.Failure == ReviewApiFailure.TimedOut)
                {
                    this.ShowMessage(WaitWording.NoAnswer, isError: true);
                    return;
                }

                if (!result.IsOk)
                {
                    this.ShowMessage(result.Message, isError: true);
                    return;
                }

                said = SystemOrderDisplay.Saved(result.Value!);
                this.ShowMessage(said.Value.Text, said.Value.IsError);
            }
            finally
            {
                this.thisIsSaving = false;
            }

            if (said is { } line)
            {
                // What was sent IS the list now - so a list that cannot be read again still shows no
                // unsaved moves.
                this.thisAsRead = this.Rows.Select(row => new SystemListingRow(row.SystemId, row.HardwareName, row.BoardName, string.Empty)).ToList();

                await this.LoadAsync();

                // The save's own sentence stays on screen after the list is read again.
                this.ShowMessage(line.Text, line.IsError);
            }

            this.UpdateState();
        }

        // Answers Save without a server - for tests; sees the order sent.
        internal Func<IReadOnlyList<string>, Task<ReviewApiResult<SystemOrderAnswer>>>? SaveOverrideForTests { get; set; }

        internal string MessageForTests =>
            this.GetControl<TextBlock>("MessageText") is { IsVisible: true } block ? block.Text ?? string.Empty : string.Empty;

        internal bool IsSaveEnabledForTests => this.GetControl<Button>("SaveButton").IsEnabled;

        private void ShowMessage(string? message, bool isError) =>
            WindowMessage.Show(this.FindControl<TextBlock>("MessageText"), message, isError);
    }

    // ###########################################################################################
    // One system in the order list, and the placeholder pair ListRowDrag drives. Both notify - a
    // binding that never heard the change draws the gap at the starting height.
    // ###########################################################################################
    public sealed class OrderListRow : INotifyPropertyChanged, IDraggableRow
    {
        public OrderListRow(string systemId, string hardwareName, string boardName)
        {
            this.SystemId = systemId;
            this.HardwareName = hardwareName;
            this.BoardName = boardName;
        }

        public event PropertyChangedEventHandler? PropertyChanged;

        public string SystemId { get; }

        public string HardwareName { get; }

        public string BoardName { get; }

        public bool ShowsRow => !this.thisIsDropPlaceholder;

        public bool IsDropPlaceholder
        {
            get => this.thisIsDropPlaceholder;
            set
            {
                if (this.thisIsDropPlaceholder == value)
                    return;

                this.thisIsDropPlaceholder = value;
                this.Raise(nameof(this.IsDropPlaceholder));
                this.Raise(nameof(this.ShowsRow));
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

        private double thisPlaceholderHeight = 28.0;

        private void Raise(string name) => this.PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }
}
