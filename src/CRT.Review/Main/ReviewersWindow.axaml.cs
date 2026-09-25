using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Avalonia.Controls;
using Avalonia.Markup.Xaml;
using Avalonia.Media;
using CRT.Review.Handlers;

namespace CRT.Review
{
    // ###########################################################################################
    // The administrator's "Reviewers" window (Phase 6 roles, 2026-09-25): which account reviews
    // which system.
    //
    // *** THE LOGIC IS IN Handlers/, NOT HERE. *** What a line says is ReviewerAssignmentDisplay's
    // decision and what the server answered is ReviewApiParser's; both are unit tested. This file
    // resolves controls, asks the client, and shows the result - the same split ReviewMain keeps.
    //
    // *** NO AUTHORITY LIVES HERE. *** The main window offers the button only to an administrator
    // as a courtesy; the server refuses every request from anyone else regardless, and a 403 lands
    // in the message line like any other refusal.
    //
    // The two lists are RE-FETCHED after every change rather than patched locally, so what is on
    // screen is always what the server holds - a change that was refused, or one made by another
    // administrator meanwhile, cannot leave the screen describing a state that does not exist.
    // ###########################################################################################
    public partial class ReviewersWindow : Window
    {
        private readonly List<ReviewSystemRow> thisSystems = [];
        private readonly List<ReviewAccountRow> thisAccounts = [];

        private ReviewApiClient? thisClient;
        private ReviewSession? thisSession;

        // The id of the selected system, held rather than re-read off the list at press time so
        // a change cannot be sent for a row the selection moved to while a request was in flight.
        private string? thisSelectedSystemId;

        public ReviewersWindow()
        {
            this.InitializeComponent();
        }

        private void InitializeComponent() => AvaloniaXamlLoader.Load(this);

        // Called by ReviewMain before showing. The client and session are the main window's own;
        // this window does not sign in and holds no credential of its own.
        public void Initialize(ReviewApiClient client, ReviewSession session)
        {
            ArgumentNullException.ThrowIfNull(client);
            ArgumentNullException.ThrowIfNull(session);

            this.thisClient = client;
            this.thisSession = session;
        }

        protected override async void OnOpened(EventArgs e)
        {
            base.OnOpened(e);
            await this.RefreshAsync();
        }

        private async void OnRefreshClick(object? sender, Avalonia.Interactivity.RoutedEventArgs e) =>
            await this.RefreshAsync();

        private void OnCloseClick(object? sender, Avalonia.Interactivity.RoutedEventArgs e) => this.Close();

        // ###########################################################################################
        // Fetches both lists. The systems list keeps its selection across a refresh when the
        // selected system is still there, so adding a reviewer does not throw the administrator
        // back to the top of the list.
        // ###########################################################################################
        private async Task RefreshAsync()
        {
            if (this.thisClient is null || this.thisSession is null)
                return;

            this.ShowMessage(null, isError: false);

            ReviewApiResult<ReviewSystemsResponse> systems = await this.thisClient.GetSystemsAsync(this.thisSession);

            if (!systems.IsOk)
            {
                this.ShowMessage(systems.Message, isError: true);
                return;
            }

            ReviewApiResult<ReviewAccountsResponse> accounts = await this.thisClient.GetAccountsAsync(this.thisSession);

            if (!accounts.IsOk)
            {
                this.ShowMessage(accounts.Message, isError: true);
                return;
            }

            this.thisSystems.Clear();
            this.thisSystems.AddRange(systems.Value!.Systems);

            this.thisAccounts.Clear();
            this.thisAccounts.AddRange(accounts.Value!.Accounts);

            this.ApplyLists();
        }

        private void ApplyLists()
        {
            var list = this.FindControl<ListBox>("SystemsList");
            var combo = this.FindControl<ComboBox>("AccountsCombo");

            if (list is null || combo is null)
                return;

            string? keep = this.thisSelectedSystemId;

            list.ItemsSource = this.thisSystems.Select(ReviewerAssignmentDisplay.SystemLine).ToList();
            combo.ItemsSource = this.thisAccounts.Select(ReviewerAssignmentDisplay.AccountChoice).ToList();

            int index = keep is null
                ? -1
                : this.thisSystems.FindIndex(system => string.Equals(system.SystemId, keep, StringComparison.Ordinal));

            list.SelectedIndex = index;

            if (index < 0)
                this.ShowSystem(null);
        }

        private void OnSystemSelectionChanged(object? sender, SelectionChangedEventArgs e)
        {
            var list = this.FindControl<ListBox>("SystemsList");

            if (list is null)
                return;

            int index = list.SelectedIndex;

            this.ShowSystem(index >= 0 && index < this.thisSystems.Count ? this.thisSystems[index] : null);
        }

        // ###########################################################################################
        // Draws the selected system's reviewers, one line and a Remove button each.
        //
        // Built in code rather than by a DataTemplate, as ReviewMain builds its sections: the
        // button has to know which reviewer it belongs to, and closing over the row is the plainest
        // way to say so.
        // ###########################################################################################
        private void ShowSystem(ReviewSystemRow? system)
        {
            var header = this.FindControl<TextBlock>("SelectedSystemText");
            var panel = this.FindControl<StackPanel>("ReviewersPanel");
            var add = this.FindControl<Button>("AddButton");
            var combo = this.FindControl<ComboBox>("AccountsCombo");

            if (header is null || panel is null || add is null || combo is null)
                return;

            panel.Children.Clear();

            this.thisSelectedSystemId = system?.SystemId;

            add.IsEnabled = system is not null;
            combo.IsEnabled = system is not null;

            if (system is null)
            {
                header.Text = "Select a system.";
                return;
            }

            header.Text = ReviewerAssignmentDisplay.SystemLine(system);

            if (system.Reviewers.Count == 0)
            {
                panel.Children.Add(new TextBlock
                {
                    Text = "Nobody is assigned. Submissions to this system go to the administrator.",
                    Opacity = 0.7,
                    TextWrapping = TextWrapping.Wrap
                });

                return;
            }

            foreach (ReviewerRow reviewer in system.Reviewers)
            {
                var row = new Grid { ColumnDefinitions = new ColumnDefinitions("*,Auto") };

                var text = new TextBlock
                {
                    Text = ReviewerAssignmentDisplay.ReviewerLine(reviewer),
                    VerticalAlignment = Avalonia.Layout.VerticalAlignment.Center,
                    TextWrapping = TextWrapping.Wrap
                };

                var remove = new Button { Content = "Remove", Margin = new Avalonia.Thickness(8, 0, 0, 0) };

                string systemId = system.SystemId;
                long accountId = reviewer.AccountId;

                remove.Click += async (_, _) => await this.RemoveAsync(systemId, accountId);

                Grid.SetColumn(remove, 1);
                row.Children.Add(text);
                row.Children.Add(remove);
                panel.Children.Add(row);
            }
        }

        private async void OnAddClick(object? sender, Avalonia.Interactivity.RoutedEventArgs e)
        {
            var combo = this.FindControl<ComboBox>("AccountsCombo");

            if (combo is null || this.thisClient is null || this.thisSession is null || this.thisSelectedSystemId is null)
                return;

            int index = combo.SelectedIndex;

            if (index < 0 || index >= this.thisAccounts.Count)
            {
                this.ShowMessage("Choose an account first.", isError: true);
                return;
            }

            ReviewAccountRow account = this.thisAccounts[index];

            // Said here, before the round trip, in the same words the list already shows. The
            // server refuses it regardless.
            string? why = ReviewerAssignmentDisplay.WhyNotGrantable(account);

            if (why is not null)
            {
                this.ShowMessage($"{account.Email} cannot be a reviewer: {why}.", isError: true);
                return;
            }

            ReviewApiResult<string> result = await this.thisClient.AddReviewerAsync(
                this.thisSession, this.thisSelectedSystemId, account.Id);

            if (!result.IsOk)
            {
                this.ShowMessage(result.Message, isError: true);
                return;
            }

            await this.RefreshAsync();
            this.ShowMessage($"{account.DisplayName} now reviews {this.thisSelectedSystemId}.", isError: false);
        }

        private async Task RemoveAsync(string systemId, long accountId)
        {
            if (this.thisClient is null || this.thisSession is null)
                return;

            ReviewApiResult<string> result = await this.thisClient.RemoveReviewerAsync(
                this.thisSession, systemId, accountId);

            if (!result.IsOk)
            {
                this.ShowMessage(result.Message, isError: true);
                return;
            }

            await this.RefreshAsync();
            this.ShowMessage("Removed. It takes effect on their next request.", isError: false);
        }

        private void ShowMessage(string? message, bool isError) =>
            WindowMessage.Show(this.FindControl<TextBlock>("MessageText"), message, isError);
    }
}
