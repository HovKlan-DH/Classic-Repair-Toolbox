using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Avalonia.Controls;
using Avalonia.Markup.Xaml;
using Avalonia.Media;
using Handlers.MaintainerHandling;
using Handlers.DataHandling;

namespace CRT
{
    // ###########################################################################################
    // "Unused files" on the Admin screen (owner decision, 2026-09-25; a window of its own until the
    // four screens replaced the windows, 2026-09-27).
    //
    // *** THE LOGIC IS ON THE SERVER AND IN Handlers/, NOT HERE. *** Which files are unused is
    // DataTreeUsage on the server; what the lines say and when Remove may be pressed is
    // UnusedFilesDisplay. This panel shows the list and sends the button press.
    //
    // *** A LIST SHOWN IS A LIST SENT. *** Remove sends exactly the paths on screen, and the server
    // removes only those it still finds unused - never one the administrator did not see.
    // ###########################################################################################
    public partial class UnusedFilesView : UserControl
    {
        private ReviewApiClient? thisClient;
        private ReviewSession? thisSession;

        // The list on screen, or null while none is. Replaced on every load, so the button can
        // never send a list for the other tree.
        private UnusedFileListing? thisListing;

        // Whether the panel has asked for a list at all yet - the tree box's first selection fires
        // while the markup is loading, before there is a client to ask with.
        private bool thisHasLoaded;

        public UnusedFilesView()
        {
            this.InitializeComponent();
        }

        private void InitializeComponent() => AvaloniaXamlLoader.Load(this);

        public void Initialize(ReviewApiClient? client, ReviewSession? session)
        {
            this.thisClient = client;
            this.thisSession = session;
        }

        private async void OnRefreshClick(object? sender, Avalonia.Interactivity.RoutedEventArgs e) =>
            await this.LoadAsync();

        // Signed out: empty, with nothing of the previous account's on screen.
        public void Clear()
        {
            this.thisHasLoaded = false;
            this.ShowListing(null);
            this.SetSummary(string.Empty);
            this.ShowMessage(null, isError: false);
        }

        private async void OnTreeSelectionChanged(object? sender, SelectionChangedEventArgs e)
        {
            if (this.thisHasLoaded)
                await this.LoadAsync();
        }

        private void OnLookedThroughChanged(object? sender, Avalonia.Interactivity.RoutedEventArgs e) =>
            this.UpdateButton();

        private string SelectedTree() =>
            (this.FindControl<ComboBox>("TreeComboBox")?.SelectedItem as ComboBoxItem)?.Tag as string ?? "beta";

        // ###########################################################################################
        // Asks the server for the selected tree's list. It reads every workbook in the tree, so it
        // says it is working rather than looking frozen.
        // ###########################################################################################
        public async Task LoadAsync()
        {
            if (this.thisClient is null || this.thisSession is null)
                return;

            this.thisHasLoaded = true;

            this.ShowListing(null);
            this.SetSummary(string.Empty);
            this.ShowMessage(null, isError: false);

            ReviewApiClient client = this.thisClient;
            ReviewSession session = this.thisSession;
            string tree = this.SelectedTree();

            // It reads every workbook in the tree, so it waits under the overlay rather than
            // looking frozen (2026-09-28).
            ReviewApiResult<UnusedFileListing> result = await ServerWait.CallAsync(
                this,
                MaintainerWaitWording.CheckingUnusedFiles,
                token => client.GetUnusedFilesAsync(session, tree, token));

            if (!result.IsOk)
            {
                this.SetSummary(string.Empty);
                this.ShowMessage(result.Message, isError: true);
                return;
            }

            this.ShowListing(result.Value);
        }

        private void ShowListing(UnusedFileListing? listing)
        {
            this.thisListing = listing;

            var files = this.FindControl<StackPanel>("FilesPanel");
            var check = this.FindControl<CheckBox>("LookedThroughCheckBox");

            files?.Children.Clear();

            if (check is not null)
            {
                // Never carried over to another list - see the markup.
                check.IsChecked = false;
                check.IsEnabled = listing is not null && listing.IsComplete && listing.Files.Count > 0;
            }

            if (listing is not null)
            {
                this.SetSummary(UnusedFilesDisplay.Summary(listing));

                foreach (UnusedFileEntry entry in listing.Files)
                {
                    files?.Children.Add(new TextBlock
                    {
                        Text = UnusedFilesDisplay.Line(entry),
                        TextWrapping = TextWrapping.Wrap
                    });
                }
            }

            if (this.FindControl<Button>("RemoveButton") is Button button)
                button.Content = UnusedFilesDisplay.RemoveButton(listing);

            this.UpdateButton();
        }

        private void UpdateButton()
        {
            var button = this.FindControl<Button>("RemoveButton");
            var check = this.FindControl<CheckBox>("LookedThroughCheckBox");

            if (button is not null)
                button.IsEnabled = UnusedFilesDisplay.CanRemove(this.thisListing, check?.IsChecked == true);
        }

        private async void OnRemoveClick(object? sender, Avalonia.Interactivity.RoutedEventArgs e)
        {
            UnusedFileListing? listing = this.thisListing;
            var check = this.FindControl<CheckBox>("LookedThroughCheckBox");
            var button = this.FindControl<Button>("RemoveButton");

            if (listing is null || button is null || this.thisClient is null || this.thisSession is null ||
                !UnusedFilesDisplay.CanRemove(listing, check?.IsChecked == true))
            {
                return;
            }

            // Off for the duration, so a second click cannot send the list twice.
            button.IsEnabled = false;
            this.ShowMessage(null, isError: false);

            ReviewApiClient client = this.thisClient;
            ReviewSession session = this.thisSession;
            List<string> paths = listing.Files.Select(file => file.Path).ToList();
            string waiting = MaintainerWaitWording.RemovingUnusedFiles(paths.Count);

            string? outcome = null;
            bool failed = true;

            await BusyOverlay.HoldAsync(this, waiting, async () =>
            {
                ReviewApiResult<UnusedFileRemovalResult> result = await ServerWait.CallAsync(
                    this, waiting, token => client.RemoveUnusedFilesAsync(session, listing.Tree, paths, token));

                // No answer in two minutes: the list, read again, says how many are still there.
                if (result.Failure == ReviewApiFailure.TimedOut)
                {
                    await this.LoadAsync();

                    int? stillThere = this.thisListing is UnusedFileListing now
                        ? now.Files.Count(file => paths.Contains(file.Path, StringComparer.Ordinal))
                        : null;

                    outcome = MaintainerWaitWording.RemovalAfterTimeout(stillThere);
                    failed = stillThere != 0;
                    return;
                }

                if (!result.IsOk)
                {
                    outcome = result.Message;
                    return;
                }

                // Reload, so what is on screen is what is left.
                await this.LoadAsync();
                outcome = UnusedFilesDisplay.Result(result.Value!);
                failed = !string.IsNullOrWhiteSpace(result.Value!.NotDoneBecause);
            });

            this.ShowMessage(outcome, isError: failed);
            this.UpdateButton();
        }

        private void SetSummary(string text)
        {
            if (this.FindControl<TextBlock>("SummaryText") is TextBlock block)
                block.Text = text;
        }

        private void ShowMessage(string? message, bool isError) =>
            WindowMessage.Show(this.FindControl<TextBlock>("MessageText"), message, isError);
    }
}
