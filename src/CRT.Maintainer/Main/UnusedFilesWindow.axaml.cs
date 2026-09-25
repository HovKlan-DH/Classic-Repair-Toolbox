using System;
using System.Linq;
using System.Threading.Tasks;
using Avalonia.Controls;
using Avalonia.Markup.Xaml;
using Avalonia.Media;
using CRT.Maintainer.Handlers;
using Handlers.DataHandling;

namespace CRT.Maintainer
{
    // ###########################################################################################
    // The administrator's "Unused files" window (owner decision, 2026-09-25).
    //
    // *** THE LOGIC IS ON THE SERVER AND IN Handlers/, NOT HERE. *** Which files are unused is
    // DataTreeUsage on the server; what the lines say and when Remove may be pressed is
    // UnusedFilesDisplay. This window shows the list and sends the button press.
    //
    // *** A LIST SHOWN IS A LIST SENT. *** Remove sends exactly the paths on screen, and the server
    // removes only those it still finds unused - never one the administrator did not see.
    // ###########################################################################################
    public partial class UnusedFilesWindow : Window
    {
        private ReviewApiClient? thisClient;
        private ReviewSession? thisSession;

        // The list on screen, or null while none is. Replaced on every load, so the button can
        // never send a list for the other tree.
        private UnusedFileListing? thisListing;

        public UnusedFilesWindow()
        {
            this.InitializeComponent();
        }

        private void InitializeComponent() => AvaloniaXamlLoader.Load(this);

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
            await this.LoadAsync();
        }

        private async void OnRefreshClick(object? sender, Avalonia.Interactivity.RoutedEventArgs e) =>
            await this.LoadAsync();

        private async void OnTreeSelectionChanged(object? sender, SelectionChangedEventArgs e)
        {
            // The first selection fires while the window is still loading; OnOpened loads then.
            if (this.IsLoaded)
                await this.LoadAsync();
        }

        private void OnCloseClick(object? sender, Avalonia.Interactivity.RoutedEventArgs e) => this.Close();

        private void OnLookedThroughChanged(object? sender, Avalonia.Interactivity.RoutedEventArgs e) =>
            this.UpdateButton();

        private string SelectedTree() =>
            (this.FindControl<ComboBox>("TreeComboBox")?.SelectedItem as ComboBoxItem)?.Tag as string ?? "beta";

        // ###########################################################################################
        // Asks the server for the selected tree's list. It reads every workbook in the tree, so it
        // says it is working rather than looking frozen.
        // ###########################################################################################
        private async Task LoadAsync()
        {
            if (this.thisClient is null || this.thisSession is null)
                return;

            this.ShowListing(null);
            this.SetSummary("Checking every file against every workbook...");
            this.ShowMessage(null, isError: false);

            ReviewApiResult<UnusedFileListing> result =
                await this.thisClient.GetUnusedFilesAsync(this.thisSession, this.SelectedTree());

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
            this.ShowMessage("Removing...", isError: false);

            ReviewApiResult<UnusedFileRemovalResult> result = await this.thisClient.RemoveUnusedFilesAsync(
                this.thisSession, listing.Tree, listing.Files.Select(file => file.Path).ToList());

            if (!result.IsOk)
            {
                this.ShowMessage(result.Message, isError: true);
                this.UpdateButton();
                return;
            }

            // Reload, so what is on screen is what is left.
            await this.LoadAsync();
            this.ShowMessage(UnusedFilesDisplay.Result(result.Value!), isError: !string.IsNullOrWhiteSpace(result.Value!.NotDoneBecause));
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
