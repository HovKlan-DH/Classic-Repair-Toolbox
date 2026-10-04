using System.Threading.Tasks;
using Avalonia.Controls;
using Avalonia.Markup.Xaml;
using Avalonia.Media;
using Handlers.MaintainerHandling;

namespace CRT
{
    // ###########################################################################################
    // "Rebuild checksum manifests" on the Account screen (owner request, 2026-10-01). See the markup
    // for what it is for and why it asks nothing first.
    //
    // *** EVERY WORD SHOWN IS THE SERVER'S. *** The headline and each tree's line come from
    // ManifestRebuildFlow, so which trees exist, what counts as skipped and how a failure is
    // described are all decided in one place - the place that actually knows. This panel presses
    // the button and prints the answer.
    //
    // *** IT DOES NOTHING ON ITS OWN. *** Unlike "Unused files", nothing is read when the screen is
    // shown: a rebuild WRITES, so it happens only when the administrator asks for it. Choosing this
    // item just shows the panel.
    // ###########################################################################################
    public partial class RebuildManifestsView : UserControl
    {
        private ReviewApiClient? thisClient;
        private ReviewSession? thisSession;

        public RebuildManifestsView()
        {
            this.InitializeComponent();
        }

        private void InitializeComponent() => AvaloniaXamlLoader.Load(this);

        public void Initialize(ReviewApiClient? client, ReviewSession? session)
        {
            this.thisClient = client;
            this.thisSession = session;
        }

        // Signed out: nothing of the previous account's result left on screen.
        public void Clear()
        {
            this.ShowResult(null);
            this.ShowMessage(null, isError: false);
        }

        private async void OnRebuildClick(object? sender, Avalonia.Interactivity.RoutedEventArgs e) =>
            await this.RebuildAsync();

        // ###########################################################################################
        // Asks the server to rebuild both trees. Under the overlay (CRT's own, up the tree - this
        // tab has none): the server hashes every file in each tree, which is seconds on a real
        // data tree and would otherwise look like a dead button.
        // ###########################################################################################
        private async Task RebuildAsync()
        {
            if (this.thisClient is null || this.thisSession is null)
                return;

            this.ShowResult(null);
            this.ShowMessage(null, isError: false);

            ReviewApiClient client = this.thisClient;
            ReviewSession session = this.thisSession;

            ReviewApiResult<ManifestRebuildResult> result = await ServerWait.CallAsync(
                this,
                MaintainerWaitWording.RebuildingManifests,
                token => client.RebuildManifestsAsync(session, token));

            if (!result.IsOk)
            {
                // A timeout is never "it failed": the server may well have rebuilt them - see
                // MaintainerWaitWording.RebuildAfterTimeout.
                this.ShowMessage(
                    result.Failure == ReviewApiFailure.TimedOut ? MaintainerWaitWording.RebuildAfterTimeout : result.Message,
                    isError: true);
                return;
            }

            this.ShowResult(result.Value);
        }

        // ###########################################################################################
        // The headline and one line per tree. A FAILED tree is drawn in CRT's red, so a rebuild that
        // half worked cannot be read as a success - the headline alone says "1 of 2" and that is
        // easy to skim past.
        // ###########################################################################################
        private void ShowResult(ManifestRebuildResult? result)
        {
            var headline = this.FindControl<TextBlock>("HeadlineText");
            var panel = this.FindControl<StackPanel>("ResultsPanel");

            panel?.Children.Clear();

            if (headline is not null)
            {
                headline.Text = result?.Headline ?? string.Empty;
                headline.IsVisible = result is not null;
            }

            if (result is null)
                return;

            foreach (ManifestRebuildTreeResult tree in result.Trees)
            {
                var line = new TextBlock
                {
                    Text = tree.Message,
                    TextWrapping = TextWrapping.Wrap
                };

                // The same red every refusal in this tab uses (WindowMessage).
                if (!tree.Skipped && tree.Entries < 0)
                    line.Foreground = Brushes.IndianRed;

                panel?.Children.Add(line);
            }
        }

        private void ShowMessage(string? message, bool isError) =>
            WindowMessage.Show(this.FindControl<TextBlock>("MessageText"), message, isError);
    }
}
