using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Avalonia.Controls;
using Avalonia.Markup.Xaml;
using Avalonia.Media;
using CRT.Review.Handlers;
using Handlers.DataHandling;

namespace CRT.Review
{
    // ###########################################################################################
    // The "Publish to production" window (maintainer request, 2026-09-25) - BETA's newer state of
    // one system, copied to what every user downloads.
    //
    // *** THE LOGIC IS IN Handlers/ AND ON THE SERVER, NOT HERE. *** What the lines say is
    // ProductionDisplay's, what may be pressed is ProductionDisplay.CanPress over the server's own
    // canPublish, and what is copied is ProductionPromotionPlan on the server - the same code that
    // performs the copy. This file shows those answers and sends the button press.
    //
    // *** A PLAN SHOWN IS A PLAN SENT. *** The publish request carries back the BETA content hash
    // this window was shown, and the server refuses when BETA has moved since - so what the
    // reviewer checked is what goes out, and a 409 says to look again rather than to retry.
    // ###########################################################################################
    public partial class ProductionWindow : Window
    {
        private readonly List<ProductionSystemRow> thisSystems = [];

        private ReviewApiClient? thisClient;
        private ReviewSession? thisSession;

        // The plan on screen. Null while none is, and replaced whenever the selection moves, so the
        // button can never send a hash for a system other than the one described.
        private ProductionPlanView? thisPlan;

        public ProductionWindow()
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
            await this.RefreshAsync();
        }

        private async void OnRefreshClick(object? sender, Avalonia.Interactivity.RoutedEventArgs e) =>
            await this.RefreshAsync();

        private void OnCloseClick(object? sender, Avalonia.Interactivity.RoutedEventArgs e) => this.Close();

        private async Task RefreshAsync()
        {
            if (this.thisClient is null || this.thisSession is null)
                return;

            this.ShowMessage(null, isError: false);

            ReviewApiResult<ProductionListResponse> result = await this.thisClient.GetProductionListAsync(this.thisSession);

            var list = this.FindControl<ListBox>("SystemsList");

            if (list is null)
                return;

            if (!result.IsOk)
            {
                // The list is left as it was: "could not ask" must not read as "nothing waiting".
                this.ShowListMessage(result.Message, isError: true);
                return;
            }

            this.thisSystems.Clear();
            this.thisSystems.AddRange(result.Value!.Systems);

            list.ItemsSource = this.thisSystems.Select(ProductionDisplay.ListEntry).ToList();
            list.SelectedIndex = -1;
            this.ShowPlan(null, null);

            if (!result.Value.Configured)
                this.ShowListMessage("Publishing to production is not switched on for this server.", isError: true);
            else if (this.thisSystems.Count == 0)
                this.ShowListMessage("Production is up to date with BETA for every system you review.", isError: false);
            else
                this.ShowListMessage(null, isError: false);
        }

        private async void OnSystemSelectionChanged(object? sender, SelectionChangedEventArgs e)
        {
            var list = this.FindControl<ListBox>("SystemsList");

            if (list is null || this.thisClient is null || this.thisSession is null)
                return;

            int index = list.SelectedIndex;

            if (index < 0 || index >= this.thisSystems.Count)
            {
                this.ShowPlan(null, null);
                return;
            }

            ProductionSystemRow row = this.thisSystems[index];

            this.ShowPlan(null, row);
            this.SetText("PlanSummaryText", "Working out what would be copied...");

            ReviewApiResult<ProductionPlanView> plan = await this.thisClient.GetProductionPlanAsync(this.thisSession, row.SystemId);

            // The selection may have moved while that was in flight; an answer for another row is
            // dropped rather than drawn under this one's name.
            if (list.SelectedIndex != index)
                return;

            if (!plan.IsOk)
            {
                this.SetText("PlanSummaryText", string.Empty);
                this.ShowMessage(plan.Message, isError: true);
                return;
            }

            this.ShowPlan(plan.Value, row);
        }

        private void ShowPlan(ProductionPlanView? plan, ProductionSystemRow? row)
        {
            this.thisPlan = plan;

            var files = this.FindControl<StackPanel>("FilesPanel");
            var check = this.FindControl<CheckBox>("CheckedInBetaCheckBox");

            if (files is null || check is null)
                return;

            files.Children.Clear();

            // Never carried from one board to the next - see the markup.
            check.IsChecked = false;
            check.IsEnabled = plan is not null && plan.CanPublish;

            this.SetText("SelectedSystemText", row is null ? "Select a system." : ProductionDisplay.SystemLine(row));
            this.SetText("PlanSummaryText", plan is null ? string.Empty : ProductionDisplay.PlanSummary(plan));

            // The two-person approval for a shared-file change, and a button that says what
            // pressing it does - both from the server's ApprovalStatus.
            string? approvalLine = ApprovalWording.StatusLine(plan?.Approval);
            this.SetText("ApprovalStatusText", approvalLine ?? string.Empty);

            if (this.FindControl<TextBlock>("ApprovalStatusText") is TextBlock approvalText)
                approvalText.IsVisible = approvalLine is not null;

            if (this.FindControl<Button>("PublishButton") is Button publish)
                publish.Content = ApprovalWording.ApproveButton(plan?.Approval, "production");

            if (plan is not null)
            {
                foreach (ReviewFindingView problem in plan.Problems)
                {
                    files.Children.Add(new TextBlock
                    {
                        Text = problem.Message,
                        TextWrapping = TextWrapping.Wrap,
                        Foreground = Brushes.IndianRed
                    });
                }

                // What publishing REMOVES from production, FIRST and in red: the least visible change
                // and the one that cannot be undone (2026-09-25). Said even when it is nothing, so
                // the reviewer knows it was checked.
                string? removals = FileRemovalWording.Headline(plan.Removals, "production");

                if (removals is not null)
                {
                    bool any = plan.Removals!.Files.Count > 0;

                    var headline = new TextBlock
                    {
                        Text = removals,
                        TextWrapping = TextWrapping.Wrap,
                        FontWeight = any ? FontWeight.SemiBold : FontWeight.Normal,
                        Margin = new Avalonia.Thickness(0, 0, 0, 2)
                    };

                    if (any)
                        headline.Foreground = Brushes.IndianRed;

                    files.Children.Add(headline);

                    foreach (string path in plan.Removals.Files)
                    {
                        files.Children.Add(new TextBlock
                        {
                            Text = FileRemovalWording.Line(path),
                            TextWrapping = TextWrapping.Wrap,
                            Foreground = Brushes.IndianRed,
                            FontWeight = FontWeight.SemiBold
                        });
                    }

                    files.Children.Add(new TextBlock { Text = " " });
                }

                foreach (PromotionFile file in plan.Files)
                {
                    files.Children.Add(new TextBlock
                    {
                        Text = ProductionDisplay.FileLine(file),
                        TextWrapping = TextWrapping.Wrap,
                        FontWeight = file.IsShared ? FontWeight.SemiBold : FontWeight.Normal
                    });
                }

                this.ShowMessage(plan.Refusal, isError: plan.Refusal is not null);
            }

            this.UpdateButton();
        }

        private void OnCheckedInBetaChanged(object? sender, Avalonia.Interactivity.RoutedEventArgs e) =>
            this.UpdateButton();

        private void UpdateButton()
        {
            var button = this.FindControl<Button>("PublishButton");
            var check = this.FindControl<CheckBox>("CheckedInBetaCheckBox");

            if (button is not null)
                button.IsEnabled = ProductionDisplay.CanPress(this.thisPlan, check?.IsChecked == true);
        }

        private async void OnPublishClick(object? sender, Avalonia.Interactivity.RoutedEventArgs e)
        {
            ProductionPlanView? plan = this.thisPlan;
            var check = this.FindControl<CheckBox>("CheckedInBetaCheckBox");
            var button = this.FindControl<Button>("PublishButton");

            if (plan is null || button is null || this.thisClient is null || this.thisSession is null ||
                !ProductionDisplay.CanPress(plan, check?.IsChecked == true))
            {
                return;
            }

            // Off for the duration, so a second click cannot send a second publish.
            button.IsEnabled = false;
            this.ShowMessage("Publishing to production...", isError: false);

            // The removal list sent back is the one on screen; the server refuses if it would now
            // remove anything else.
            ReviewApiResult<ProductionPublishResult> result = await this.thisClient.PublishToProductionAsync(
                this.thisSession, plan.SystemId, plan.BetaContentHash, plan.Removals?.Files ?? []);

            if (!result.IsOk)
            {
                this.ShowMessage(result.Message, isError: true);
                this.UpdateButton();
                return;
            }

            // The first of two approvals: recorded, nothing copied.
            string done = result.Value!.IsAwaitingApproval
                ? ApprovalWording.Recorded(result.Value.WaitingFor, "production")
                : $"{plan.SystemId} is published to production ({result.Value.FilesCopied} file(s) copied)." +
                  FileRemovalWording.Done(result.Value.RemovedFiles) +
                  " Everyone gets it the next time their data updates.";

            await this.RefreshAsync();
            this.ShowMessage(done, isError: false);
        }

        private void SetText(string name, string text)
        {
            var block = this.FindControl<TextBlock>(name);

            if (block is not null)
                block.Text = text;
        }

        private void ShowMessage(string? message, bool isError) =>
            WindowMessage.Show(this.FindControl<TextBlock>("MessageText"), message, isError);

        private void ShowListMessage(string? message, bool isError) =>
            WindowMessage.Show(this.FindControl<TextBlock>("ListMessageText"), message, isError);
    }
}
