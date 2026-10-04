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
    // The right-hand side of the "BETA" screen (owner request, 2026-09-27) - one system's BETA
    // state, copied to what every user downloads, or pushed back to the queue. It was the "Publish
    // to production" window (2026-09-25) until the four screens replaced the windows; the list it
    // sat beside is now TabMaintainer.Beta.cs.
    //
    // *** THE LOGIC IS IN Handlers/ AND ON THE SERVER, NOT HERE. *** What the lines say is
    // ProductionDisplay's, what may be pressed is ProductionDisplay.CanPress over the server's own
    // canPublish, and what is copied is ProductionPromotionPlan on the server - the same code that
    // performs the copy. This file shows those answers and sends the button press.
    //
    // *** A PLAN SHOWN IS A PLAN SENT. *** The publish request carries back the BETA content hash
    // this panel was shown, and the server refuses when BETA has moved since - so what the
    // maintainer checked is what goes out, and a 409 says to look again rather than to retry.
    // ###########################################################################################
    public partial class BetaView : UserControl
    {
        private ReviewApiClient? thisClient;
        private ReviewSession? thisSession;

        // The plan on screen. Null while none is, and replaced whenever the system shown changes, so
        // the button can never send a hash for a system other than the one described.
        private ProductionPlanView? thisPlan;

        // The system on screen, and a counter that drops a plan answer the selection has moved past.
        private ProductionSystemRow? thisRow;
        private int thisRequest;

        public BetaView()
        {
            this.InitializeComponent();
        }

        private void InitializeComponent() => AvaloniaXamlLoader.Load(this);

        // The main window's client and session - this panel signs in to nothing of its own.
        public void Initialize(ReviewApiClient? client, ReviewSession? session)
        {
            this.thisClient = client;
            this.thisSession = session;
        }

        // ###########################################################################################
        // What the main window does after a publish or a push-back changed things: read the lists
        // again (the system leaves BETA's, a rollback's submissions return to the queue). Awaited
        // BEFORE the outcome is said, so re-showing the system cannot clear the sentence saying
        // what just happened.
        // ###########################################################################################
        public Func<Task>? AfterChange { get; set; }

        // ###########################################################################################
        // "PLEASE WAIT" OVER THE WHOLE WINDOW (owner request, 2026-09-27) while a push-back or a
        // publish to production runs - the main window's BusyOverlay, found from this panel
        // (BusyOverlay.For). Since 2026-09-28 every wait in the application uses it, with its
        // two-minute limit; see ServerWait.
        // ###########################################################################################

        // The system on screen, or null.
        public ProductionSystemRow? ShownRow => this.thisRow;

        // ###########################################################################################
        // Shows one system: its name at once, then the plan when the server has worked it out. Null
        // empties the panel. An answer arriving after another system was chosen is dropped rather
        // than drawn under that one's name.
        // ###########################################################################################
        public async Task ShowSystemAsync(ProductionSystemRow? row)
        {
            int request = ++this.thisRequest;

            this.ShowMessage(null, isError: false);
            this.ShowPlan(null, row);

            if (row is null || this.thisClient is null || this.thisSession is null)
                return;

            this.SetText("PlanSummaryText", "Working out what would be copied...");

            ReviewApiResult<ProductionPlanView> plan = await this.thisClient.GetProductionPlanAsync(this.thisSession, row.SystemId);

            if (request != this.thisRequest)
                return;

            if (!plan.IsOk)
            {
                this.SetText("PlanSummaryText", string.Empty);
                this.ShowMessage(plan.Message, isError: true);
                return;
            }

            this.ShowPlan(plan.Value, row);
        }

        // ###########################################################################################
        // The system on screen has left the list because somebody else published it or pushed it
        // back - found by the queue's own check. The panel empties and says why, rather than keep
        // offering buttons for a BETA state that no longer waits.
        // ###########################################################################################
        public void ShowGoneElsewhere()
        {
            this.thisRequest++;
            this.ShowPlan(null, null);
            this.ShowMessage("This system is no longer waiting to go to stable - it was published or pushed back meanwhile.", isError: true);
        }

        // Signed out: empty, with nothing of the previous account's on screen.
        public void Clear()
        {
            this.thisRequest++;
            this.ShowPlan(null, null);
            this.ShowMessage(null, isError: false);
        }

        // Says one sentence under the panel - for the main window, and for the outcomes below.
        public void ShowMessage(string? message, bool isError) =>
            WindowMessage.Show(this.FindControl<TextBlock>("MessageText"), message, isError);

        // ###########################################################################################
        // WHOSE WORK THIS CARRIES, at the top of the panel (owner request, 2026-09-27) - see the
        // markup for why this comes before the files rather than after them.
        //
        // The headline is shown even when nothing is carried: a board brought level by hand really
        // does carry no contribution, and leaving the space blank would read as though the question
        // had not been asked.
        // ###########################################################################################
        private static void ShowCarrying(StackPanel panel, ProductionPlanView plan)
        {
            IReadOnlyList<CarriedSubmission> carrying = plan.Carrying ?? [];

            panel.Children.Add(new TextBlock
            {
                Text = ProductionDisplay.CarryingHeadline(carrying),
                TextWrapping = TextWrapping.Wrap,
                FontWeight = carrying.Count > 0 ? FontWeight.SemiBold : FontWeight.Normal,
                Margin = new Avalonia.Thickness(0, 0, 0, 2)
            });

            DateTimeOffset now = DateTimeOffset.UtcNow;

            foreach (CarriedSubmission submission in carrying)
            {
                panel.Children.Add(new TextBlock
                {
                    Text = ProductionDisplay.CarryingLine(submission, now),
                    TextWrapping = TextWrapping.Wrap
                });

                // Its contributor discarded their own draft since (owner request, 2026-09-28): said
                // under their line, with what to do - push it back and ask them.
                if (submission.DraftDiscardedUtc is not null)
                {
                    panel.Children.Add(new TextBlock
                    {
                        Text = DraftDiscardWording.BetaWarning(submission),
                        TextWrapping = TextWrapping.Wrap,
                        FontWeight = FontWeight.SemiBold,
                        Foreground = Brushes.IndianRed,
                        Margin = new Avalonia.Thickness(0, 0, 0, 4)
                    });
                }
            }
        }

        // ###########################################################################################
        // Draws a plan into the panel without a server - the panel is otherwise only drawn from a
        // live plan request, and would be pinned by nothing at all.
        // ###########################################################################################
        internal void ShowPlanForTests(ProductionPlanView? plan, ProductionSystemRow? row) =>
            this.ShowPlan(plan, row);

        private void ShowPlan(ProductionPlanView? plan, ProductionSystemRow? row)
        {
            this.thisPlan = plan;
            this.thisRow = row;

            var check = this.FindControl<CheckBox>("CheckedInBetaCheckBox");
            var carrying = this.FindControl<StackPanel>("CarryingPanel");
            var attention = this.FindControl<StackPanel>("AttentionPanel");

            if (check is null || carrying is null || attention is null)
                return;

            carrying.Children.Clear();
            attention.Children.Clear();

            // ###########################################################################################
            // THE FILE TREE (owner request, 2026-09-28): production after the publish, from the plan
            // - its copies, its removals and what it found already the same. Emptied HERE, with the
            // panels it sits under, so a refresh, a deselect or a failed plan request never leaves
            // the previous board's files on screen (code review, 2026-09-27, about the list it
            // replaced).
            // ###########################################################################################
            this.ShowFiles(plan, row);

            // Never carried from one board to the next - see the markup.
            check.IsChecked = false;
            check.IsEnabled = plan is not null && plan.CanPublish;

            // The name, or - with nothing chosen - what to do, in the Review screen's placeholder look.
            if (this.FindControl<TextBlock>("SelectedSystemText") is TextBlock title)
            {
                title.Text = row is null ? "Select a system" : ProductionDisplay.SystemLine(row);
                title.FontWeight = row is null ? FontWeight.Normal : FontWeight.SemiBold;
                title.Opacity = row is null ? 0.7 : 1;
            }

            if (this.FindControl<StackPanel>("ActionPanel") is StackPanel actions)
                actions.IsVisible = row is not null;

            this.SetText("PlanSummaryText", plan is null ? string.Empty : ProductionDisplay.PlanSummary(plan));

            // The two-person approval for a shared-file change, and a button that says what
            // pressing it does - both from the server's ApprovalStatus.
            string? approvalLine = ApprovalWording.StatusLine(plan?.Approval);
            this.SetText("ApprovalStatusText", approvalLine ?? string.Empty);

            if (this.FindControl<TextBlock>("ApprovalStatusText") is TextBlock approvalText)
                approvalText.IsVisible = approvalLine is not null;

            if (this.FindControl<Button>("PublishButton") is Button publish)
                publish.Content = ApprovalWording.ApproveButton(plan?.Approval, "stable");

            if (plan is not null)
            {
                BetaView.ShowCarrying(carrying, plan);

                foreach (ReviewFindingView problem in plan.Problems)
                {
                    attention.Children.Add(new TextBlock
                    {
                        Text = problem.Message,
                        TextWrapping = TextWrapping.Wrap,
                        Foreground = Brushes.IndianRed
                    });
                }

                // What publishing REMOVES from production, FIRST and in red: the least visible change
                // and the one that cannot be undone (2026-09-25). Said even when it is nothing, so
                // the maintainer knows it was checked.
                string? removals = FileRemovalWording.Headline(plan.Removals, "the stable data");

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

                    attention.Children.Add(headline);

                    foreach (string path in plan.Removals.Files)
                    {
                        attention.Children.Add(new TextBlock
                        {
                            Text = FileRemovalWording.Line(path),
                            TextWrapping = TextWrapping.Wrap,
                            Foreground = Brushes.IndianRed,
                            FontWeight = FontWeight.SemiBold
                        });
                    }
                }

                this.ShowMessage(plan.Refusal, isError: plan.Refusal is not null);
            }

            this.UpdateButton();
        }

        // ###########################################################################################
        // The tree, or nothing. Its counts are PlanSummaryText's line above it, so it does not say
        // them twice. A file is read from where the entry says - BETA's copy, or production's for
        // one the publish removes - and opened under this window's "please wait" (FileTreeFiles).
        // ###########################################################################################
        private void ShowFiles(ProductionPlanView? plan, ProductionSystemRow? row)
        {
            if (this.FindControl<FileTreeView>("FileTree") is not FileTreeView tree)
                return;

            tree.ShowSummary = false;

            if (plan is null || row is null)
            {
                tree.Clear();
                tree.IsVisible = false;
                return;
            }

            ReviewApiClient? client = this.thisClient;
            ReviewSession? session = this.thisSession;

            tree.Files = client is null ? null : new FileTreeFiles(this, client, session, submissionId: null, plan.BetaDataUrl, plan.ProductionDataUrl);
            // With each file's size, from the plan (2026-10-04).
            IReadOnlyDictionary<string, long>? sizes = plan.FileSizes;

            tree.Show(SystemFileEntries.WithSizes(
                SystemFileEntries.ForPromotion(plan.Files, plan.Removals?.Files, plan.UnchangedFiles),
                entry => sizes is not null && sizes.TryGetValue(entry.Path, out long size) ? size : null));
            tree.IsVisible = true;
        }

        // The tree on screen, for tests.
        internal FileTreeView FileTreeForTests => this.FindControl<FileTreeView>("FileTree")!;

        private void OnCheckedInBetaChanged(object? sender, Avalonia.Interactivity.RoutedEventArgs e) =>
            this.UpdateButton();

        private void UpdateButton()
        {
            var button = this.FindControl<Button>("PublishButton");
            var check = this.FindControl<CheckBox>("CheckedInBetaCheckBox");

            if (button is not null)
                button.IsEnabled = ProductionDisplay.CanPress(this.thisPlan, check?.IsChecked == true);

            // ###########################################################################################
            // Pushing back needs only a system on screen - NOT the tick, and not the server's
            // canPublish. The tick says "I checked this and it is right", which is the opposite of
            // what this button does, and a board whose publish is blocked (a shared file awaiting
            // the administrator, say) is exactly one a maintainer may want to push back.
            // ###########################################################################################
            if (this.FindControl<Button>("RollBackButton") is Button rollBack)
                rollBack.IsEnabled = this.thisPlan is not null;

            // Reject is the same roll back (2026-09-28), so the same rule.
            if (this.FindControl<Button>("RejectButton") is Button reject)
                reject.IsEnabled = this.thisPlan is not null;
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
            this.ShowMessage(null, isError: false);

            ReviewApiClient client = this.thisClient;
            ReviewSession session = this.thisSession;
            string waiting = ProductionDisplay.PublishingWait(plan.SystemId);

            string? outcome = null;
            bool failed = true;

            // The window is held for the publish AND the lists read after it, so nothing on screen
            // can be pressed against a state that is about to change.
            await BusyOverlay.HoldAsync(this, waiting, async () =>
            {
                // The removal list sent back is the one on screen; the server refuses if it would now
                // remove anything else.
                ReviewApiResult<ProductionPublishResult> result = await ServerWait.CallAsync(
                    this,
                    waiting,
                    token => client.PublishToProductionAsync(session, plan.SystemId, plan.BetaContentHash, plan.Removals?.Files ?? [], token));

                // No answer in two minutes: a publish carries on regardless, so the BETA list says
                // whether it landed - a system published to production leaves it.
                if (result.Failure == ReviewApiFailure.TimedOut)
                {
                    bool? stillInBeta = await this.IsStillInBetaAsync(client, session, plan.SystemId);

                    outcome = MaintainerWaitWording.PublishAfterTimeout(plan.SystemId, stillInBeta);
                    failed = stillInBeta != false;
                    await ServerWait.RunAsync(this, MaintainerWaitWording.ReadingQueue, this.AfterChangeAsync);
                    return;
                }

                if (!result.IsOk)
                {
                    outcome = result.Message;
                    return;
                }

                await ServerWait.RunAsync(this, MaintainerWaitWording.ReadingQueue, this.AfterChangeAsync);

                // The first of two approvals: recorded, nothing copied.
                outcome = result.Value!.IsAwaitingApproval
                    ? ApprovalWording.Recorded(result.Value.WaitingFor, "the stable source")
                    : $"{plan.SystemId} is published to the stable source ({result.Value.FilesCopied} file(s) copied)." +
                      FileRemovalWording.Done(result.Value.RemovedFiles) +
                      " Everyone gets it the next time their data updates.";
                failed = false;
            });

            this.ShowMessage(outcome ?? "The publish did not complete.", isError: failed);

            if (failed)
                this.UpdateButton();
        }

        // ###########################################################################################
        // After a timeout: is the system still waiting in BETA? A publish to production and a
        // push-back both take it off the list when they land. Null when the list could not be read.
        // ###########################################################################################
        private async Task<bool?> IsStillInBetaAsync(ReviewApiClient client, ReviewSession session, string systemId)
        {
            ReviewApiResult<ProductionListResponse> list = await ServerWait.CallAsync(
                this, WaitWording.Checking, token => client.GetProductionListAsync(session, token));

            return list.IsOk
                ? list.Value!.Systems.Any(row => string.Equals(row.SystemId, systemId, StringComparison.Ordinal))
                : null;
        }

        // ###########################################################################################
        // *** PUSHING A BOARD BACK TO THE QUEUE (owner decision, 2026-09-27). *** The plan is
        // fetched first and shown in the confirmation, because a rollback takes back EVERY
        // submission merged since the last promotion and the maintainer has to see whose work that
        // is before deciding - see RollBackBetaWindow.
        //
        // Deliberately NOT gated on the "I checked this in BETA" tick: that box says the data is
        // right, which is the opposite of what pushing back means.
        //
        // *** "REJECT" IS THE SAME, with the submissions rejected (owner request, 2026-09-28). ***
        // One path for both, so what a rejection does to BETA cannot drift from a push-back.
        // ###########################################################################################
        private async void OnRollBackClick(object? sender, Avalonia.Interactivity.RoutedEventArgs e) =>
            await this.TakeOutOfBetaAsync(reject: false);

        private async void OnRejectClick(object? sender, Avalonia.Interactivity.RoutedEventArgs e) =>
            await this.TakeOutOfBetaAsync(reject: true);

        private async Task TakeOutOfBetaAsync(bool reject)
        {
            ProductionPlanView? shown = this.thisPlan;

            if (shown is null || this.thisClient is null || this.thisSession is null ||
                TopLevel.GetTopLevel(this) is not Window owner)
            {
                return;
            }

            // Both off for the duration, so neither can send a second one.
            foreach (string name in new[] { "RollBackButton", "RejectButton" })
            {
                if (this.FindControl<Button>(name) is Button button)
                    button.IsEnabled = false;
            }

            this.ShowMessage(null, isError: false);

            ReviewApiClient client = this.thisClient;
            ReviewSession session = this.thisSession;

            ReviewApiResult<BetaRollbackPlanView> planned = await ServerWait.CallAsync(
                this,
                MaintainerWaitWording.ReadingRollbackPlan(shown.SystemId),
                token => client.GetBetaRollbackPlanAsync(session, shown.SystemId, token));

            if (!planned.IsOk)
            {
                this.ShowMessage(planned.Message, isError: true);
                this.UpdateButton();
                return;
            }

            var confirm = new RollBackBetaWindow();
            confirm.Initialize(planned.Value!, reject);

            await confirm.ShowDialog(owner);

            if (!confirm.WasConfirmed)
            {
                this.ShowMessage(null, isError: false);
                this.UpdateButton();
                return;
            }

            string waiting = reject ? ProductionDisplay.RejectingWait(shown.SystemId) : ProductionDisplay.PushingBackWait(shown.SystemId);
            string? outcome = null;
            bool failed = true;

            // Held until it is done, the lists included (owner request, 2026-09-27).
            await BusyOverlay.HoldAsync(this, waiting, async () =>
            {
                ReviewApiResult<BetaRollbackResult> result = await ServerWait.CallAsync(
                    this,
                    waiting,
                    token => client.RollBackBetaAsync(session, shown.SystemId, confirm.Comment, reject, token));

                // No answer in two minutes: whether the system left the BETA list says whether it landed.
                if (result.Failure == ReviewApiFailure.TimedOut)
                {
                    bool? stillInBeta = await this.IsStillInBetaAsync(client, session, shown.SystemId);

                    outcome = reject
                        ? MaintainerWaitWording.RejectAfterTimeout(shown.SystemId, stillInBeta)
                        : MaintainerWaitWording.PushBackAfterTimeout(shown.SystemId, stillInBeta);
                    failed = stillInBeta != false;
                    await ServerWait.RunAsync(this, MaintainerWaitWording.ReadingQueue, this.AfterChangeAsync);
                    return;
                }

                if (!result.IsOk)
                {
                    outcome = result.Message;
                    return;
                }

                // The system leaves the list (its BETA state is level with production, or gone) and
                // its submissions are back in the queue, so the lists are read again rather than patched.
                await ServerWait.RunAsync(this, MaintainerWaitWording.ReadingQueue, this.AfterChangeAsync);

                // Asked to reject, an older server pushes back instead - said as it is.
                if (reject && !result.Value!.Rejected)
                {
                    outcome = ProductionDisplay.PushedBackInsteadOfRejected(result.Value);
                    return;
                }

                outcome = reject ? ProductionDisplay.Rejected(result.Value!) : ProductionDisplay.RolledBack(result.Value!);
                failed = false;
            });

            this.ShowMessage(outcome ?? (reject ? "The rejection did not complete." : "The push-back did not complete."), isError: failed);

            if (failed)
                this.UpdateButton();
        }

        private async Task AfterChangeAsync()
        {
            if (this.AfterChange is not null)
                await this.AfterChange();
        }

        private void SetText(string name, string text)
        {
            if (this.FindControl<TextBlock>(name) is TextBlock block)
                block.Text = text;
        }
    }
}
