using System;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Markup.Xaml;
using Avalonia.Media;
using Handlers.MaintainerHandling;
using Handlers.DataHandling;

namespace CRT
{
    // ###########################################################################################
    // CONFIRMING A BETA ROLLBACK (owner decision, 2026-09-27) - the BETA screen's "push back
    // to queue".
    //
    // *** THE POINT OF THE DIALOG IS THE NAMES. *** A rollback is per SYSTEM, so it takes back
    // EVERY submission merged since the last promotion; one contributor's work cannot be picked
    // out (ProductionPromotionPlan's header explains why in the other direction). Listing them
    // here, before anything happens, is what stops a maintainer discarding two other people's
    // accepted work believing they are returning one.
    //
    // *** ENTER AND ESCAPE BOTH CANCEL, on the Tunnel route *** - the same rule and the same reason
    // as DeleteWorkbookWindow and UnsavedTableEditsWindow: a focused Button consumes Enter on the
    // bubbling route, so a reflexive keypress would confirm an operation that takes a board out of
    // BETA. Escape is a plain cancel; Enter must not be a shortcut for this one.
    //
    // The comment is required here AND on the server: the button stays off until something is
    // typed, and BetaRollbackFlow refuses a blank one regardless.
    //
    // *** ALSO THE CONFIRMATION FOR "REJECT" (owner request, 2026-09-28). *** The same rollback with
    // its submissions rejected, so the same names and the same required reason - only the words
    // differ (ProductionDisplay.Reject*).
    // ###########################################################################################
    public partial class RollBackBetaWindow : Window
    {
        public RollBackBetaWindow()
        {
            this.InitializeComponent();

            this.AddHandler(KeyDownEvent, this.OnWindowPreviewKeyDown, RoutingStrategies.Tunnel);
        }

        private void InitializeComponent() => AvaloniaXamlLoader.Load(this);

        // True when the maintainer confirmed; the comment is then what they typed.
        public bool WasConfirmed { get; private set; }

        public string Comment { get; private set; } = string.Empty;

        // `reject`: Beta > Prod's "Reject" rather than "Push back to queue".
        public void Initialize(BetaRollbackPlanView plan, bool reject = false)
        {
            ArgumentNullException.ThrowIfNull(plan);

            this.Title = reject ? "Reject" : "Push back to queue";
            this.SetText("HeadlineText", reject ? ProductionDisplay.RejectHeadline(plan) : ProductionDisplay.RollBackHeadline(plan));
            this.SetText("ExplanationText", reject ? ProductionDisplay.RejectExplanation(plan) : ProductionDisplay.RollBackExplanation(plan));

            if (this.FindControl<Button>("ConfirmButton") is Button confirm)
                confirm.Content = reject ? ProductionDisplay.RejectConfirmButton(plan) : ProductionDisplay.RollBackConfirmButton(plan);

            if (this.FindControl<StackPanel>("ReturningPanel") is not StackPanel returning)
                return;

            returning.Children.Clear();

            DateTimeOffset now = DateTimeOffset.UtcNow;

            foreach (CarriedSubmission submission in plan.Returning)
            {
                returning.Children.Add(new TextBlock
                {
                    Text = ProductionDisplay.CarryingLine(submission, now),
                    TextWrapping = TextWrapping.Wrap
                });
            }

            returning.IsVisible = plan.Returning.Count > 0;

            if (this.FindControl<StackPanel>("SharedPanel") is not StackPanel shared)
                return;

            shared.Children.Clear();

            if (ProductionDisplay.RollBackSharedFiles(plan) is string sentence)
            {
                shared.Children.Add(new TextBlock { Text = sentence, TextWrapping = TextWrapping.Wrap, FontWeight = FontWeight.SemiBold });

                foreach (string path in ProductionDisplay.RollBackSharedPaths(plan))
                    shared.Children.Add(new TextBlock { Text = path, TextWrapping = TextWrapping.Wrap });
            }

            shared.IsVisible = shared.Children.Count > 0;
        }

        private void SetText(string name, string text)
        {
            if (this.FindControl<TextBlock>(name) is TextBlock block)
                block.Text = text;
        }

        private void OnCommentChanged(object? sender, TextChangedEventArgs e)
        {
            if (this.FindControl<Button>("ConfirmButton") is Button confirm)
                confirm.IsEnabled = !string.IsNullOrWhiteSpace(this.FindControl<TextBox>("CommentTextBox")?.Text);
        }

        private void OnConfirmClick(object? sender, RoutedEventArgs e)
        {
            string comment = (this.FindControl<TextBox>("CommentTextBox")?.Text ?? string.Empty).Trim();

            if (comment.Length == 0)
                return;

            this.WasConfirmed = true;
            this.Comment = comment;
            this.Close();
        }

        private void OnCancelClick(object? sender, RoutedEventArgs e) => this.Close();

        // ###########################################################################################
        // Enter and Escape both CANCEL. Enter is NOT handled while the comment box has focus, or
        // the maintainer could not type a second line of explanation - the box accepts returns.
        // ###########################################################################################
        private void OnWindowPreviewKeyDown(object? sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
            {
                e.Handled = true;
                this.Close();
                return;
            }

            if (e.Key != Key.Enter)
                return;

            if (this.FindControl<TextBox>("CommentTextBox")?.IsFocused == true)
                return;

            e.Handled = true;
            this.Close();
        }
    }
}
