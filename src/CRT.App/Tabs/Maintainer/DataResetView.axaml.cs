using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Avalonia.Controls;
using Avalonia.Markup.Xaml;
using Handlers.DataHandling;
using Handlers.MaintainerHandling;

namespace CRT
{
    // ###########################################################################################
    // "Reset contribution data" on the Account screen (owner request, 2026-10-04). See the markup for
    // what it is; the server's DataResetFlow for what a reset deletes and keeps.
    //
    // *** COUNTS, CONFIRM, RESET - in that order. *** The counts are read when this item is chosen;
    // the button sends back their fingerprint, so the server deletes nothing that arrived after they
    // were shown (it answers Conflict, and the counts are read again). After a reset the counts are
    // read again - after a timeout too, when they are what says whether it went through
    // (DataResetWording.HasLanded), the tab then reading its lists again as after any reset.
    //
    // NOTHING IS READ ON THE MINUTE CHECK: the counts are read on choosing this item and after the
    // button, never in the background.
    // ###########################################################################################
    public partial class DataResetView : UserControl
    {
        private ReviewApiClient? thisClient;
        private ReviewSession? thisSession;
        private DataResetPlanAnswer? thisPlan;

        public DataResetView()
        {
            this.InitializeComponent();

            this.SetText("HeadingText", DataResetWording.Heading);
            this.SetText("ExplanationText", DataResetWording.Explanation);
            this.SetText("KeptText", DataResetWording.Kept);
            this.SetText("ConfirmPromptText", DataResetWording.ConfirmPrompt);

            if (this.FindControl<Button>("ResetButton") is Button button)
                button.Content = DataResetWording.ResetButton;

            // The Text PROPERTY, not TextChanged: that event is raised a dispatcher turn later, and
            // the button must follow the box exactly - never on for a word no longer there.
            if (this.FindControl<TextBox>("ConfirmBox") is TextBox box)
            {
                box.PropertyChanged += (_, e) =>
                {
                    if (e.Property == TextBox.TextProperty)
                        this.UpdateButton();
                };
            }
        }

        private void InitializeComponent() => AvaloniaXamlLoader.Load(this);

        // After a reset went through: the tab reads its queue, BETA list and boards again.
        public Func<Task>? AfterReset { get; set; }

        public void Initialize(ReviewApiClient? client, ReviewSession? session)
        {
            this.thisClient = client;
            this.thisSession = session;
        }

        // Signed out: nothing of the previous account's counts or result left on screen - nor the
        // word typed (code review, 2026-11-02: the next administrator on this computer got new counts
        // with the button already on).
        public void Clear()
        {
            this.EmptyConfirmBox();
            this.ShowPlan(null);
            this.ShowDone(null);
            this.ShowMessage(null, isError: false);
        }

        private void EmptyConfirmBox()
        {
            if (this.FindControl<TextBox>("ConfirmBox") is TextBox box)
                box.Text = string.Empty;
        }

        // ###########################################################################################
        // Reads the counts, under CRT's overlay. Counts that cannot be read leave the button off, with
        // the reason - a reset is never offered on counts nobody has seen. And never on the word typed
        // for the counts before these: it is typed again for what is shown now (code review, 2026-11-02).
        // ###########################################################################################
        public async Task LoadAsync()
        {
            this.EmptyConfirmBox();

            if (this.thisClient is null || this.thisSession is null)
                return;

            ReviewApiClient client = this.thisClient;
            ReviewSession session = this.thisSession;

            ReviewApiResult<DataResetPlanAnswer> result = await ServerWait.CallAsync(
                this,
                MaintainerWaitWording.ReadingResetCounts,
                token => client.GetDataResetPlanAsync(session, token));

            if (!result.IsOk)
            {
                this.ShowPlan(null);
                this.ShowMessage(result.Message, isError: true);
                return;
            }

            this.ShowPlan(result.Value);
        }

        private async void OnResetClick(object? sender, Avalonia.Interactivity.RoutedEventArgs e) =>
            await HandlerGuard.RunAsync(this.ResetAsync, message => this.ShowMessage(message, isError: true));

        // ###########################################################################################
        // Sends the reset with the shown counts' fingerprint. Whatever happens, the counts are read
        // again afterwards and the confirmation word is cleared - a second reset is typed again.
        // ###########################################################################################
        internal async Task ResetAsync()
        {
            if (this.thisClient is null || this.thisSession is null ||
                !DataResetWording.CanReset(this.thisPlan, this.FindControl<TextBox>("ConfirmBox")?.Text))
            {
                return;
            }

            ReviewApiClient client = this.thisClient;
            ReviewSession session = this.thisSession;
            DataResetPlanAnswer shown = this.thisPlan!;
            string fingerprint = shown.Fingerprint;

            this.ShowDone(null);
            this.ShowMessage(null, isError: false);

            ReviewApiResult<DataResetAnswer> result = await ServerWait.CallAsync(
                this,
                MaintainerWaitWording.ResettingData,
                token => client.ResetDataAsync(session, fingerprint, token));

            await this.LoadAsync();

            if (result.IsOk)
            {
                this.ShowDone(result.Value);

                if (this.AfterReset is not null)
                    await this.AfterReset();

                return;
            }

            // ###########################################################################################
            // *** NO ANSWER IS JUDGED FROM THE COUNTS JUST READ (code review, 2026-11-02). *** It said "if
            // they are all zero, it did", which they never are - and left the queue and the boards as they
            // were until the next minute check, though every submission might have gone.
            // ###########################################################################################
            if (result.Failure == ReviewApiFailure.TimedOut)
            {
                bool? landed = DataResetWording.HasLanded(shown, this.thisPlan);

                this.ShowMessage(DataResetWording.AfterTimeout(landed), isError: landed != true);

                if (landed == true && this.AfterReset is not null)
                    await this.AfterReset();

                return;
            }

            this.ShowMessage(result.Failure == ReviewApiFailure.Conflict ? DataResetWording.Changed : result.Message, isError: true);
        }

        private void ShowPlan(DataResetPlanAnswer? plan)
        {
            this.thisPlan = plan;

            if (this.FindControl<StackPanel>("CountLines") is StackPanel lines)
            {
                lines.Children.Clear();

                foreach (IReadOnlyList<ReviewNoteRun> line in plan is null ? [] : DataResetWording.Lines(plan))
                {
                    var block = new TextBlock { TextWrapping = Avalonia.Media.TextWrapping.Wrap };
                    TabMaintainer.ShowCounts(block, line);
                    lines.Children.Add(block);
                }
            }

            if (this.FindControl<TextBlock>("SwitchedOffText") is TextBlock off)
            {
                off.Text = plan?.NotEnabledBecause ?? string.Empty;
                off.IsVisible = plan is { IsEnabled: false };
            }

            if (this.FindControl<TextBox>("ConfirmBox") is TextBox box)
                box.IsEnabled = plan is { IsEnabled: true };

            this.UpdateButton();
        }

        private void ShowDone(DataResetAnswer? answer)
        {
            if (this.FindControl<TextBlock>("DoneText") is not TextBlock done)
                return;

            done.IsVisible = answer is not null;
            TabMaintainer.ShowCounts(done, answer is null ? [] : DataResetWording.Done(answer));
        }

        private void UpdateButton()
        {
            if (this.FindControl<Button>("ResetButton") is Button button)
                button.IsEnabled = DataResetWording.CanReset(this.thisPlan, this.FindControl<TextBox>("ConfirmBox")?.Text);
        }

        private void ShowMessage(string? message, bool isError) =>
            WindowMessage.Show(this.FindControl<TextBlock>("MessageText"), message, isError);

        private void SetText(string name, string text)
        {
            if (this.FindControl<TextBlock>(name) is TextBlock block)
                block.Text = text;
        }

        // ---- For tests ----------------------------------------------------------------------------

        internal IReadOnlyList<string> CountLinesForTests() =>
            this.FindControl<StackPanel>("CountLines")?.Children.OfType<TextBlock>().Select(TabMaintainer.TextOf).ToList() ?? [];

        internal Button ResetButtonForTests => this.FindControl<Button>("ResetButton")!;

        internal TextBox ConfirmBoxForTests => this.FindControl<TextBox>("ConfirmBox")!;

        internal string? SwitchedOffForTests =>
            this.FindControl<TextBlock>("SwitchedOffText") is { IsVisible: true } off ? off.Text : null;

        internal string? DoneForTests =>
            this.FindControl<TextBlock>("DoneText") is { IsVisible: true } done ? TabMaintainer.TextOf(done) : null;

        internal string? MessageForTests =>
            this.FindControl<TextBlock>("MessageText") is { IsVisible: true } message ? message.Text : null;
    }
}
