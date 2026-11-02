using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Layout;
using Avalonia.Markup.Xaml;
using Avalonia.Media;
using Handlers.DataHandling;
using Handlers.MaintainerHandling;
using Handlers.Theming;

namespace CRT
{
    // ###########################################################################################
    // Account > "Oscilloscope models" (CRT 4.0 step 5) - see the markup for what it is;
    // OscilloscopeModelsDisplay words it and CRT.Server's OscilloscopeModelFlows decides.
    //
    // Read when the entry is chosen, when Account is shown again with it chosen, on Refresh and after
    // each decision. On the minute check while on screen, and off screen only when the count the queue
    // answered differs from what is on show (FollowWaitingCountAsync, code review 2026-11-02: the badge
    // said 1 beside "No oscilloscope model is waiting"), and then quietly. A decision is made under CRT's overlay; one that got no
    // answer may have happened, so the list is read again and whether the model is still in it is
    // said (a model decided is no longer in it - MaintainerWaitWording.OscilloscopeModelAfterTimeout). The
    // cards are made again from each read: nothing is typed into a card but its reason - kept for
    // every model still listed (code review, 2026-10-29: a reason half typed on one card was gone once
    // another was decided) - and no quiet read lands under a reason being typed. A model a publish put
    // into BETA only offers Publish alone: it cannot be rejected (the server's rule).
    // ###########################################################################################
    public partial class OscilloscopeModelsView : UserControl
    {
        private ReviewApiClient? thisClient;
        private ReviewSession? thisSession;
        private bool thisIsDeciding;

        // The reason box of each card, by model id.
        private readonly Dictionary<long, TextBox> thisReasons = new();

        // The models on show, as last read - null until a read went well.
        private IReadOnlyList<OscilloscopeModelQueueEntry>? thisShown;

        public OscilloscopeModelsView()
        {
            this.InitializeComponent();

            this.GetControl<TextBlock>("HeadingText").Text = OscilloscopeModelsDisplay.Heading;
            this.GetControl<TextBlock>("ExplanationText").Text = OscilloscopeModelsDisplay.Explanation;
        }

        private void InitializeComponent() => AvaloniaXamlLoader.Load(this);

        public void Initialize(ReviewApiClient? client, ReviewSession? session)
        {
            this.thisClient = client;
            this.thisSession = session;
        }

        // Signed out: nothing of the previous account's list left.
        public void Clear()
        {
            this.thisReasons.Clear();
            this.thisShown = null;
            this.GetControl<StackPanel>("ModelsPanel").Children.Clear();
            this.ShowMessage(null, isError: false);
        }

        // How many models the list read shows waiting - the tab's badges follow it at once (code review,
        // 2026-10-29: after the last decision they still counted it until the next minute check).
        internal Action<int>? ModelsWaitingRead { get; set; }

        // Answers a read without a server - for tests.
        internal Func<Task<ReviewApiResult<OscilloscopeModelQueueAnswer>>>? ReadOverrideForTests { get; set; }

        // Answers a decision without a server - for tests: the id, the reason (null for a publish).
        internal Func<long, string?, Task<ReviewApiResult<OscilloscopeModelDecisionAnswer>>>? DecideOverrideForTests { get; set; }

        // ###########################################################################################
        // Reads the models waiting, under CRT's overlay. An answer that cannot be read leaves what is
        // shown, with the reason.
        // ###########################################################################################
        public Task LoadAsync() => this.ReadAndShowAsync(MaintainerWaitWording.ReadingOscilloscopeModels);

        // The list read and shown, or null when it could not be - the reason then said. A sign-out
        // or another sign-in meanwhile wins: nothing of this answer goes on show.
        private async Task<OscilloscopeModelQueueAnswer?> ReadAndShowAsync(string waiting)
        {
            ReviewSession? asked = this.thisSession;

            if (this.Reader() is not { } read)
                return null;

            ReviewApiResult<OscilloscopeModelQueueAnswer> result = await ServerWait.CallAsync(this, waiting, read);

            if (!ReferenceEquals(asked, this.thisSession))
                return null;

            if (!result.IsOk)
            {
                this.ShowMessage(result.Message, isError: true);
                return null;
            }

            this.ShowModels(result.Value!);
            return result.Value;
        }

        // The one way the list is read - a test's answer, or the server's. Null when signed out.
        private Func<CancellationToken, Task<ReviewApiResult<OscilloscopeModelQueueAnswer>>>? Reader() =>
            this.ReadOverrideForTests is { } answer
                ? _ => answer()
                : this.thisClient is ReviewApiClient client && this.thisSession is ReviewSession session
                    ? token => client.GetOscilloscopeModelsAsync(session, token)
                    : null;

        // ###########################################################################################
        // *** THE MODELS ON SHOW FOLLOW THE COUNT THE QUEUE ANSWERED (code review, 2026-11-02). *** The
        // badge moves every minute with that answer; this list read only when asked, so it could say
        // "No oscilloscope model is waiting" beside a badge saying 1.
        //
        // *** AND ON SCREEN THEY ARE READ AT EVERY FULL CHECK, whatever the count says (the same review). ***
        // A model re-sent under its name replaces the pending copy, so the count stands still: the card kept
        // the old commands and the old id, and its Publish answered "decided meanwhile". The other lists are
        // read that way too. Off screen the count is all that is asked - showing the list reads it.
        //
        // QUIET: nobody pressed anything, so no overlay and no word if it fails - the next check asks
        // again. Never while a decision is on its way (that reads by itself) or while a reason is being
        // typed: the cards are made again from each read, which keeps the words but takes the box from
        // under the caret. The next check asks again then too. A read with nothing new makes no card
        // again, for the same reason and for the press going down on a button.
        // ###########################################################################################
        internal async Task FollowWaitingCountAsync(int waiting, bool onScreen)
        {
            if ((!onScreen && this.thisShown?.Count == waiting) || this.thisIsDeciding || this.IsTypingReason)
                return;

            ReviewSession? asked = this.thisSession;

            if (this.Reader() is not { } read)
                return;

            ReviewApiResult<OscilloscopeModelQueueAnswer> result = await read(CancellationToken.None);

            if (!result.IsOk || !ReferenceEquals(asked, this.thisSession) || this.thisIsDeciding || this.IsTypingReason)
                return;

            if (OscilloscopeModelsDisplay.SameModels(this.thisShown, result.Value!.Entries))
                this.ModelsWaitingRead?.Invoke(result.Value.Entries.Count);
            else
                this.ShowModels(result.Value);
        }

        private bool IsTypingReason => this.thisReasons.Values.Any(box => box.IsKeyboardFocusWithin);

        private async void OnRefreshClick(object? sender, Avalonia.Interactivity.RoutedEventArgs e)
        {
            this.ShowMessage(null, isError: false);
            await HandlerGuard.RunAsync(this.LoadAsync, message => this.ShowMessage(message, isError: true));
        }

        // ###########################################################################################
        // One card per model - or the line saying none waits, never an empty box.
        // ###########################################################################################
        internal void ShowModels(OscilloscopeModelQueueAnswer answer)
        {
            StackPanel panel = this.GetControl<StackPanel>("ModelsPanel");

            // What is typed in each reason box, to put back on its model's new card.
            Dictionary<long, string> typed = this.thisReasons
                .Where(reason => !string.IsNullOrEmpty(reason.Value.Text))
                .ToDictionary(reason => reason.Key, reason => reason.Value.Text!);

            panel.Children.Clear();
            this.thisReasons.Clear();
            this.thisShown = answer.Entries;

            this.ModelsWaitingRead?.Invoke(answer.Entries.Count);

            if (answer.Entries.Count == 0)
            {
                panel.Children.Add(new TextBlock { Text = OscilloscopeModelsDisplay.Empty, Opacity = 0.8 });
                return;
            }

            foreach (OscilloscopeModelQueueEntry entry in answer.Entries)
                panel.Children.Add(this.Card(entry, typed.GetValueOrDefault(entry.Id)));
        }

        private Control Card(OscilloscopeModelQueueEntry entry, string? typedReason)
        {
            var body = new StackPanel { Spacing = 4 };

            body.Children.Add(new TextBlock
            {
                Text = OscilloscopeModelsDisplay.Title(entry),
                FontWeight = FontWeight.SemiBold,
                FontSize = 13,
            });

            var kind = new TextBlock { Text = OscilloscopeModelsDisplay.Kind(entry), TextWrapping = TextWrapping.Wrap };

            if (entry.InBeta)
                kind.Bind(TextBlock.ForegroundProperty, kind.GetResourceObservable("Text_Warning_Fg"));
            else
                kind.Opacity = 0.8;

            body.Children.Add(kind);

            var sentBy = new TextBlock { TextWrapping = TextWrapping.Wrap };
            TabMaintainer.ShowCounts(sentBy, OscilloscopeModelsDisplay.SentBy(entry));
            body.Children.Add(sentBy);

            // Whether that is a known person, and the suite as what the sender's CRT reported: nothing
            // a stranger can type is drawn as good news (code review, 2026-11-02 - Security). It was
            // green.
            body.Children.Add(OscilloscopeModelsView.Said(OscilloscopeModelsDisplay.SenderLine(entry), warn: !entry.SignedIn));
            body.Children.Add(OscilloscopeModelsView.Said(OscilloscopeModelsDisplay.SuiteLine(entry), warn: !entry.FullTestSuitePassed));

            body.Children.Add(new TextBlock
            {
                Text = OscilloscopeModelsDisplay.Message(entry),
                FontStyle = FontStyle.Italic,
                TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(0, 4, 0, 4),
            });

            body.Children.Add(OscilloscopeModelsView.ValuesGrid(entry));

            var publish = new Button { Content = "Publish", Classes = { "Accept" } };
            var buttons = new StackPanel { Orientation = Orientation.Horizontal, Spacing = ButtonRow.Gap, Margin = new Thickness(0, 6, 0, 0) };

            publish.Click += async (_, _) =>
                await HandlerGuard.RunAsync(() => this.DecideAsync(entry.Id, reason: null), message => this.ShowMessage(message, isError: true));
            buttons.Children.Add(publish);

            if (!entry.InBeta)
            {
                var reason = new TextBox
                {
                    PlaceholderText = OscilloscopeModelsDisplay.ReasonPlaceholder,
                    AcceptsReturn = true,
                    TextWrapping = TextWrapping.Wrap,
                    MinHeight = 48,
                    Margin = new Thickness(0, 8, 0, 0),
                    Text = typedReason,
                };

                this.thisReasons[entry.Id] = reason;
                body.Children.Add(reason);

                var reject = new Button { Content = "Reject", Classes = { "Danger" } };

                reject.Click += async (_, _) =>
                    await HandlerGuard.RunAsync(() => this.DecideAsync(entry.Id, reason.Text ?? string.Empty), message => this.ShowMessage(message, isError: true));
                buttons.Children.Add(reject);
            }

            body.Children.Add(buttons);

            var card = new Border
            {
                Child = body,
                Padding = new Thickness(12),
                BorderThickness = new Thickness(1),
                CornerRadius = new CornerRadius(4),
            };
            card.Bind(Border.BorderBrushProperty, card.GetResourceObservable("Panel_Border"));

            return card;
        }

        // A line of the card: in the warning colour when it is something to weigh, else plain and dimmed.
        private static TextBlock Said(string text, bool warn)
        {
            var line = new TextBlock { Text = text, TextWrapping = TextWrapping.Wrap };

            if (warn)
                line.Bind(TextBlock.ForegroundProperty, line.GetResourceObservable("Text_Warning_Fg"));
            else
                line.Opacity = 0.8;

            return line;
        }

        // Every value: its label, what was sent, and - for an update - what is published now, a
        // change in the warning colour.
        private static Grid ValuesGrid(OscilloscopeModelQueueEntry entry)
        {
            bool update = entry.PublishedValues is not null;
            var grid = new Grid { ColumnDefinitions = new ColumnDefinitions(update ? "Auto,12,*,12,*" : "Auto,12,*") };
            IReadOnlyList<ScopeModelValueRow> rows = OscilloscopeModelsDisplay.Rows(entry);
            int line = 0;

            void Add(Control control, int column)
            {
                while (grid.RowDefinitions.Count <= line)
                    grid.RowDefinitions.Add(new RowDefinition(GridLength.Auto));

                Grid.SetRow(control, line);
                Grid.SetColumn(control, column);
                grid.Children.Add(control);
            }

            Add(new TextBlock { Text = string.Empty }, 0);
            Add(new TextBlock { Text = OscilloscopeModelsDisplay.ValueColumn, FontWeight = FontWeight.SemiBold }, 2);

            if (update)
                Add(new TextBlock { Text = OscilloscopeModelsDisplay.PublishedColumn, FontWeight = FontWeight.SemiBold }, 4);

            foreach (ScopeModelValueRow row in rows)
            {
                line++;
                Add(new TextBlock { Text = row.Label, Opacity = 0.8, FontSize = 11 }, 0);

                var sent = new TextBlock { Text = row.Sent, FontFamily = new FontFamily("Consolas"), FontSize = 11, TextWrapping = TextWrapping.Wrap };

                if (row.Differs)
                {
                    sent.FontWeight = FontWeight.Bold;
                    sent.Bind(TextBlock.ForegroundProperty, sent.GetResourceObservable("Text_Warning_Fg"));
                }

                Add(sent, 2);

                if (update)
                    Add(new TextBlock { Text = row.Published ?? string.Empty, FontFamily = new FontFamily("Consolas"), FontSize = 11, Opacity = 0.8, TextWrapping = TextWrapping.Wrap }, 4);
            }

            return grid;
        }

        // ###########################################################################################
        // Publish (reason null) or Reject one model, under CRT's overlay; then the list read again,
        // the decision's own sentence kept on screen. A reject with no reason asks for one first.
        // ###########################################################################################
        internal async Task DecideAsync(long id, string? reason)
        {
            if (this.thisIsDeciding)
                return;

            if (reason is not null && reason.Trim().Length == 0)
            {
                this.ShowMessage(OscilloscopeModelsDisplay.NoReason, isError: true);
                return;
            }

            Func<CancellationToken, Task<ReviewApiResult<OscilloscopeModelDecisionAnswer>>>? decide =
                this.DecideOverrideForTests is { } answer
                    ? _ => answer(id, reason)
                    : this.thisClient is ReviewApiClient client && this.thisSession is ReviewSession session
                        ? reason is null
                            ? token => client.PublishOscilloscopeModelAsync(session, id, token)
                            : token => client.RejectOscilloscopeModelAsync(session, id, reason.Trim(), token)
                        : null;

            if (decide is null)
                return;

            this.thisIsDeciding = true;
            (string Text, bool IsError) said;
            bool timedOut;

            // Named now: a model decided is gone from the list read below.
            string name = this.thisShown?.FirstOrDefault(entry => entry.Id == id) is { } shown
                ? OscilloscopeModelsDisplay.Title(shown)
                : "The model";

            try
            {
                ReviewApiResult<OscilloscopeModelDecisionAnswer> result = await ServerWait.CallAsync(
                    this,
                    reason is null ? MaintainerWaitWording.PublishingOscilloscopeModel : MaintainerWaitWording.RejectingOscilloscopeModel,
                    decide);

                timedOut = result.Failure == ReviewApiFailure.TimedOut;

                said = !result.IsOk
                    ? (result.Message ?? "The server refused this.", true)
                    : result.Value!.Problem is string problem
                        ? ($"{result.Value.Message} {problem}", true)
                        : (result.Value.Message, false);
            }
            finally
            {
                this.thisIsDeciding = false;
            }

            OscilloscopeModelQueueAnswer? now = await this.ReadAndShowAsync(
                timedOut ? WaitWording.Checking : MaintainerWaitWording.ReadingOscilloscopeModels);

            // ###########################################################################################
            // *** NO ANSWER IS NEVER "TRY AGAIN" (code review, 2026-11-02; CLAUDE.md: "A timeout is never
            // it failed"). *** The sentence for a read said it, in red, also when the server had finished -
            // the list then showed the card gone with no word of what became of it, and a model still
            // listed answered a second press with "decided meanwhile". The list just read says which.
            // ###########################################################################################
            if (timedOut)
            {
                bool? stillListed = now?.Entries.Any(entry => entry.Id == id);

                said = (MaintainerWaitWording.OscilloscopeModelAfterTimeout(reason is null, name, stillListed), stillListed != false);
            }

            this.ShowMessage(said.Text, said.IsError);
        }

        private void ShowMessage(string? message, bool isError) =>
            WindowMessage.Show(this.FindControl<TextBlock>("MessageText"), message, isError);

        // ---- For tests ----------------------------------------------------------------------------

        internal string? MessageForTests =>
            this.FindControl<TextBlock>("MessageText") is { IsVisible: true } message ? message.Text : null;

        // Each card's lines, in order - the values grid's cells joined by " | ".
        internal IReadOnlyList<IReadOnlyList<string>> CardsForTests() =>
            this.GetControl<StackPanel>("ModelsPanel").Children
                .OfType<Border>()
                .Select(card => (IReadOnlyList<string>)((StackPanel)card.Child!).Children
                    .Select(child => child switch
                    {
                        TextBlock text => TabMaintainer.TextOf(text),
                        Grid grid => string.Join(" | ", grid.Children.OfType<TextBlock>().Select(cell => cell.Text)),
                        TextBox box => "[reason]",
                        StackPanel row => string.Join(" ", row.Children.OfType<Button>().Select(button => $"[{button.Content}:{string.Join(",", button.Classes)}]")),
                        _ => child.GetType().Name,
                    })
                    .ToList())
                .ToList();

        internal void TypeReasonForTests(long id, string reason) => this.thisReasons[id].Text = reason;

        internal string? ReasonForTests(long id) => this.thisReasons[id].Text;

        internal string? NoneWaitingLineForTests =>
            this.GetControl<StackPanel>("ModelsPanel").Children.OfType<TextBlock>().FirstOrDefault()?.Text;
    }
}
