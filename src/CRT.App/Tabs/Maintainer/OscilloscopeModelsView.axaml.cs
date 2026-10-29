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

namespace CRT
{
    // ###########################################################################################
    // Account > "Oscilloscope models" (CRT 4.0 step 5) - see the markup for what it is;
    // OscilloscopeModelsDisplay words it and CRT.Server's OscilloscopeModelFlows decides.
    //
    // Read when the entry is chosen, on Refresh, and after each decision - never on the minute check.
    // A decision is made under CRT's overlay; a publish that got no answer may have happened, so it
    // is said so and the list read again shows the truth (a model decided is no longer in it). The
    // cards are made again from each read: nothing is typed into a card but its reason - kept for
    // every model still listed (code review, 2026-10-29: a reason half typed on one card was gone once
    // another was decided) - and a read only follows a press elsewhere or a decision, never under a
    // press on a card. A model a publish put into BETA only offers Publish alone: it cannot be
    // rejected (the server's rule).
    // ###########################################################################################
    public partial class OscilloscopeModelsView : UserControl
    {
        private ReviewApiClient? thisClient;
        private ReviewSession? thisSession;
        private bool thisIsDeciding;

        // The reason box of each card, by model id.
        private readonly Dictionary<long, TextBox> thisReasons = new();

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
        public async Task LoadAsync()
        {
            Func<CancellationToken, Task<ReviewApiResult<OscilloscopeModelQueueAnswer>>>? read =
                this.ReadOverrideForTests is { } answer
                    ? _ => answer()
                    : this.thisClient is ReviewApiClient client && this.thisSession is ReviewSession session
                        ? token => client.GetOscilloscopeModelsAsync(session, token)
                        : null;

            if (read is null)
                return;

            ReviewApiResult<OscilloscopeModelQueueAnswer> result = await ServerWait.CallAsync(
                this, MaintainerWaitWording.ReadingOscilloscopeModels, read);

            if (!result.IsOk)
            {
                this.ShowMessage(result.Message, isError: true);
                return;
            }

            this.ShowModels(result.Value!);
        }

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

            var suite = new TextBlock { Text = OscilloscopeModelsDisplay.SuiteLine(entry), TextWrapping = TextWrapping.Wrap };
            suite.Bind(TextBlock.ForegroundProperty, suite.GetResourceObservable(entry.FullTestSuitePassed ? "Text_Success_Fg" : "Text_Warning_Fg"));
            body.Children.Add(suite);

            body.Children.Add(new TextBlock
            {
                Text = OscilloscopeModelsDisplay.Message(entry),
                FontStyle = FontStyle.Italic,
                TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(0, 4, 0, 4),
            });

            body.Children.Add(OscilloscopeModelsView.ValuesGrid(entry));

            var publish = new Button { Content = "Publish", Classes = { "Accept" } };
            var buttons = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 8, Margin = new Thickness(0, 6, 0, 0) };

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

            try
            {
                ReviewApiResult<OscilloscopeModelDecisionAnswer> result = await ServerWait.CallAsync(
                    this,
                    reason is null ? MaintainerWaitWording.PublishingOscilloscopeModel : MaintainerWaitWording.RejectingOscilloscopeModel,
                    decide);

                said = result.Failure == ReviewApiFailure.TimedOut
                    ? (WaitWording.NoAnswer, true)
                    : !result.IsOk
                        ? (result.Message ?? "The server refused this.", true)
                        : result.Value!.Problem is string problem
                            ? ($"{result.Value.Message} {problem}", true)
                            : (result.Value.Message, false);
            }
            finally
            {
                this.thisIsDeciding = false;
            }

            await this.LoadAsync();
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
