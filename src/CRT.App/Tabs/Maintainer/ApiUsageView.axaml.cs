using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Markup.Xaml;
using Avalonia.Media;
using Handlers.DataHandling;
using Handlers.MaintainerHandling;

namespace CRT
{
    // ###########################################################################################
    // "API usage" on the Account screen (owner request, 2026-10-04). See the markup for what it is; the
    // server's ApiUsageFlow for what is counted, and ApiUsageDisplay for the words and the order.
    //
    // Read when this item is chosen, when the days are changed and on Refresh - never on the minute
    // check, which this screen does not need.
    // ###########################################################################################
    public partial class ApiUsageView : UserControl
    {
        private ReviewApiClient? thisClient;
        private ReviewSession? thisSession;

        // Set while the box is filled in code, so that does not count as the user choosing.
        private bool thisFillingDays;

        public ApiUsageView()
        {
            this.InitializeComponent();

            this.SetText("HeadingText", ApiUsageDisplay.Heading);
            this.SetText("ExplanationText", ApiUsageDisplay.Explanation);

            if (this.FindControl<ComboBox>("DaysBox") is ComboBox days)
            {
                this.thisFillingDays = true;
                days.ItemsSource = ApiUsageDisplay.DayChoices.Select(ApiUsageDisplay.DayChoiceLabel).ToList();
                days.SelectedIndex = ApiUsageDisplay.DayChoices.ToList().IndexOf(ApiUsageDisplay.DefaultDays);
                this.thisFillingDays = false;
            }
        }

        private void InitializeComponent() => AvaloniaXamlLoader.Load(this);

        public void Initialize(ReviewApiClient? client, ReviewSession? session)
        {
            this.thisClient = client;
            this.thisSession = session;
        }

        // Signed out: nothing of the previous account's answer left on screen.
        public void Clear()
        {
            this.FindControl<StackPanel>("UsagePanel")?.Children.Clear();
            this.ShowMessage(null, isError: false);
        }

        // The days chosen in the box.
        internal int DaysChosen =>
            this.FindControl<ComboBox>("DaysBox")?.SelectedIndex is int index && index >= 0 && index < ApiUsageDisplay.DayChoices.Count
                ? ApiUsageDisplay.DayChoices[index]
                : ApiUsageDisplay.DefaultDays;

        // ###########################################################################################
        // Reads the usage, under CRT's overlay. An answer that cannot be read leaves the previous one
        // on screen, with the reason.
        // ###########################################################################################
        public async Task LoadAsync()
        {
            if (this.thisClient is null || this.thisSession is null)
                return;

            ReviewApiClient client = this.thisClient;
            ReviewSession session = this.thisSession;
            int days = this.DaysChosen;

            ReviewApiResult<ApiUsageAnswer> result = await ServerWait.CallAsync(
                this,
                MaintainerWaitWording.ReadingApiUsage,
                token => client.GetApiUsageAsync(session, days, token));

            if (!result.IsOk)
            {
                this.ShowMessage(result.Message, isError: true);
                return;
            }

            this.ShowMessage(null, isError: false);
            this.ShowUsage(result.Value!);
        }

        private async void OnDaysChanged(object? sender, SelectionChangedEventArgs e)
        {
            if (!this.thisFillingDays)
                await this.LoadAsync();
        }

        private async void OnRefreshClick(object? sender, Avalonia.Interactivity.RoutedEventArgs e) =>
            await this.LoadAsync();

        // ###########################################################################################
        // The launches first - who still runs what - then each group of routes: its heading, and per
        // route its name, its summary and a line per version, indented under it.
        // ###########################################################################################
        internal void ShowUsage(ApiUsageAnswer answer)
        {
            if (this.FindControl<StackPanel>("UsagePanel") is not StackPanel panel)
                return;

            panel.Children.Clear();

            panel.Children.Add(ApiUsageView.Heading(ApiUsageDisplay.InstallationsHeading, top: 0));

            foreach (IReadOnlyList<ReviewNoteRun> line in ApiUsageDisplay.InstallationLines(answer))
                panel.Children.Add(ApiUsageView.Line(line, indent: 12));

            foreach (ApiUsageGroup group in ApiUsageDisplay.Groups(answer))
            {
                panel.Children.Add(ApiUsageView.Heading(group.Heading, top: 16));

                foreach (ApiUsageRouteLines route in group.Routes)
                {
                    panel.Children.Add(new TextBlock
                    {
                        Text = route.Title,
                        FontWeight = FontWeight.Medium,
                        Margin = new Thickness(12, 6, 0, 0),
                        TextWrapping = TextWrapping.Wrap
                    });

                    TextBlock summary = ApiUsageView.Line(route.Summary, indent: 24);
                    summary.Opacity = 0.8;
                    panel.Children.Add(summary);

                    foreach (IReadOnlyList<ReviewNoteRun> version in route.Versions)
                        panel.Children.Add(ApiUsageView.Line(version, indent: 36));
                }
            }
        }

        private static TextBlock Heading(string text, double top) =>
            new()
            {
                Text = text,
                FontWeight = FontWeight.SemiBold,
                Margin = new Thickness(0, top, 0, 2),
                TextWrapping = TextWrapping.Wrap
            };

        private static TextBlock Line(IReadOnlyList<ReviewNoteRun> runs, double indent)
        {
            var block = new TextBlock { TextWrapping = TextWrapping.Wrap, Margin = new Thickness(indent, 0, 0, 0) };
            TabMaintainer.ShowCounts(block, runs);
            return block;
        }

        private void ShowMessage(string? message, bool isError) =>
            WindowMessage.Show(this.FindControl<TextBlock>("MessageText"), message, isError);

        private void SetText(string name, string text)
        {
            if (this.FindControl<TextBlock>(name) is TextBlock block)
                block.Text = text;
        }

        // ---- For tests ----------------------------------------------------------------------------

        // Every line shown, in order.
        internal IReadOnlyList<string> LinesForTests() =>
            this.FindControl<StackPanel>("UsagePanel")?.Children.OfType<TextBlock>().Select(TabMaintainer.TextOf).ToList() ?? [];

        internal ComboBox DaysBoxForTests => this.FindControl<ComboBox>("DaysBox")!;

        internal string? MessageForTests =>
            this.FindControl<TextBlock>("MessageText") is { IsVisible: true } message ? message.Text : null;
    }
}
