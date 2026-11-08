using Avalonia.Controls;
using Avalonia.Input.Platform;
using Avalonia.Interactivity;
using Avalonia.Threading;
using Handlers.DataHandling;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Threading;

namespace CRT
{
    // ###########################################################################################
    // "ERRORS AND WARNINGS" - every warning and error of the run, newest first (ProblemsLog; owner request, 2026-11-07:
    // "the end-goal here would be not to be dependent on manually viewing that logfile"). One window at a time, kept up
    // to date while it is open; "Copy all" for the "Feedback" tab. Opened from Main's "Something went wrong" card and from
    // the "Configuration" tab (ShowFor).
    // ###########################################################################################
    public partial class ProblemsWindow : Window
    {
        private static ProblemsWindow? OpenOne;

        private readonly ProblemsLog thisLog;

        public ProblemsWindow()
            : this(ProblemsLog.Current)
        {
        }

        internal ProblemsWindow(ProblemsLog log)
        {
            this.thisLog = log;
            this.InitializeComponent();
            this.ProblemsList.ItemsSource = this.thisShown;

            this.thisLog.Added += this.OnProblemAdded;
            this.Closed += (_, _) => this.thisLog.Added -= this.OnProblemAdded;

            this.ShowProblems();
        }

        // ###########################################################################################
        // The window, over `owner` - the one open brought to the front rather than a second. `log` for a test.
        // ###########################################################################################
        internal static ProblemsWindow ShowFor(Window owner, ProblemsLog? log = null)
        {
            if (ProblemsWindow.OpenOne is { IsVisible: true } open)
            {
                open.Activate();
                return open;
            }

            var window = new ProblemsWindow(log ?? ProblemsLog.Current);
            ProblemsWindow.OpenOne = window;
            window.Closed += (_, _) =>
            {
                if (ReferenceEquals(ProblemsWindow.OpenOne, window))
                {
                    ProblemsWindow.OpenOne = null;
                }
            };
            window.Show(owner);

            return window;
        }

        // What it lists - for the tests.
        internal IReadOnlyList<LoggedProblem> ShownForTests => this.ProblemsList.ItemsSource?.Cast<LoggedProblem>().ToList() ?? [];

        // ###########################################################################################
        // *** WHAT IS LOGGED WHILE IT IS OPEN IS SHOWN ONCE PER BURST, THE LIST CHANGED IN PLACE (code review, 2026-11-08).
        // *** Every line logged posted the whole list built again - a data check giving up on many files logs a warning
        // each - and every build replaced the rows and the selection, the details box under it flickering. Now one
        // catching-up is posted until it has run (from any thread: the log is written from the pool too), and the rows
        // are changed in place (InPlaceList), the one picked kept as long as it is listed.
        // ###########################################################################################
        private readonly ObservableCollection<LoggedProblem> thisShown = new();

        private int thisShowPosted;

        // How often the list was brought up to date - for the tests.
        internal int ShowsForTests { get; private set; }

        private void OnProblemAdded(LoggedProblem problem)
        {
            if (Interlocked.Exchange(ref this.thisShowPosted, 1) == 0)
            {
                Dispatcher.UIThread.Post(() =>
                {
                    Interlocked.Exchange(ref this.thisShowPosted, 0);
                    this.ShowProblems();
                });
            }
        }

        private void ShowProblems()
        {
            this.ShowsForTests++;
            bool warnings = this.ShowWarningsCheckBox.IsChecked == true;

            List<LoggedProblem> shown = this.thisLog.Newest().Where(problem => warnings || problem.IsError).ToList();

            InPlaceList.Sync(
                this.thisShown,
                shown,
                (item, wanted) => ReferenceEquals(item, wanted),
                wanted => wanted,
                keep: this.ProblemsList.SelectedItem as LoggedProblem);

            this.NothingText.IsVisible = shown.Count == 0;
            this.ProblemsList.IsVisible = shown.Count > 0;
            this.NothingText.Text = warnings || this.thisLog.Count == 0
                ? "Nothing has gone wrong since CRT started."
                : "No error since CRT started - only warnings, which \"Show warnings too\" lists.";
        }

        private void OnShowWarningsChanged(object? sender, RoutedEventArgs e) => this.ShowProblems();

        private void OnProblemSelected(object? sender, SelectionChangedEventArgs e) =>
            this.DetailsBox.Text = (this.ProblemsList.SelectedItem as LoggedProblem)?.Message.Trim() ?? string.Empty;

        private async void OnCopyAllClick(object? sender, RoutedEventArgs e)
        {
            try
            {
                if (TopLevel.GetTopLevel(this)?.Clipboard is { } clipboard)
                {
                    await clipboard.SetTextAsync(ProblemsLog.AsText(this.thisLog.Newest(), AppConfig.AppDisplayVersionString));
                    this.CopiedText.Text = "Copied to the clipboard";
                }
            }
            catch (Exception ex)
            {
                // Said here, never on the list it is about: copying it again would copy this too.
                this.CopiedText.Text = "The clipboard could not be used just now - try again";
                Logger.Info($"Copying the errors and warnings failed - [{ex.Message}]");
            }
        }

        private void OnCloseClick(object? sender, RoutedEventArgs e) => this.Close();
    }
}
