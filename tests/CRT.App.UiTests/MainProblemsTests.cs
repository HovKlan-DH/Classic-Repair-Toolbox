using Avalonia.Controls;
using Avalonia.Interactivity;
using Avalonia.Threading;
using CRT;
using Handlers.DataHandling;
using Handlers.Theming;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// "SOMETHING WENT WRONG" AND "CRT CLOSED UNEXPECTEDLY" (Main.Problems.cs; owner request, 2026-11-07: "I think we must show
// every ERROR and WARNING in here in the UI ... the end-goal here would be not to be dependent on manually viewing that
// logfile"). Each window watches a log of the test's own - the run's is Logger's, which a test never starts. Main's
// StartAsync is never called (MainWindowTests' header).
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class MainProblemsTests : IDisposable
{
    private readonly TempWorkspace thisWorkspace = new();

    public MainProblemsTests()
    {
        WorklogManager.LoadFrom(this.thisWorkspace.Path_("Workbook"));
        UserSettings.LoadFrom(this.thisWorkspace.WriteFile("settings.json", "{}"));
        DraftManager.LoadFrom(this.thisWorkspace.Path_("Drafts"));
    }

    public void Dispose()
    {
        DraftManager.LoadFrom(string.Empty);
        WorklogManager.LoadFrom(this.thisWorkspace.Path_("Workbook-after"));
        UserSettings.LoadFrom(this.thisWorkspace.WriteFile("settings-after.json", "{}"));
        this.thisWorkspace.Dispose();
    }

    private static LoggedProblem Problem(ProblemLevel level, string message, bool shownElsewhere = false) =>
        new(DateTime.Now, level, message, shownElsewhere);

    private static string CardText(CRT.Main main, string name) => NoticeEmphasis.TextOf(main.GetControl<TextBlock>(name));

    private static void Click(CRT.Main main, string button) =>
        main.GetControl<Button>(button).RaiseEvent(new RoutedEventArgs(Button.ClickEvent));

    // ###########################################################################################
    // An error says itself at once and stays - one logged before the window was there too - and counts the rest until
    // closed; a warning, or an error the place it happened says itself, gets no card.
    // ###########################################################################################
    [Fact]
    public void An_error_is_said_on_the_card_and_counted_until_closed_and_a_warning_is_not()
    {
        UiTest.Run(() =>
        {
            var log = new ProblemsLog();
            log.Add(Problem(ProblemLevel.Error, "Failed to load board file [Commodore/C64/250407/board.json]"));
            log.Add(Problem(ProblemLevel.Warning, "Download given up"));

            var main = new CRT.Main();
            main.WatchProblems(log);

            Border card = main.GetControl<Border>("ProblemBanner");
            Assert.True(card.IsVisible);
            Assert.Equal("Something went wrong: Failed to load board file [Commodore/C64/250407/board.json]", CardText(main, "ProblemBannerText"));

            // A warning, and an error said elsewhere: nothing more on the card.
            log.Add(Problem(ProblemLevel.Warning, "Another download given up"));
            log.Add(Problem(ProblemLevel.Error, "The update could not be installed", shownElsewhere: true));
            Dispatcher.UIThread.RunJobs();
            Assert.Equal(1, main.ProblemsOnCardForTests);

            log.Add(Problem(ProblemLevel.Error, "The Contribute tab's [Table] failed - [System.InvalidOperationException: Sequence contains no elements\n   at x"));
            Dispatcher.UIThread.RunJobs();
            Assert.Equal("Something went wrong [2] times - the latest: The Contribute tab's [Table] failed - [Sequence contains no elements]", CardText(main, "ProblemBannerText"));

            Click(main, "ProblemBannerDismissButton");
            Assert.False(card.IsVisible);
            Assert.Equal(0, main.ProblemsOnCardForTests);

            // The next one starts again from one.
            log.Add(Problem(ProblemLevel.Error, "Failed to export workbook [#3]: [Access denied]"));
            Dispatcher.UIThread.RunJobs();
            Assert.True(card.IsVisible);
            Assert.StartsWith("Something went wrong: Failed to export workbook", CardText(main, "ProblemBannerText"), StringComparison.Ordinal);
        });
    }

    // ###########################################################################################
    // A crash last time: said once, at the next launch, its report listed among the errors - never on the error card,
    // whose count is of what went wrong in this run.
    // ###########################################################################################
    [Fact]
    public void A_crash_last_time_is_said_once_and_its_report_listed()
    {
        string crashFile =
            "=== CRASH REPORT ===\r\nTime      : 2026-11-07 19:09:51\r\nSeverity  : FATAL\r\nMessage   : Attempt to call InvalidateArrange on wrong LayoutManager.\r\n\r\n";

        UiTest.Run(() =>
        {
            // A launch of this CRT before the crash: what it said up to is stored.
            UserSettings.LastCrashReportSaid = "2026-11-07 08:00:00";

            var log = new ProblemsLog();
            var main = new CRT.Main();
            main.WatchProblems(log);

            main.SayLastCrashIfNotSaid(crashFile, new DateTime(2026, 11, 7, 20, 0, 0));

            Assert.True(main.GetControl<Border>("LastCrashBanner").IsVisible);
            Assert.StartsWith("CRT closed unexpectedly the last time it ran, on 2026-11-07 at 19:09.", CardText(main, "LastCrashBannerText"), StringComparison.Ordinal);
            Assert.Equal("2026-11-07 19:09:51", UserSettings.LastCrashReportSaid);

            LoggedProblem listed = Assert.Single(log.Newest());
            Assert.True(listed.IsError);
            Assert.True(listed.ShownElsewhere);
            Assert.False(main.GetControl<Border>("ProblemBanner").IsVisible);

            // The next launch: said already.
            var next = new CRT.Main();
            next.WatchProblems(new ProblemsLog());
            next.SayLastCrashIfNotSaid(crashFile, new DateTime(2026, 11, 7, 21, 0, 0));
            Assert.False(next.GetControl<Border>("LastCrashBanner").IsVisible);
        });
    }

    // ###########################################################################################
    // *** THE FIRST LAUNCH OF A CRT THAT REMEMBERS WHAT IT SAID SAYS NO OLD CRASH (code review, 2026-11-07). *** The crash
    // file is never emptied, and with nothing stored its newest fatal report - September's, from an earlier pre-release -
    // was said as "the last time it ran" though several runs had ended well since. Nothing is said then, and the launch's
    // own time is stored: a crash of this run is said at the next launch.
    // ###########################################################################################
    [Fact]
    public void A_first_launch_says_no_old_crash_and_a_crash_of_its_run_is_said_next_time()
    {
        string september =
            "=== CRASH REPORT ===\r\nTime      : 2026-09-14 10:02:00\r\nSeverity  : FATAL\r\nMessage   : Something long ago.\r\n\r\n";
        string tonight = september +
            "=== CRASH REPORT ===\r\nTime      : 2026-11-07 20:15:00\r\nSeverity  : FATAL\r\nMessage   : Something tonight.\r\n\r\n";

        UiTest.Run(() =>
        {
            Assert.Null(UserSettings.LastCrashReportSaid);

            var log = new ProblemsLog();
            var first = new CRT.Main();
            first.WatchProblems(log);
            first.SayLastCrashIfNotSaid(september, new DateTime(2026, 11, 7, 20, 0, 0));

            Assert.False(first.GetControl<Border>("LastCrashBanner").IsVisible);
            Assert.Empty(log.Newest());
            Assert.Equal("2026-11-07 20:00:00", UserSettings.LastCrashReportSaid);

            var next = new CRT.Main();
            next.WatchProblems(new ProblemsLog());
            next.SayLastCrashIfNotSaid(tonight, new DateTime(2026, 11, 7, 21, 0, 0));

            Assert.True(next.GetControl<Border>("LastCrashBanner").IsVisible);
            Assert.StartsWith("CRT closed unexpectedly the last time it ran, on 2026-11-07 at 20:15.", CardText(next, "LastCrashBannerText"), StringComparison.Ordinal);
        });
    }

    // ###########################################################################################
    // "Errors and warnings": the run's, newest first; warnings left out when unticked; the whole line of the one picked
    // under the list; and it follows what is logged while it is open.
    // ###########################################################################################
    [Fact]
    public void Errors_and_warnings_lists_the_newest_first_and_follows_the_log()
    {
        UiTest.Run(() =>
        {
            var log = new ProblemsLog();
            log.Add(new LoggedProblem(new DateTime(2026, 11, 7, 19, 0, 0), ProblemLevel.Warning, "A download given up"));
            log.Add(new LoggedProblem(new DateTime(2026, 11, 7, 19, 1, 0), ProblemLevel.Error, "Failed to export workbook [#3]: [Access denied]\n   at stack"));

            var owner = new Window();
            owner.Show();

            ProblemsWindow window = ProblemsWindow.ShowFor(owner, log);

            try
            {
                Assert.Equal(["Failed to export workbook [#3]: [Access denied]\n   at stack", "A download given up"], window.ShownForTests.Select(problem => problem.Message));

                window.GetControl<CheckBox>("ShowWarningsCheckBox").IsChecked = false;
                Dispatcher.UIThread.RunJobs();
                Assert.Single(window.ShownForTests);

                window.GetControl<ListBox>("ProblemsList").SelectedIndex = 0;
                Dispatcher.UIThread.RunJobs();
                Assert.Equal("Failed to export workbook [#3]: [Access denied]\n   at stack", window.GetControl<TextBox>("DetailsBox").Text);

                log.Add(new LoggedProblem(new DateTime(2026, 11, 7, 19, 2, 0), ProblemLevel.Error, "Failed to load board file [x]"));
                Dispatcher.UIThread.RunJobs();
                Assert.Equal("Failed to load board file [x]", window.ShownForTests[0].Message);

                // Asked for again: the same window, to the front.
                Assert.Same(window, ProblemsWindow.ShowFor(owner, log));
            }
            finally
            {
                window.Close();
                owner.Close();
            }

            // Nothing at all: it says so.
            var another = new Window();
            another.Show();
            ProblemsWindow empty = ProblemsWindow.ShowFor(another, new ProblemsLog());
            Assert.True(empty.GetControl<TextBlock>("NothingText").IsVisible);
            empty.Close();
            another.Close();
        });
    }

    // ###########################################################################################
    // *** A BURST LOGGED WHILE IT IS OPEN IS SHOWN ONCE, THE ONE PICKED STAYING PICKED (code review, 2026-11-08). *** Each
    // line logged posted the whole list built again - a data check giving up on many files logs a warning each - and
    // each build replaced the rows and the selection, the details under the list flickering. Logged from the pool too,
    // as the log is.
    // ###########################################################################################
    [Fact]
    public void A_burst_logged_while_it_is_open_is_shown_once_and_the_one_picked_stays()
    {
        UiTest.Run(() =>
        {
            var log = new ProblemsLog();
            log.Add(new LoggedProblem(new DateTime(2026, 11, 8, 9, 0, 0), ProblemLevel.Error, "Failed to export workbook [#3]"));

            var owner = new Window();
            owner.Show();
            ProblemsWindow window = ProblemsWindow.ShowFor(owner, log);

            try
            {
                ListBox list = window.GetControl<ListBox>("ProblemsList");
                list.SelectedIndex = 0;
                Dispatcher.UIThread.RunJobs();

                LoggedProblem picked = window.ShownForTests[0];
                int selectionChanges = 0;
                list.SelectionChanged += (_, _) => selectionChanges++;
                int showsBefore = window.ShowsForTests;

                // Logged on the pool while the UI thread waits - nothing posted runs until the burst is over.
                Task.Run(() =>
                {
                    for (int line = 0; line < 50; line++)
                    {
                        log.Add(new LoggedProblem(new DateTime(2026, 11, 8, 9, 1, 0).AddSeconds(line), ProblemLevel.Warning, $"A download given up [{line}]"));
                    }
                }).Wait(TestContext.Current.CancellationToken);

                Dispatcher.UIThread.RunJobs();

                Assert.Equal(51, window.ShownForTests.Count);
                Assert.Equal(1, window.ShowsForTests - showsBefore);
                Assert.Same(picked, list.SelectedItem);
                Assert.Equal(0, selectionChanges);
                Assert.Equal("Failed to export workbook [#3]", window.GetControl<TextBox>("DetailsBox").Text);
            }
            finally
            {
                window.Close();
                owner.Close();
            }
        });
    }
}
