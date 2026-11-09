using Avalonia.Interactivity;
using Avalonia.Threading;
using Handlers.DataHandling;
using Handlers.Theming;
using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace CRT
{
    // ###########################################################################################
    // *** "SOMETHING WENT WRONG", AND "CRT CLOSED UNEXPECTEDLY" (owner request, 2026-11-07: "I think we must show every
    // ERROR and WARNING in here in the UI ... the end-goal here would be not to be dependent on manually viewing that
    // logfile"; "implement what you believe is the most optimal solution as seen from a user perspective"). ***
    //
    // Two cards of Main's NoticeStack:
    //
    //   - PROBLEM: every error CRT logs (ProblemsLog - a CRIT line) that the place it happened does not say itself - a
    //     button whose work failed, a board or a file that could not be read, something that could not be saved. It says
    //     the latest in plain words, and how many since it was last closed; it STAYS until closed, as somebody's work may
    //     not have been done. Errors logged before the window was there - at launch, while the data was read - are said
    //     as it opens. A warning needs nothing done and gets no card: it is listed in "Errors and warnings" only.
    //   - LAST CRASH: a fatal error ends CRT with nothing on screen, so the next launch says so, once per crash report
    //     (CrashReports, UserSettings.LastCrashReportSaid), the report listed with the rest.
    //
    // "Show details" on either opens "Errors and warnings" (ProblemsWindow). Started by StartAsync - a test's window
    // watches the log it hands over (WatchProblems).
    //
    // Part of the Main partial class - see Main.axaml.cs for the file map.
    // ###########################################################################################
    public partial class Main
    {
        private ProblemsLog? thisProblemsWatched;

        // Errors said since the card was last closed.
        private int thisProblemsSinceClosed;

        internal int ProblemsOnCardForTests => this.thisProblemsSinceClosed;

        // ###########################################################################################
        // Watches `log` from now on, and says what it already holds - each error once, whatever watches it next.
        // ###########################################################################################
        internal void WatchProblems(ProblemsLog log)
        {
            ArgumentNullException.ThrowIfNull(log);

            if (this.thisProblemsWatched is not null)
            {
                return;
            }

            this.thisProblemsWatched = log;
            this.TabConfiguration.WatchProblems(log);

            // What it holds and what comes after, in one step (ProblemsLog.Watch) - each error said once.
            foreach (LoggedProblem earlier in log.Watch(problem => Dispatcher.UIThread.Post(() => this.ShowProblem(problem))).Reverse())
            {
                this.ShowProblem(earlier);
            }
        }

        // Whether the card says it: an error nothing else on screen says.
        internal static bool CardSays(LoggedProblem problem) => problem.IsError && !problem.ShownElsewhere;

        internal void ShowProblem(LoggedProblem problem)
        {
            if (!Main.CardSays(problem))
            {
                return;
            }

            this.thisProblemsSinceClosed++;

            string said = this.thisProblemsSinceClosed == 1
                ? $"Something went wrong: {problem.Readable}"
                : $"Something went wrong [{this.thisProblemsSinceClosed}] times - the latest: {problem.Readable}";

            NoticeEmphasis.Show(this.ProblemBannerText, said);
            this.ProblemBanner.IsVisible = true;
        }

        private void OnProblemBannerDismiss(object? sender, RoutedEventArgs e)
        {
            this.ProblemBanner.IsVisible = false;
            this.thisProblemsSinceClosed = 0;
        }

        private void OnProblemBannerDetailsClick(object? sender, RoutedEventArgs e) =>
            ProblemsWindow.ShowFor(this, this.thisProblemsWatched);

        // ###########################################################################################
        // The last fatal crash, said once: its card, and its report among the errors. `crashFileText` is the crash file as
        // read (null with none), `now` the launch's time. With nothing stored of what was said - the first launch of a CRT
        // that keeps it - nothing is said and `now` is stored, so a crash of this run is said at the next launch
        // (CrashReports.IsNotSaidYet says why).
        // ###########################################################################################
        internal void SayLastCrashIfNotSaid(string? crashFileText, DateTime now)
        {
            if (!CrashReports.KnowsWhatWasSaid(UserSettings.LastCrashReportSaid))
            {
                UserSettings.LastCrashReportSaid = CrashReports.TimeTextOf(now);
                return;
            }

            if (CrashReports.LatestFatal(crashFileText) is not { } report ||
                !CrashReports.IsNotSaidYet(report, UserSettings.LastCrashReportSaid))
            {
                return;
            }

            (this.thisProblemsWatched ?? ProblemsLog.Current).Add(new LoggedProblem(report.When, ProblemLevel.Error, report.Text, ShownElsewhere: true));

            NoticeEmphasis.Show(this.LastCrashBannerText, CrashReports.NoticeOf(report));
            this.LastCrashBanner.IsVisible = true;
            UserSettings.LastCrashReportSaid = CrashReports.TimeTextOf(report);
        }

        private void OnLastCrashBannerDismiss(object? sender, RoutedEventArgs e) => this.LastCrashBanner.IsVisible = false;

        // StartAsync's part: the run's log watched, and the crash file read off the UI thread.
        private async Task StartSayingProblemsAsync()
        {
            this.WatchProblems(ProblemsLog.Current);

            try
            {
                string path = CrashLogger.ResolveCrashFilePath();
                string? text = await Task.Run(() => File.Exists(path) ? File.ReadAllText(path) : null);

                this.SayLastCrashIfNotSaid(text, DateTime.Now);
            }
            catch (Exception ex)
            {
                Logger.Warning($"The crash file could not be read to see whether CRT closed unexpectedly last time - [{ex.Message}]");
            }
        }
    }
}
