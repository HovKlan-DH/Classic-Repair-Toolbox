using System;
using System.Collections.Generic;
using System.Globalization;
using Handlers.DataHandling;

namespace Handlers.MaintainerHandling
{
    // ###########################################################################################
    // "Account" > SERVER VERSION (owner requests, 2026-10-04: "I would like to see the server version listed, so
    // it is clear to me what has been deployed", then as an entry of its own, every maintainer's).
    //
    // *** PANELS, EACH SAYING AT A GLANCE HOW IT STANDS (owner request, 2026-11-24: "make the "Account" page a
    // little prettier to look at, so it is more easy to overview" - "I meant for the "Server version" menu").
    // *** Until then a row of lines, "Server version [4.6.0]", "API server version [1]", ... under a paragraph
    // explaining what an API revision is. Now one panel per thing the server is asked about, as "Board status"
    // on the "Contribute" tab is drawn (ServerVersionView): its title, a line saying how it stands - in words,
    // its color and icon saying the same - and its facts in two columns, each value bold:
    //
    //   "Server"             the version deployed, its API revision and this CRT's - and the line says what
    //                        the two revisions mean together, so nobody has to know what one is;
    //   "Automatic updates"  how the server keeps itself up to date from GitHub (server 10.3.0);
    //   "Mail"               only while mails wait that the server could not send (2026-10-13).
    //
    // Under them the server's own changelog (ServerStatusAnswer.Changelog, server 12.2.0), drawn as the "About"
    // tab draws CHANGELOG.md.
    //
    // The version is GET /api/health's (public, asked with no session), the rest GET /api/review/server-status's
    // (with the session), both read each time the entry is shown - so a deploy shows the next time.
    // ###########################################################################################
    public static class ServerVersionDisplay
    {
        public const string Heading = "Server version";

        public const string Introduction =
            "What is deployed on the server this CRT talks to, how the server keeps itself up to date, and what changed in each of its versions.";

        public const string ChangelogTitle = "Full changelog for server";

        // A server that answered but sent no changelog - one older than 12.2.0.
        public const string NoChangelog = "This server sends no changelog - it is older than 12.2.0.";

        // -----------------------------------------------------------------------------------
        // "Server"
        // -----------------------------------------------------------------------------------

        public const string ServerTitle = "Server";

        public const string VersionLabel = "Version";

        public const string ServerApiLabel = "API revision";

        public const string ApplicationApiLabel = "This CRT's API revision";

        public const string Asking = "Asking the server";

        public const string NoAnswer = "The server did not answer";

        public const string SameApi = "This CRT and the server speak the same API";

        public const string CrtBehind = "This CRT is behind the server - update CRT to submit and to use this tab";

        public const string ServerBehind = "The server is behind this CRT - its newer version is not deployed yet";

        // A server older than 4.6.0 answers with its version but no API revision.
        public const string NoApiRevision = "The server is older than 4.6.0 and does not say its API revision";

        // -----------------------------------------------------------------------------------
        // "Automatic updates" (owner request, 2026-11-09: "If this is possible it could show the status in
        // 'Maintainer' => 'Account' => 'Server version'") - from ServerStatusAnswer.Updates, which CRT.Server's
        // updater writes. The updater's own sentence about its last look is a fact of its own ("Found"); a look
        // older than ten minutes says so, as the timer looks every minute (owner request, 2026-11-09), so it may
        // have stopped - an install, with its download and two minutes of health checks, is well inside that.
        // Nothing at all from an older server, which has no updater.
        // -----------------------------------------------------------------------------------

        public const string UpdatesTitle = "Automatic updates";

        public const string InstallsLabel = "Installs";

        public const string ReleasesOnly = "releases only";

        public const string EveryCommit = "every commit to main";

        public const string NotSetUp = "Not set up on this server";

        public const string LastLookLabel = "Last look";

        public const string FoundLabel = "Found";

        public const string LastInstalledLabel = "Last installed";

        // A look this old means the timer may not be running.
        public static readonly TimeSpan LookIsStaleAfter = TimeSpan.FromMinutes(10);

        public const string StaleLook = " - over ten minutes ago";

        public const string TimerMayHaveStopped = "No look for over ten minutes - its timer may not be running";

        public const string UpToDate = "Up to date";

        public const string JustInstalled = "Up to date - the newest build was installed";

        public const string WouldInstall = "A newer build would be installed (a trial run)";

        public const string Held = "A newer build is held back";

        public const string Failed = "The last install failed";

        public const string Refused = "A build was refused";

        public const string LookFailed = "The last look went wrong";

        public const string SetUp = "Set up";

        // -----------------------------------------------------------------------------------
        // "Mail" (code review, 2026-10-13): the mails waiting in the server's outbox because the mail system
        // refused them - said only when above zero, as a mail system that is down is something to look at.
        // -----------------------------------------------------------------------------------

        public const string MailTitle = "Mail";

        public const string UnsentMailsLabel = "Not sent yet";

        public const string MailsWaiting = "Mails the server could not send yet - it keeps trying";

        // ###########################################################################################
        // The panels, in order. `version` null is no answer (yet, while `asking`); the application's revision is
        // known either way, so the server's panel always says it.
        // ###########################################################################################
        public static IReadOnlyList<ServerPanel> Panels(
            string? version,
            int? serverApiRevision,
            int applicationApiRevision,
            bool asking = false,
            int? unsentMails = null,
            ServerUpdatesAnswer? updates = null,
            DateTimeOffset? now = null)
        {
            var panels = new List<ServerPanel> { ServerVersionDisplay.ServerPanelOf(version, serverApiRevision, applicationApiRevision, asking) };

            if (updates != null)
                panels.Add(ServerVersionDisplay.UpdatesPanelOf(updates, now ?? DateTimeOffset.Now));

            if (unsentMails is int unsent && unsent > 0)
            {
                panels.Add(new ServerPanel(
                    ServerVersionDisplay.MailTitle,
                    ServerPanelState.Problem,
                    ServerVersionDisplay.MailsWaiting,
                    [ServerVersionDisplay.Row(ServerVersionDisplay.UnsentMailsLabel, unsent.ToString("N0", CultureInfo.InvariantCulture))]));
            }

            return panels;
        }

        private static ServerPanel ServerPanelOf(string? version, int? serverApiRevision, int applicationApiRevision, bool asking)
        {
            var rows = new List<ServerPanelRow>();
            string application = applicationApiRevision.ToString(CultureInfo.InvariantCulture);

            if (string.IsNullOrWhiteSpace(version))
            {
                rows.Add(ServerVersionDisplay.Row(ServerVersionDisplay.ApplicationApiLabel, application));

                return asking
                    ? new ServerPanel(ServerVersionDisplay.ServerTitle, ServerPanelState.Neutral, ServerVersionDisplay.Asking, rows)
                    : new ServerPanel(ServerVersionDisplay.ServerTitle, ServerPanelState.Problem, ServerVersionDisplay.NoAnswer, rows);
            }

            rows.Add(ServerVersionDisplay.Row(ServerVersionDisplay.VersionLabel, version.Trim()));

            if (serverApiRevision is not int server)
            {
                rows.Add(ServerVersionDisplay.Row(ServerVersionDisplay.ApplicationApiLabel, application));
                return new ServerPanel(ServerVersionDisplay.ServerTitle, ServerPanelState.Neutral, ServerVersionDisplay.NoApiRevision, rows);
            }

            rows.Add(ServerVersionDisplay.Row(ServerVersionDisplay.ServerApiLabel, server.ToString(CultureInfo.InvariantCulture)));
            rows.Add(ServerVersionDisplay.Row(ServerVersionDisplay.ApplicationApiLabel, application));

            // The same number: they speak the same API. The server's higher: it turns this CRT away (ClientVersionPolicy).
            // This CRT's higher: the server this CRT was built for is not deployed yet.
            (ServerPanelState state, string summary) = server.CompareTo(applicationApiRevision) switch
            {
                0 => (ServerPanelState.Good, ServerVersionDisplay.SameApi),
                > 0 => (ServerPanelState.Problem, ServerVersionDisplay.CrtBehind),
                _ => (ServerPanelState.Waiting, ServerVersionDisplay.ServerBehind),
            };

            return new ServerPanel(ServerVersionDisplay.ServerTitle, state, summary, rows);
        }

        private static ServerPanel UpdatesPanelOf(ServerUpdatesAnswer updates, DateTimeOffset now)
        {
            if (!updates.SetUp)
                return new ServerPanel(ServerVersionDisplay.UpdatesTitle, ServerPanelState.Neutral, ServerVersionDisplay.NotSetUp, []);

            var rows = new List<ServerPanelRow>();

            string channel = updates.Channel switch
            {
                ServerUpdateOutcomes.ReleasesChannel => ServerVersionDisplay.ReleasesOnly,
                ServerUpdateOutcomes.MainChannel => ServerVersionDisplay.EveryCommit,
                _ => updates.Channel?.Trim() ?? string.Empty,
            };

            if (channel.Length > 0)
                rows.Add(ServerVersionDisplay.Row(ServerVersionDisplay.InstallsLabel, channel));

            bool stale = false;

            if (updates.CheckedUtc is DateTimeOffset checkedUtc)
            {
                stale = now - checkedUtc > ServerVersionDisplay.LookIsStaleAfter;

                rows.Add(new ServerPanelRow(
                    ServerVersionDisplay.LastLookLabel,
                    stale
                        ? [new ReviewNoteRun(ServerVersionDisplay.When(checkedUtc), IsCount: true), new ReviewNoteRun(ServerVersionDisplay.StaleLook, IsCount: false)]
                        : [new ReviewNoteRun(ServerVersionDisplay.When(checkedUtc), IsCount: true)]));
            }

            // The updater's own sentence - what it found, or what waits and why. Never bold: it is no value.
            if (updates.Message?.Trim() is { Length: > 0 } message)
                rows.Add(new ServerPanelRow(ServerVersionDisplay.FoundLabel, [new ReviewNoteRun(message, IsCount: false)]));

            if (!string.IsNullOrWhiteSpace(updates.InstalledVersion))
            {
                var installed = new List<ReviewNoteRun> { new(updates.InstalledVersion.Trim(), IsCount: true) };

                if (updates.InstalledUtc is DateTimeOffset installedUtc)
                {
                    installed.Add(new ReviewNoteRun(" on ", IsCount: false));
                    installed.Add(new ReviewNoteRun(ServerVersionDisplay.When(installedUtc), IsCount: true));
                }

                if (!string.IsNullOrWhiteSpace(updates.InstalledFromVersion))
                {
                    installed.Add(new ReviewNoteRun(", replacing ", IsCount: false));
                    installed.Add(new ReviewNoteRun(updates.InstalledFromVersion.Trim(), IsCount: true));
                }

                rows.Add(new ServerPanelRow(ServerVersionDisplay.LastInstalledLabel, installed));
            }

            (ServerPanelState state, string summary) = updates.Outcome switch
            {
                ServerUpdateOutcomes.UpToDate => (ServerPanelState.Good, ServerVersionDisplay.UpToDate),
                ServerUpdateOutcomes.Installed => (ServerPanelState.Good, ServerVersionDisplay.JustInstalled),
                ServerUpdateOutcomes.WouldInstall => (ServerPanelState.Waiting, ServerVersionDisplay.WouldInstall),
                ServerUpdateOutcomes.Held => (ServerPanelState.Waiting, ServerVersionDisplay.Held),
                ServerUpdateOutcomes.Failed => (ServerPanelState.Problem, ServerVersionDisplay.Failed),
                ServerUpdateOutcomes.Refused => (ServerPanelState.Problem, ServerVersionDisplay.Refused),
                ServerUpdateOutcomes.Error => (ServerPanelState.Problem, ServerVersionDisplay.LookFailed),
                _ => (ServerPanelState.Neutral, ServerVersionDisplay.SetUp),
            };

            // A timer that may have stopped outweighs good news, never a problem already said.
            if (stale && state != ServerPanelState.Problem)
                (state, summary) = (ServerPanelState.Waiting, ServerVersionDisplay.TimerMayHaveStopped);

            return new ServerPanel(ServerVersionDisplay.UpdatesTitle, state, summary, rows);
        }

        // "2026-November-9 at 14:20", local time - the Maintainer tab's way (ReviewTableWording).
        private static string When(DateTimeOffset at) =>
            SubmissionReceiptPresenter.FormatDate(at) + " at " + at.ToLocalTime().ToString("HH:mm", CultureInfo.InvariantCulture);

        // A fact whose value is one bold piece.
        private static ServerPanelRow Row(string label, string value) =>
            new(label, [new ReviewNoteRun(value, IsCount: true)]);
    }

    // How a panel stands - its edge, its icon and the color of its line.
    public enum ServerPanelState
    {
        Good,
        Waiting,
        Problem,
        Neutral,
    }

    // A fact of a panel: its label in the left column, its value beside it - the pieces marked IsCount bold.
    public sealed record ServerPanelRow(string Label, IReadOnlyList<ReviewNoteRun> Value);

    // A panel: its title, the line saying how it stands, and its facts.
    public sealed record ServerPanel(string Title, ServerPanelState State, string Summary, IReadOnlyList<ServerPanelRow> Rows);
}
