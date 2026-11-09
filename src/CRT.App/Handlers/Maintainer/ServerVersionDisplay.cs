using System;
using System.Collections.Generic;
using System.Globalization;
using Handlers.DataHandling;

namespace Handlers.MaintainerHandling
{
    // ###########################################################################################
    // "Account" > SERVER VERSION (owner requests, 2026-10-04: "I would like to see
    // the server version listed, so it is clear to me what has been deployed" - and then, as its own
    // entry in the list, worded as three lines:
    //
    //     Server version [4.6.0]
    //     API server version [1]
    //     API application version [1]
    //
    // with each value bold). It was one line under the administrator's tools, above the account row,
    // until then.
    //
    // The first two are GET /api/health's answer, read each time the entry is shown, so a deploy shows
    // the next time it is chosen. The API revisions are CRT.Data's ClientVersionContract.ApiRevision -
    // the server's, and this CRT's own: the same number means they speak the same API, a lower one
    // for this CRT means the server tells it to update. A server that did not answer is SAID, never
    // left blank - a blank would read as nothing deployed - and a server older than 4.6.0 reports no
    // revision, which is said too.
    // ###########################################################################################
    public static class ServerVersionDisplay
    {
        public const string ServerVersionLabel = "Server version";

        public const string ApiServerVersionLabel = "API server version";

        public const string ApiApplicationVersionLabel = "API application version";

        // Before the first answer, and when none came.
        public const string Asking = "Server version: asking the server";

        public const string NoAnswer = "Server version: the server did not answer";

        // A server older than 4.6.0 answers with its version but no API revision.
        public const string NoApiRevision = "API server version: not reported - the server is older than 4.6.0";

        // ###########################################################################################
        // The lines, each "Label [value]" with the value bold. `version` null is no answer (yet, while
        // `asking`); the application's revision is known either way, so it is always the last line.
        // ###########################################################################################
        // ###########################################################################################
        // *** THE MAILS THE SERVER COULD NOT SEND (code review, 2026-10-13). *** Read from the
        // signed-in /api/review/server-status (ServerStatusAnswer.UnsentMails), never the public
        // health answer, and said as a last line only when above zero: "Mails the server could not
        // send yet [3]" - a mail system that is down is something to look at, not to read past.
        // ###########################################################################################
        public const string UnsentMailsLabel = "Mails the server could not send yet";

        // ###########################################################################################
        // *** HOW THE SERVER KEEPS ITSELF UP TO DATE (owner request, 2026-11-09: "If this is possible it
        // could show the status in 'Maintainer' => 'Account' => 'Server version'"). *** From the signed-in
        // server-status answer (ServerStatusAnswer.Updates - CRT.Server's updater writes it), after the
        // version lines:
        //
        //     Automatic updates [releases only]
        //     Last look [2026-November-9 at 14:20]: Up to date: 10.3.0 is the newest server build ...
        //     Last installed [10.3.0] on [2026-November-9 at 14:05], replacing [10.2.1]
        //
        // The middle line is the updater's own sentence - what it found, or what waits and why. A look
        // older than ten minutes says so: the timer looks every minute (owner request, 2026-11-09), so it may
        // have stopped - an install, with its download and two minutes of health checks, is well inside that.
        // Nothing at all from an older server, which has no updater.
        // ###########################################################################################
        public const string AutomaticUpdatesLabel = "Automatic updates";

        public const string ReleasesOnly = "releases only";

        public const string EveryCommit = "every commit to main";

        public const string NotSetUp = "not set up on this server";

        public const string LastLookLabel = "Last look";

        public const string LastInstalledLabel = "Last installed";

        // A look this old means the timer may not be running.
        public static readonly TimeSpan LookIsStaleAfter = TimeSpan.FromMinutes(10);

        public const string StaleLook = " - over ten minutes ago, so its timer may not be running";

        public static IReadOnlyList<IReadOnlyList<ReviewNoteRun>> Lines(
            string? version,
            int? serverApiRevision,
            int applicationApiRevision,
            bool asking = false,
            int? unsentMails = null,
            ServerUpdatesAnswer? updates = null,
            DateTimeOffset? now = null)
        {
            List<IReadOnlyList<ReviewNoteRun>> lines = ServerVersionDisplay.VersionLines(version, serverApiRevision, applicationApiRevision, asking);

            if (unsentMails is int unsent && unsent > 0)
                lines.Add(ServerVersionDisplay.Labelled(ServerVersionDisplay.UnsentMailsLabel, unsent.ToString("N0", CultureInfo.InvariantCulture)));

            if (updates != null)
                lines.AddRange(ServerVersionDisplay.UpdateLines(updates, now ?? DateTimeOffset.Now));

            return lines;
        }

        private static IEnumerable<IReadOnlyList<ReviewNoteRun>> UpdateLines(ServerUpdatesAnswer updates, DateTimeOffset now)
        {
            if (!updates.SetUp)
            {
                yield return ServerVersionDisplay.Labelled(ServerVersionDisplay.AutomaticUpdatesLabel, ServerVersionDisplay.NotSetUp);
                yield break;
            }

            string channel = updates.Channel switch
            {
                ServerUpdateOutcomes.ReleasesChannel => ServerVersionDisplay.ReleasesOnly,
                ServerUpdateOutcomes.MainChannel => ServerVersionDisplay.EveryCommit,
                _ => updates.Channel ?? string.Empty,
            };

            yield return channel.Length > 0
                ? ServerVersionDisplay.Labelled(ServerVersionDisplay.AutomaticUpdatesLabel, channel)
                : [new ReviewNoteRun(ServerVersionDisplay.AutomaticUpdatesLabel, IsCount: false)];

            string message = updates.Message?.Trim() ?? string.Empty;

            if (updates.CheckedUtc is DateTimeOffset checkedUtc)
            {
                var line = new List<ReviewNoteRun>(ServerVersionDisplay.Labelled(ServerVersionDisplay.LastLookLabel, ServerVersionDisplay.When(checkedUtc)));

                if (now - checkedUtc > ServerVersionDisplay.LookIsStaleAfter)
                    line.Add(new ReviewNoteRun(ServerVersionDisplay.StaleLook, IsCount: false));

                if (message.Length > 0)
                    line.Add(new ReviewNoteRun(": " + message, IsCount: false));

                yield return line;
            }
            else if (message.Length > 0)
            {
                yield return [new ReviewNoteRun(message, IsCount: false)];
            }

            if (!string.IsNullOrWhiteSpace(updates.InstalledVersion))
            {
                var line = new List<ReviewNoteRun>(ServerVersionDisplay.Labelled(ServerVersionDisplay.LastInstalledLabel, updates.InstalledVersion.Trim()));

                if (updates.InstalledUtc is DateTimeOffset installedUtc)
                {
                    line.Add(new ReviewNoteRun(" on [", IsCount: false));
                    line.Add(new ReviewNoteRun(ServerVersionDisplay.When(installedUtc), IsCount: true));
                    line.Add(new ReviewNoteRun("]", IsCount: false));
                }

                if (!string.IsNullOrWhiteSpace(updates.InstalledFromVersion))
                {
                    line.Add(new ReviewNoteRun(", replacing [", IsCount: false));
                    line.Add(new ReviewNoteRun(updates.InstalledFromVersion.Trim(), IsCount: true));
                    line.Add(new ReviewNoteRun("]", IsCount: false));
                }

                yield return line;
            }
        }

        // "2026-November-9 at 14:20", local time - the Maintainer tab's way (ReviewTableWording).
        private static string When(DateTimeOffset at) =>
            SubmissionReceiptPresenter.FormatDate(at) + " at " + at.ToLocalTime().ToString("HH:mm", CultureInfo.InvariantCulture);

        private static List<IReadOnlyList<ReviewNoteRun>> VersionLines(string? version, int? serverApiRevision, int applicationApiRevision, bool asking)
        {
            var lines = new List<IReadOnlyList<ReviewNoteRun>>();

            if (string.IsNullOrWhiteSpace(version))
            {
                lines.Add([new ReviewNoteRun(asking ? ServerVersionDisplay.Asking : ServerVersionDisplay.NoAnswer, IsCount: false)]);
            }
            else
            {
                lines.Add(ServerVersionDisplay.Labelled(ServerVersionDisplay.ServerVersionLabel, version.Trim()));

                lines.Add(serverApiRevision is int server
                    ? ServerVersionDisplay.Labelled(ServerVersionDisplay.ApiServerVersionLabel, server.ToString(CultureInfo.InvariantCulture))
                    : [new ReviewNoteRun(ServerVersionDisplay.NoApiRevision, IsCount: false)]);
            }

            lines.Add(ServerVersionDisplay.Labelled(
                ServerVersionDisplay.ApiApplicationVersionLabel,
                applicationApiRevision.ToString(CultureInfo.InvariantCulture)));

            return lines;
        }

        // "Label [value]", the value bold - as every count and name in brackets in this tab.
        private static IReadOnlyList<ReviewNoteRun> Labelled(string label, string value) =>
        [
            new ReviewNoteRun(label + " [", IsCount: false),
            new ReviewNoteRun(value, IsCount: true),
            new ReviewNoteRun("]", IsCount: false)
        ];
    }
}
