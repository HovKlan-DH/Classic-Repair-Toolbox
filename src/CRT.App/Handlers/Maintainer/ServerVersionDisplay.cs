using System.Collections.Generic;
using System.Globalization;

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
        public static IReadOnlyList<IReadOnlyList<ReviewNoteRun>> Lines(string? version, int? serverApiRevision, int applicationApiRevision, bool asking = false)
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
