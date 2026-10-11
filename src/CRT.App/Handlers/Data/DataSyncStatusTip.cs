using System;
using System.Collections.Generic;
using System.Globalization;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // The words of the data-sync icon's tooltip (Main.DataSyncStatus.cs): while a check runs, how
    // far it got - the same lines the data-check notice shows (DataCheckNotice) - and otherwise
    // what a click does and what the last check found (owner report, 2026-10-11).
    //
    // "All data files are up to date; protected contribution related files are [N]" is gone: the
    // protected files are CRT 2.x's contribution workbooks, which 3.0.0 no longer makes (owner,
    // 2026-10-11).
    // ###########################################################################################
    internal static class DataSyncStatusTip
    {
        // ###########################################################################################
        // What a running check is doing: a heading, the sync's own status line, and the file it is
        // on - each left out until the sync has said it. Shared by the tooltip and the notice.
        // ###########################################################################################
        public static string Checking(string sourceLabel, string progress, string file)
        {
            var lines = new List<string> { $"Checking data from the {sourceLabel} - please wait..." };

            if (!string.IsNullOrWhiteSpace(progress))
            {
                lines.Add(DataSyncStatusTip.OneLine(progress));
            }

            if (!string.IsNullOrWhiteSpace(file))
            {
                lines.Add(DataSyncStatusTip.OneLine(file));
            }

            return string.Join("\n", lines);
        }

        // ###########################################################################################
        // The tooltip. While a check runs: what it is doing and the file it is on. Otherwise: what
        // a click does, and the outcome of the last check when there was one.
        // ###########################################################################################
        public static string Build(
            bool isChecking,
            bool isEnabled,
            bool isHovering,
            string sourceLabel,
            string progress,
            string file,
            string? lastCheck)
        {
            if (isChecking)
            {
                return DataSyncStatusTip.Checking(sourceLabel, progress, file);
            }

            var lines = new List<string>();

            lines.Add(isEnabled
                ? "Data update is enabled. Click to refresh data now"
                : isHovering
                    ? "Data update at launch is disabled. Click to run a manual refresh now"
                    : "Data update at launch is disabled. Hover to show manual refresh");

            if (!string.IsNullOrWhiteSpace(lastCheck))
            {
                lines.Add(string.Empty);
                lines.Add(lastCheck);
            }

            return string.Join("\n", lines);
        }

        // ###########################################################################################
        // Whether a status the sync reports is progress worth showing. Its closing summaries are
        // not ("... is up to date", "Update complete ...", "Sync complete (3 new, ...)"): the
        // outcome line says that, in words a person reads, once the check is over.
        // ###########################################################################################
        public static bool IsProgress(string? status)
        {
            return !string.IsNullOrWhiteSpace(status) &&
                   !status.Contains("up to date", StringComparison.OrdinalIgnoreCase) &&
                   !status.TrimStart().StartsWith("Update complete", StringComparison.OrdinalIgnoreCase) &&
                   !status.TrimStart().StartsWith("Sync complete", StringComparison.OrdinalIgnoreCase);
        }

        // ###########################################################################################
        // The outcome of a finished check, for the line under the tooltip's idle text. A negative
        // count is a check that could not finish (no data folder, or no manifest).
        // ###########################################################################################
        public static string LastCheck(int changedCount, DateTime finishedAt)
        {
            string at = finishedAt.ToString("HH:mm", CultureInfo.InvariantCulture);

            return changedCount switch
            {
                < 0 => $"Last check at {at} failed - see the log",
                0 => $"Last check at {at}: all data files were up to date",
                1 => $"Last check at {at}: [1] file updated",
                _ => $"Last check at {at}: [{changedCount}] files updated",
            };
        }

        private static string OneLine(string text)
        {
            return text
                .Replace("\r", string.Empty, StringComparison.Ordinal)
                .Replace("\n", " ", StringComparison.Ordinal)
                .Trim();
        }
    }
}
