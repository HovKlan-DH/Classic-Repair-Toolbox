using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using Handlers.DataHandling;

namespace Handlers.MaintainerHandling
{
    // ###########################################################################################
    // WHERE A SYSTEM IS, AS THREE STAGES under its name on the Systems screen (owner request,
    // 2026-10-04: "I see this system here ... where does this sit now, as I do not think it is in
    // BETA nor stable?" - and the answer chosen: a stage line, with a BETA / Stable switch on Board
    // data and Files).
    //
    //   Submitted - its newest submission: number, state in CRT's own words, when it was sent or
    //               decided, and how many there have been;
    //   BETA      - its revision there, and whether it waits to be pushed to stable - or not there;
    //   Stable    - its revision there and when it was published - or not there, or not known when
    //               the server has no stable source to look in.
    //
    // A system "not published yet" then says why at a glance: its submission's state, and nothing
    // in either tree. Pure, so what the maintainer reads is tested.
    // ###########################################################################################
    public enum SystemStageState
    {
        // There: a submission, or the board in that tree.
        Reached,

        // Not there - said, never left blank.
        NotThere,

        // Not known yet (the submissions not read) or not knowable (no stable source).
        NotKnown
    }

    public sealed record SystemStage(string Label, string Value, string? Detail, SystemStageState State);

    public static class SystemStagesDisplay
    {
        public const string SubmittedLabel = "Submitted";

        public const string BetaLabel = "BETA";

        public const string StableLabel = "Stable";

        public const string NotThere = "Not there";

        // The three, left to right. `submissions` is the system's detail's - null until it is read.
        public static IReadOnlyList<SystemStage> For(SystemOverviewEntry system, IReadOnlyList<SystemSubmissionEntry>? submissions)
        {
            ArgumentNullException.ThrowIfNull(system);

            return [SystemStagesDisplay.Submitted(submissions), SystemStagesDisplay.Beta(system), SystemStagesDisplay.Stable(system)];
        }

        // ###########################################################################################
        // The NEWEST submission - the one that says where the latest work is. "#5 Not accepted",
        // "decided 21 Sep 2026 - the latest of 2".
        // ###########################################################################################
        public static SystemStage Submitted(IReadOnlyList<SystemSubmissionEntry>? submissions)
        {
            if (submissions is null)
                return new SystemStage(SystemStagesDisplay.SubmittedLabel, string.Empty, null, SystemStageState.NotKnown);

            SystemSubmissionEntry? latest = submissions
                .OrderByDescending(submission => submission.CreatedUtc)
                .ThenByDescending(submission => submission.Id)
                .FirstOrDefault();

            if (latest is null)
                return new SystemStage(SystemStagesDisplay.SubmittedLabel, "No submissions", null, SystemStageState.NotThere);

            string when = latest.DecidedUtc is DateTimeOffset decided
                ? $"decided {SubmissionReceiptPresenter.FormatDate(decided)}"
                : $"sent {SubmissionReceiptPresenter.FormatDate(latest.CreatedUtc)}";

            string of = submissions.Count > 1
                ? $" - the latest of {submissions.Count.ToString(CultureInfo.InvariantCulture)}"
                : string.Empty;

            return new SystemStage(
                SystemStagesDisplay.SubmittedLabel,
                $"#{latest.Id.ToString(CultureInfo.InvariantCulture)} {SubmissionReceiptPresenter.DescribeState(latest.State)}",
                when + of,
                SystemStageState.Reached);
        }

        public static SystemStage Beta(SystemOverviewEntry system)
        {
            ArgumentNullException.ThrowIfNull(system);

            if (!system.InBeta)
                return new SystemStage(SystemStagesDisplay.BetaLabel, SystemStagesDisplay.NotThere, null, SystemStageState.NotThere);

            return new SystemStage(
                SystemStagesDisplay.BetaLabel,
                SystemStagesDisplay.Revision(system.BetaRevision, "In BETA"),
                system.IsAwaitingProduction ? $"ahead of stable - waiting under {MaintainerScreenWording.BetaQueueQuoted}" : null,
                SystemStageState.Reached);
        }

        public static SystemStage Stable(SystemOverviewEntry system)
        {
            ArgumentNullException.ThrowIfNull(system);

            return system.InProduction switch
            {
                true => new SystemStage(
                    SystemStagesDisplay.StableLabel,
                    SystemStagesDisplay.Revision(system.ProductionRevision, "In the stable source"),
                    system.ProductionPublishedUtc is DateTimeOffset published
                        ? $"published {SubmissionReceiptPresenter.FormatDate(published)}"
                        : null,
                    SystemStageState.Reached),
                false => new SystemStage(SystemStagesDisplay.StableLabel, SystemStagesDisplay.NotThere, null, SystemStageState.NotThere),
                _ => new SystemStage(SystemStagesDisplay.StableLabel, "Not known", "this server has no stable source to look in", SystemStageState.NotKnown)
            };
        }

        // "Revision 2026-September-25", or the plain fact for a board with no revision on record.
        private static string Revision(string? revision, string withoutOne) =>
            string.IsNullOrWhiteSpace(revision) ? withoutOne : $"Revision {revision.Trim()}";
    }
}
