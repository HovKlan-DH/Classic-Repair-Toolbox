using CRT.Server.Handlers.Accounts;
using Handlers.DataHandling;

namespace CRT.Server.Handlers.Submissions
{
    // ###########################################################################################
    // WHAT HAS HAPPENED TO A SYSTEM, NEWEST FIRST (owner request, 2026-09-27: "I would like to see
    // the date, newest first, to understand what has happened to a system").
    //
    // Two sources, merged on their dates:
    //   - the system's SUBMISSIONS: each one's sending, and its decision (the state it was left in
    //     and who decided it), from the rows themselves;
    //   - the AUDIT TRAIL rows naming the system (a promotion, a push-back, a pool change, an
    //     invitation, its place in the drop-down lists) or one of its submissions (an amendment).
    //
    // Only a decision that ENDED somewhere is a line of its own: merged (into BETA), rejected,
    // changes requested, replaced. A push-back leaves its submissions pending again - that is the
    // push-back's own line - and a first approval of two publishes nothing.
    //
    // Pure, so what the history says is tested; SystemOverviewFlow reads the two sources.
    // ###########################################################################################
    public static class SystemHistoryRules
    {
        // Plenty to see what happened lately, and a bounded answer however old the system is.
        public const int Limit = 100;

        // The audit subject a submission is recorded under (AmendSubmissionFlow writes it so).
        public static string SubmissionSubject(long submissionId) => $"#{submissionId}";

        private static readonly HashSet<string> ShownActions = new(StringComparer.Ordinal)
        {
            SystemHistoryEvents.Amended,
            SystemHistoryEvents.PublishedToProduction,
            SystemHistoryEvents.FoundInProduction,
            SystemHistoryEvents.PushedBack,
            SystemHistoryEvents.RejectedFromBeta,
            SystemHistoryEvents.MaintainerAdded,
            SystemHistoryEvents.MaintainerRemoved,
            SystemHistoryEvents.Invited,
            SystemHistoryEvents.InvitationWithdrawn,
            SystemHistoryEvents.InvitationAccepted,
            SystemHistoryEvents.Placed,

            // The contributor discarded their own draft after sending (2026-09-28).
            SystemHistoryEvents.DraftDiscarded
        };

        private static readonly HashSet<string> DecidedStates = new(StringComparer.Ordinal)
        {
            SubmissionState.Merged,
            SubmissionState.Rejected,
            SubmissionState.ChangesRequested,
            SubmissionState.Withdrawn
        };

        public static IReadOnlyList<SystemHistoryEntry> Build(
            IReadOnlyList<SystemSubmissionRecord> submissions,
            IReadOnlyDictionary<long, AccountRecord> named,
            IReadOnlyList<AuditEntry> audit)
        {
            ArgumentNullException.ThrowIfNull(submissions);
            ArgumentNullException.ThrowIfNull(named);
            ArgumentNullException.ThrowIfNull(audit);

            var entries = new List<SystemHistoryEntry>();

            foreach (SystemSubmissionRecord record in submissions)
            {
                SubmissionRecord submission = record.Submission;

                entries.Add(new SystemHistoryEntry(
                    submission.CreatedUtc,
                    SystemHistoryEvents.Sent,
                    submission.AccountId is long sender && named.TryGetValue(sender, out AccountRecord? account)
                        ? account.DisplayName
                        : SystemHistoryRules.Blank(submission.ContactEmail),
                    submission.Id,
                    SystemHistoryRules.Blank(submission.Summary)));

                if (submission.DecidedUtc is DateTimeOffset decided && DecidedStates.Contains(submission.State))
                {
                    entries.Add(new SystemHistoryEntry(
                        decided,
                        SystemHistoryEvents.Decided,
                        record.DecidedByAccountId is long by && named.TryGetValue(by, out AccountRecord? decider)
                            ? decider.DisplayName
                            : null,
                        submission.Id,
                        submission.State,
                        SystemHistoryRules.Blank(submission.DecisionComment)));
                }
            }

            foreach (AuditEntry row in audit)
            {
                if (!ShownActions.Contains(row.Action))
                    continue;

                entries.Add(new SystemHistoryEntry(
                    row.AtUtc,
                    row.Action,
                    SystemHistoryRules.Blank(row.ActorLabel),
                    SystemHistoryRules.SubmissionOf(row.Subject),
                    SystemHistoryRules.DetailOf(row)));
            }

            // Newest first; OrderByDescending is stable, so two events at one instant keep the order
            // they were added in reverse - a decision above its own sending.
            return entries
                .Select((entry, index) => (entry, index))
                .OrderByDescending(item => item.entry.AtUtc)
                .ThenByDescending(item => item.index)
                .Take(SystemHistoryRules.Limit)
                .Select(item => item.entry)
                .ToList();
        }

        // ###########################################################################################
        // The part of an audit row's detail a person reads. Pool rows are written "account 7
        // (anna@example.com)" - the address is the part; an amendment "Commodore/C64/250407:
        // amendment 2, 3 file(s)" - the part after the system; a push-back "<kind>; 3 file(s)
        // restored, ..." - the part after its kind, which is a code word. Anything else as written.
        // ###########################################################################################
        public static string? DetailOf(AuditEntry row)
        {
            ArgumentNullException.ThrowIfNull(row);

            string detail = row.Detail?.Trim() ?? string.Empty;

            if (detail.Length == 0)
                return null;

            switch (row.Action)
            {
                case SystemHistoryEvents.MaintainerAdded:
                case SystemHistoryEvents.MaintainerRemoved:
                case SystemHistoryEvents.InvitationAccepted:
                    int open = detail.IndexOf('(');
                    int close = detail.LastIndexOf(')');
                    return open >= 0 && close > open + 1 ? detail[(open + 1)..close] : detail;

                case SystemHistoryEvents.Amended:
                    int colon = detail.IndexOf(": ", StringComparison.Ordinal);
                    return colon >= 0 ? detail[(colon + 2)..] : detail;

                case SystemHistoryEvents.PushedBack:
                case SystemHistoryEvents.RejectedFromBeta:
                    int semicolon = detail.IndexOf("; ", StringComparison.Ordinal);
                    return semicolon >= 0 ? detail[(semicolon + 2)..] : detail;

                default:
                    return detail;
            }
        }

        // "#41" -> 41; a system id -> null.
        private static long? SubmissionOf(string? subject) =>
            subject is not null && subject.StartsWith('#') && long.TryParse(subject.AsSpan(1), out long id) ? id : null;

        private static string? Blank(string? text) =>
            string.IsNullOrWhiteSpace(text) ? null : text.Trim();
    }
}
