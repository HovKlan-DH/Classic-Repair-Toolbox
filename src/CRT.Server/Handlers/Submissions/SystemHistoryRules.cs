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
        // The system's whole history (owner request, 2026-10-04: "it should show the full history of
        // what has happened with this board") - a bound no board comes near, so the answer stays
        // bounded however old the system is.
        public const int Limit = 1000;

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
            SystemHistoryEvents.DraftDiscarded,

            // The administrator deleted the system (2026-10-03) - seen only by a system created
            // again under the same id, whose history then starts with it.
            SystemHistoryEvents.Deleted
        };

        private static readonly HashSet<string> DecidedStates = new(StringComparer.Ordinal)
        {
            SubmissionState.Merged,
            SubmissionState.Rejected,
            SubmissionState.ChangesRequested,
            SubmissionState.Withdrawn
        };

        // ###########################################################################################
        // `showAddresses` false (owner request, 2026-10-05 - an account that does not maintain the
        // system): NO email address in any line. A person is named by their account where there is
        // one (`named` must then hold whoever did each audited thing, and whoever a pool change
        // names - AccountNamedBy), a label that is no address is kept ("the contributor"), and
        // anything else is left out, which the tab reads as "somebody".
        // ###########################################################################################
        public static IReadOnlyList<SystemHistoryEntry> Build(
            IReadOnlyList<SystemSubmissionRecord> submissions,
            IReadOnlyDictionary<long, AccountRecord> named,
            IReadOnlyList<AuditEntry> audit,
            bool showAddresses = true)
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
                        : showAddresses ? SystemHistoryRules.Blank(submission.ContactEmail) : null,
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
                    showAddresses ? SystemHistoryRules.Blank(row.ActorLabel) : SystemHistoryRules.ActorWithoutAddress(row, named),
                    SystemHistoryRules.SubmissionOf(row.Subject),
                    showAddresses ? SystemHistoryRules.DetailOf(row) : SystemHistoryRules.DetailWithoutAddress(row, named)));
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

        // ###########################################################################################
        // The account a pool change names - "account 7 (anna@example.com)" -> 7 - for a maintainer
        // made, removed, or made by an accepted invitation; null for anything else, or a detail
        // not in that shape.
        // ###########################################################################################
        public static long? AccountNamedBy(AuditEntry row)
        {
            ArgumentNullException.ThrowIfNull(row);

            if (row.Action is not (SystemHistoryEvents.MaintainerAdded or SystemHistoryEvents.MaintainerRemoved or SystemHistoryEvents.InvitationAccepted))
                return null;

            const string Prefix = "account ";
            string detail = row.Detail?.Trim() ?? string.Empty;

            if (!detail.StartsWith(Prefix, StringComparison.Ordinal))
                return null;

            string digits = new(detail[Prefix.Length..].TakeWhile(char.IsAsciiDigit).ToArray());

            return long.TryParse(digits, System.Globalization.NumberStyles.None, System.Globalization.CultureInfo.InvariantCulture, out long id) ? id : null;
        }

        // Who did an audited thing, without an address: the name on their account, else the label
        // when it is no address ("the contributor"), else nobody.
        private static string? ActorWithoutAddress(AuditEntry row, IReadOnlyDictionary<long, AccountRecord> named)
        {
            if (row.ActorAccountId is long id && named.TryGetValue(id, out AccountRecord? account))
                return SystemHistoryRules.Blank(account.DisplayName);

            string? label = SystemHistoryRules.Blank(row.ActorLabel);

            return label is null || label.Contains('@') ? null : label;
        }

        // An audit row's detail without an address: a pool change names its person by their
        // account; an invitation's detail IS an address, so it says nothing; the rest are as
        // DetailOf has them, none of which carries an address.
        private static string? DetailWithoutAddress(AuditEntry row, IReadOnlyDictionary<long, AccountRecord> named)
        {
            switch (row.Action)
            {
                case SystemHistoryEvents.MaintainerAdded:
                case SystemHistoryEvents.MaintainerRemoved:
                case SystemHistoryEvents.InvitationAccepted:
                    return SystemHistoryRules.AccountNamedBy(row) is long id && named.TryGetValue(id, out AccountRecord? account)
                        ? SystemHistoryRules.Blank(account.DisplayName)
                        : null;

                case SystemHistoryEvents.Invited:
                case SystemHistoryEvents.InvitationWithdrawn:
                    return null;

                default:
                    string? detail = SystemHistoryRules.DetailOf(row);
                    return detail is not null && detail.Contains('@') ? null : detail;
            }
        }

        // "#41" -> 41; a system id -> null.
        private static long? SubmissionOf(string? subject) =>
            subject is not null && subject.StartsWith('#') && long.TryParse(subject.AsSpan(1), out long id) ? id : null;

        private static string? Blank(string? text) =>
            string.IsNullOrWhiteSpace(text) ? null : text.Trim();
    }
}
