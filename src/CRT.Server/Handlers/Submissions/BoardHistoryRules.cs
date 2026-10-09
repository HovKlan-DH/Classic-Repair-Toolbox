using CRT.Server.Handlers.Accounts;
using Handlers.DataHandling;

namespace CRT.Server.Handlers.Submissions
{
    // ###########################################################################################
    // WHAT HAS HAPPENED TO A BOARD, NEWEST FIRST (owner request, 2026-09-27: "I would like to see
    // the date, newest first, to understand what has happened to a system").
    //
    // Two sources, merged on their dates:
    //   - the board's SUBMISSIONS: each one's sending, and its decision (the state it was left in
    //     and who decided it), from the rows themselves;
    //   - the AUDIT TRAIL rows naming the board (a promotion, a push-back, a pool change, an
    //     invitation, its place in the drop-down lists) or one of its submissions (an amendment).
    //
    // Only a decision that ENDED somewhere is a line of its own: merged (into BETA), rejected,
    // changes requested, replaced. A push-back leaves its submissions pending again - that is the
    // push-back's own line - and a first approval of two publishes nothing.
    //
    // Pure, so what the history says is tested; BoardOverviewFlow reads the two sources.
    // ###########################################################################################
    public static class BoardHistoryRules
    {
        // The board's whole history (owner request, 2026-10-04: "it should show the full history of
        // what has happened with this board") - a bound no board comes near, so the answer stays
        // bounded however old the board is.
        public const int Limit = 1000;

        // The audit subject a submission is recorded under (AmendSubmissionFlow writes it so).
        public static string SubmissionSubject(long submissionId) => $"#{submissionId}";

        private static readonly HashSet<string> ShownActions = new(StringComparer.Ordinal)
        {
            BoardHistoryEvents.Amended,
            BoardHistoryEvents.PublishedToProduction,
            BoardHistoryEvents.FoundInProduction,
            BoardHistoryEvents.PushedBack,
            BoardHistoryEvents.RejectedFromBeta,
            BoardHistoryEvents.MaintainerAdded,
            BoardHistoryEvents.MaintainerRemoved,
            BoardHistoryEvents.Invited,
            BoardHistoryEvents.InvitationWithdrawn,
            BoardHistoryEvents.InvitationAccepted,
            BoardHistoryEvents.Placed,

            // The contributor discarded their own draft after sending (2026-09-28).
            BoardHistoryEvents.DraftDiscarded,

            // The administrator deleted the board (2026-10-03) - seen only by a board created
            // again under the same id, whose history then starts with it.
            BoardHistoryEvents.Deleted
        };

        private static readonly HashSet<string> DecidedStates = new(StringComparer.Ordinal)
        {
            SubmissionState.Merged,
            SubmissionState.Rejected,
            SubmissionState.ChangesRequested,
            SubmissionState.Withdrawn
        };

        // ###########################################################################################
        // *** A PERSON IS NAMED BY THE NAME ON THEIR ACCOUNT (owner request, 2026-10-09: "Please use
        // name (if available, otherwise mail address) for all maintainers"). *** Whoever did an
        // audited thing, and whoever a pool change names, read by their account's name wherever
        // `named` holds it - so it must hold every such account (AccountNamedBy) - and by the
        // address only when there is no name. "dennis@... made a maintainer, by dennis@..." read as
        // nobody could tell who was who.
        //
        // `showAddresses` false (owner request, 2026-10-05 - an account that does not maintain the
        // board): NO email address in any line. A label that is no address is kept ("the
        // contributor"), and anything else is left out, which the tab reads as "somebody".
        // ###########################################################################################
        public static IReadOnlyList<BoardHistoryEntry> Build(
            IReadOnlyList<BoardSubmissionRecord> submissions,
            IReadOnlyDictionary<long, AccountRecord> named,
            IReadOnlyList<AuditEntry> audit,
            bool showAddresses = true)
        {
            ArgumentNullException.ThrowIfNull(submissions);
            ArgumentNullException.ThrowIfNull(named);
            ArgumentNullException.ThrowIfNull(audit);

            var entries = new List<BoardHistoryEntry>();

            foreach (BoardSubmissionRecord record in submissions)
            {
                SubmissionRecord submission = record.Submission;

                entries.Add(new BoardHistoryEntry(
                    submission.CreatedUtc,
                    BoardHistoryEvents.Sent,
                    submission.AccountId is long sender && named.TryGetValue(sender, out AccountRecord? account)
                        ? account.DisplayName
                        : showAddresses ? BoardHistoryRules.Blank(submission.ContactEmail) : null,
                    submission.Id,
                    BoardHistoryRules.Blank(submission.Summary)));

                if (submission.DecidedUtc is DateTimeOffset decided && DecidedStates.Contains(submission.State))
                {
                    entries.Add(new BoardHistoryEntry(
                        decided,
                        BoardHistoryEvents.Decided,
                        record.DecidedByAccountId is long by && named.TryGetValue(by, out AccountRecord? decider)
                            ? decider.DisplayName
                            : null,
                        submission.Id,
                        submission.State,
                        BoardHistoryRules.Blank(submission.DecisionComment)));
                }
            }

            foreach (AuditEntry row in audit)
            {
                if (!ShownActions.Contains(row.Action))
                    continue;

                entries.Add(new BoardHistoryEntry(
                    row.AtUtc,
                    row.Action,
                    BoardHistoryRules.ActorName(row, named) ??
                        (showAddresses ? BoardHistoryRules.Blank(row.ActorLabel) : BoardHistoryRules.ActorWithoutAddress(row)),
                    BoardHistoryRules.SubmissionOf(row.Subject),
                    BoardHistoryRules.NamedPerson(row, named) ??
                        (showAddresses ? BoardHistoryRules.DetailOf(row) : BoardHistoryRules.DetailWithoutAddress(row))));
            }

            // Newest first; OrderByDescending is stable, so two events at one instant keep the order
            // they were added in reverse - a decision above its own sending.
            return entries
                .Select((entry, index) => (entry, index))
                .OrderByDescending(item => item.entry.AtUtc)
                .ThenByDescending(item => item.index)
                .Take(BoardHistoryRules.Limit)
                .Select(item => item.entry)
                .ToList();
        }

        // ###########################################################################################
        // The part of an audit row's detail a person reads. Pool rows are written "account 7
        // (anna@example.com)" - the address is the part; an amendment "Commodore/C64/250407:
        // amendment 2, 3 file(s)" - the part after the board; a push-back "<kind>; 3 file(s)
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
                case BoardHistoryEvents.MaintainerAdded:
                case BoardHistoryEvents.MaintainerRemoved:
                case BoardHistoryEvents.InvitationAccepted:
                    int open = detail.IndexOf('(');
                    int close = detail.LastIndexOf(')');
                    return open >= 0 && close > open + 1 ? detail[(open + 1)..close] : detail;

                case BoardHistoryEvents.Amended:
                    int colon = detail.IndexOf(": ", StringComparison.Ordinal);
                    return colon >= 0 ? detail[(colon + 2)..] : detail;

                case BoardHistoryEvents.PushedBack:
                case BoardHistoryEvents.RejectedFromBeta:
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

            if (row.Action is not (BoardHistoryEvents.MaintainerAdded or BoardHistoryEvents.MaintainerRemoved or BoardHistoryEvents.InvitationAccepted))
                return null;

            const string Prefix = "account ";
            string detail = row.Detail?.Trim() ?? string.Empty;

            if (!detail.StartsWith(Prefix, StringComparison.Ordinal))
                return null;

            string digits = new(detail[Prefix.Length..].TakeWhile(char.IsAsciiDigit).ToArray());

            return long.TryParse(digits, System.Globalization.NumberStyles.None, System.Globalization.CultureInfo.InvariantCulture, out long id) ? id : null;
        }

        // Who did an audited thing, by the name on their account - null when it has none.
        private static string? ActorName(AuditEntry row, IReadOnlyDictionary<long, AccountRecord> named) =>
            row.ActorAccountId is long id && named.TryGetValue(id, out AccountRecord? account)
                ? BoardHistoryRules.Blank(account.DisplayName)
                : null;

        // The person a pool change names, by the name on their account - null for any other row,
        // or an account with no name.
        private static string? NamedPerson(AuditEntry row, IReadOnlyDictionary<long, AccountRecord> named) =>
            BoardHistoryRules.AccountNamedBy(row) is long id && named.TryGetValue(id, out AccountRecord? account)
                ? BoardHistoryRules.Blank(account.DisplayName)
                : null;

        // Who did an audited thing, with no name and no address to give: the label when it is no
        // address ("the contributor"), else nobody.
        private static string? ActorWithoutAddress(AuditEntry row)
        {
            string? label = BoardHistoryRules.Blank(row.ActorLabel);

            return label is null || label.Contains('@') ? null : label;
        }

        // An audit row's detail with no name and no address to give: a pool change says nothing;
        // an invitation's detail IS an address, so it says nothing either; the rest are as DetailOf
        // has them, none of which carries an address.
        private static string? DetailWithoutAddress(AuditEntry row)
        {
            switch (row.Action)
            {
                case BoardHistoryEvents.MaintainerAdded:
                case BoardHistoryEvents.MaintainerRemoved:
                case BoardHistoryEvents.InvitationAccepted:
                    return null;

                case BoardHistoryEvents.Invited:
                case BoardHistoryEvents.InvitationWithdrawn:
                    return null;

                default:
                    string? detail = BoardHistoryRules.DetailOf(row);
                    return detail is not null && detail.Contains('@') ? null : detail;
            }
        }

        // "#41" -> 41; a board id -> null.
        private static long? SubmissionOf(string? subject) =>
            subject is not null && subject.StartsWith('#') && long.TryParse(subject.AsSpan(1), out long id) ? id : null;

        private static string? Blank(string? text) =>
            string.IsNullOrWhiteSpace(text) ? null : text.Trim();
    }
}
