using System;
using System.Collections.Generic;
using System.Linq;

namespace CRT.Server.Handlers.Submissions
{
    // ###########################################################################################
    // A NEWER SUBMISSION FROM THE SAME CONTRIBUTOR REPLACES THE OLDER ONE (owner decision,
    // 2026-09-26): "if it is from same person, then the newest one always wins".
    //
    // Every submission is the contributor's WHOLE draft of a board, and the draft stays on their
    // machine after it is sent, so a second submission normally holds everything the first did.
    // Left in the queue, the older one is at best work a maintainer does twice - and at worst,
    // approved after the newer one, it publishes the older content back over it.
    //
    // Which submissions a newly queued one replaces (ReplacedBy), all of these holding:
    //
    //   - the SAME SYSTEM;
    //   - the SAME CONTRIBUTOR: the same signed-in account, or - for the ordinary contributor, who
    //     has none - the same contact email. *** The email is not verified ***, so anyone who knows
    //     a contributor's address and board can replace their waiting submission with their own;
    //     the project owner accepted that (2026-09-26) - the replacing one still has to pass review;
    //   - OLDER (a lower id - created earlier), so two sends finishing out of order never withdraw
    //     the newer; both then stay, which is harmless;
    //   - still WAITING ('pending') and UNTOUCHED: a maintainer's amendment (their table edits,
    //     which the contributor's draft never had) or the first of two approvals ('approved') keeps
    //     it - "IF this is the case, it is valid and it should probably stay". The amendment is
    //     checked by the store, inside the transaction that withdraws it
    //     (ISubmissionStore.WithdrawReplacedAsync), since a maintainer may be saving one right now.
    //
    // Withdrawn, not deleted: the contributor's CRT reads it as "Replaced by a newer submission"
    // with ReplacedComment under it, and its bytes are collected like any retired submission's.
    // ###########################################################################################
    public static class SubmissionReplacementRules
    {
        // What the contributor reads under the replaced submission - the decision comment, the one
        // channel an anonymous contributor has. CRT shows no submission numbers, so it names none.
        public const string ReplacedComment =
            "Replaced by the newer submission you sent for this board - that one is waiting for review now.";

        public static IReadOnlyList<SubmissionRecord> ReplacedBy(SubmissionRecord arrived, IEnumerable<SubmissionRecord> waiting)
        {
            ArgumentNullException.ThrowIfNull(arrived);
            ArgumentNullException.ThrowIfNull(waiting);

            return waiting
                .Where(older =>
                    older.Id < arrived.Id &&
                    older.State == SubmissionState.Pending &&
                    string.Equals(older.SystemId, arrived.SystemId, StringComparison.Ordinal) &&
                    SubmissionReplacementRules.IsSameContributor(older, arrived))
                .OrderBy(older => older.Id)
                .ToList();
        }

        // ###########################################################################################
        // The same person: both signed in as the same account, or both anonymous with the same
        // contact email (trimmed, any case - "Dennis@Example.com" is the same mailbox). A signed-in
        // submission and an anonymous one are never the same person here: a signed-in submission
        // stores no contact email to compare.
        // ###########################################################################################
        public static bool IsSameContributor(SubmissionRecord a, SubmissionRecord b)
        {
            ArgumentNullException.ThrowIfNull(a);
            ArgumentNullException.ThrowIfNull(b);

            if (a.AccountId is not null || b.AccountId is not null)
                return a.AccountId == b.AccountId;

            string first = a.ContactEmail?.Trim() ?? string.Empty;
            string second = b.ContactEmail?.Trim() ?? string.Empty;

            return first.Length > 0 && string.Equals(first, second, StringComparison.OrdinalIgnoreCase);
        }
    }
}
