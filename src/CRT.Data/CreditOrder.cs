using System;
using System.Collections.Generic;
using System.Linq;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // *** THE BOARD'S CREDITS ARE ITS HISTORY, OLDEST FIRST (owner requests, 2026-11-14: "have the credits now sorted by oldest
    // contribution first ... With this, we get a cronological order"; owner decision: "Always kept in date order"; and the next
    // day: "the existing credits should actually have NO date at all - but all new credits must have it set"). *** A credit's
    // date is when it was published, set by the server as it publishes (PublishMerge.WithCreditsDated, Dated) in the board's
    // own shape, "2026-November-14" (BoardWorkbookStyle.FormatRevisionDate), read by DraftRevisionComparer's one rule. So:
    //
    //   - a credit with NO date in the published data is one from before the history was kept - older than every dated one,
    //     so the undated come FIRST, in the board's order;
    //   - then the dated, by day, credits of one day in the board's order;
    //   - a date that cannot be read (only the table can type one) comes last - and is an error (BoardDataChecks'
    //     "credit.date_unreadable"), so it is never published.
    //
    // In a DRAFT a new credit has no date either, until it is published - so the order is made where every undated credit is an
    // old one: by the server at every publish. A draft keeps the board's order, a credit added standing last ("Board: Credits"
    // has no drag handle), which is where the publish puts it too. A new credit is told from an old undated one by the
    // published board (Dated, NotPublished): each published credit answers for one credit sent, its very values first, then
    // its category, sub-category and name or handle - so the same person credited for the same kind of work again is new.
    // ###########################################################################################
    public static class CreditOrder
    {
        // The credits, oldest first (above). The same list, in the same order, when it already is.
        public static List<CreditEntry> Sorted(IEnumerable<CreditEntry> credits)
        {
            ArgumentNullException.ThrowIfNull(credits);

            return credits
                .Select((credit, index) => (Credit: credit, Index: index, Key: CreditOrder.KeyOf(credit.Date)))
                .OrderBy(item => item.Key.Rank)
                .ThenBy(item => item.Key.Day)
                .ThenBy(item => item.Index)
                .Select(item => item.Credit)
                .ToList();
        }

        // A date given that cannot be read as one - an empty one can (from before the history, or not published yet).
        public static bool IsUnreadable(string? date) =>
            !string.IsNullOrWhiteSpace(date) && !DraftRevisionComparer.TryParse(date.Trim(), out _);

        // ###########################################################################################
        // The credits as a publish on `publishedUtc` writes them: one the published board has keeps the date it was sent with
        // (a maintainer's correction of it), else the date it has there - none, for one from before the history; and every
        // other one, new in this submission, is dated that day, WHATEVER DATE IT WAS SENT WITH. Then oldest first.
        //
        // *** A NEW CREDIT CANNOT BE BACK-DATED (owner decision, 2026-11-16: "Should not be possible to back-date - new
        // contributions, not seen before, must be set when board maintainer pushes this to BETA and STABLE"). *** A date typed
        // on a new credit - in the "Contribute" tab's table, or by a maintainer - kept it until then, sorted among the old
        // history. A promotion to stable copies the date BETA's publish set.
        // ###########################################################################################
        public static List<CreditEntry> Dated(IEnumerable<CreditEntry> credits, IEnumerable<CreditEntry>? published, DateTimeOffset publishedUtc)
        {
            ArgumentNullException.ThrowIfNull(credits);

            string today = BoardWorkbookStyle.FormatRevisionDate(publishedUtc);
            List<CreditEntry> sent = credits.ToList();
            string?[] matched = CreditOrder.PublishedDatesOf(sent, published);

            // A date given is written trimmed, as every reading of it trims (code review, 2026-11-16).
            return CreditOrder.Sorted(sent.Select((credit, index) =>
                matched[index] is not { } date ? credit.WithDate(today)
                : !string.IsNullOrWhiteSpace(credit.Date) ? (credit.Date == credit.Date.Trim() ? credit : credit.WithDate(credit.Date.Trim()))
                : date.Length == 0 ? credit
                : credit.WithDate(date)));
        }

        // ###########################################################################################
        // *** A CREDIT SENT UNDATED THAT THE BOARD HAS PUBLISHED TAKES ITS PUBLISHED DATE (PublishMerge.Build). *** The
        // contributor's draft has its own credit undated after the publish dated it, and a draft made before credits had dates
        // sends every one undated: the review compares them with what the board has, never as changed. A credit the board does
        // not have is undated, for the publish to date (Dated) - a date it was sent with is let go of, as the publish lets go of
        // it (owner decision, 2026-11-16: "Should not be possible to back-date"), so the review never shows one.
        // ###########################################################################################
        public static List<CreditEntry> WithPublishedDates(IEnumerable<CreditEntry> credits, IEnumerable<CreditEntry>? published)
        {
            ArgumentNullException.ThrowIfNull(credits);

            List<CreditEntry> sent = credits.ToList();
            string?[] matched = CreditOrder.PublishedDatesOf(sent, published);

            return sent
                .Select((credit, index) =>
                    matched[index] is not { } date ? (string.IsNullOrWhiteSpace(credit.Date) ? credit : credit.WithDate(string.Empty))
                    : string.IsNullOrWhiteSpace(credit.Date) && date.Length > 0 ? credit.WithDate(date)
                    : credit)
                .ToList();
        }

        // The credits of `credits` the published board does not have - new in a draft, to be dated as they are published (the
        // pane's "Dated when published"). By reference.
        public static IReadOnlySet<CreditEntry> NotPublished(IReadOnlyList<CreditEntry> credits, IEnumerable<CreditEntry>? published)
        {
            ArgumentNullException.ThrowIfNull(credits);

            string?[] matched = CreditOrder.PublishedDatesOf(credits, published);

            return credits.Where((credit, index) => matched[index] is null).ToHashSet<CreditEntry>(ReferenceEqualityComparer.Instance);
        }

        // ###########################################################################################
        // For each credit sent, the date of the published credit that answers for it ("" for one with none), or null for one the
        // board does not have. Each published credit answers for ONE: first a credit with its very values (category,
        // sub-category, name or handle, contact, main contribution), then one with its category, sub-category and name or
        // handle - a credit changed. A credit sent WITH a date answers by its values like any other, so it uses up its own.
        // ###########################################################################################
        // Each pass works out every published credit's identity ONCE, its candidates in the board's order by identity (code
        // review, 2026-11-16): asked for each credit sent, they were worked out again for every one - a join and a note's
        // markup normalised for every pair - and this runs after every draft edit (SubmissionCredit.Follow).
        //
        // *** THIRD, A "Schematic image" CREDIT BY EVERY VALUE BUT WHICH PART OF IT (code review, 2026-11-16) - its image renamed
        // on its card, the credit following (SchematicRename): PublishedCredits.PartlessKeyOf, the check's own rule. *** By its
        // values and by who and what it no longer matched: the author's credit read as new - sent, dated the day of the publish.
        private static string?[] PublishedDatesOf(IReadOnlyList<CreditEntry> sent, IEnumerable<CreditEntry>? published)
        {
            List<CreditEntry> pool = (published ?? []).ToList();
            var used = new bool[pool.Count];
            var dates = new string?[sent.Count];

            foreach (Func<CreditEntry, string?> identity in new Func<CreditEntry, string?>[] { CreditOrder.AllValuesOf, CreditOrder.WhoAndWhatOf, PublishedCredits.PartlessKeyOf })
            {
                var candidates = new Dictionary<string, Queue<int>>(StringComparer.OrdinalIgnoreCase);

                for (int candidate = 0; candidate < pool.Count; candidate++)
                {
                    if (used[candidate] || identity(pool[candidate]) is not { } key)
                        continue;

                    if (!candidates.TryGetValue(key, out Queue<int>? ofKey))
                        candidates[key] = ofKey = new Queue<int>();

                    ofKey.Enqueue(candidate);
                }

                for (int index = 0; index < sent.Count; index++)
                {
                    if (dates[index] is not null)
                        continue;

                    if (identity(sent[index]) is { } wanted && candidates.TryGetValue(wanted, out Queue<int>? ofKey) && ofKey.TryDequeue(out int at))
                    {
                        used[at] = true;
                        dates[index] = pool[at].Date?.Trim() ?? string.Empty;
                    }
                }
            }

            return dates;
        }

        private static string AllValuesOf(CreditEntry credit) =>
            string.Join('\u001F', CreditOrder.WhoAndWhatOf(credit), credit.Contact?.Trim(), RichNote.NormalisedMarkup(credit.MainContribution));

        private static string WhoAndWhatOf(CreditEntry credit) =>
            string.Join('\u001F', credit.Category?.Trim(), credit.SubCategory?.Trim(), credit.NameOrHandle?.Trim());

        // The undated first (from before the history), then the dated by day, then the unreadable.
        private static (int Rank, DateTime Day) KeyOf(string? date) =>
            string.IsNullOrWhiteSpace(date) ? (0, DateTime.MinValue)
            : DraftRevisionComparer.TryParse(date.Trim(), out DateTime day) ? (1, day.Date)
            : (2, DateTime.MinValue);
    }

    // ###########################################################################################
    // *** A VALUE THE SERVER SETS AS IT PUBLISHES - EMPTY IN A DRAFT IS NO CHANGE (2026-11-14). *** A credit's date: added in a
    // draft it has none, and the server dates it at the publish (CreditOrder). Compared as any other value, the draft's own
    // credit read as changed against the published one the moment it was published - "1 row changed" over a draft fully
    // published, and never retired (DraftRetirement). So an EMPTY value on the draft's side of a comparison is no change
    // (BoardDataDiffer.DifferingFields, and the table's cells by their sheet's SetWhenPublishedColumns - BoardWorkbookSchemaTests
    // holds the two together); a value given that differs still is - a maintainer's correction of a date.
    // ###########################################################################################
    [AttributeUsage(AttributeTargets.Property)]
    public sealed class SetWhenPublishedAttribute : Attribute
    {
    }
}
