using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // THE CREDIT FOR A SUBMISSION'S WORK IS A ROW OF "Board: Credits" (owner request, 2026-11-15: "remove it from this modal, and
    // have this directly done in the "Board: Credits" data, so the user can see how it will show ... have it mandatory, that the
    // contributor has created a new line in the "Board: Credits" with the changes done by the user ... as soon as the user does a
    // modification for an existing board, then it will populate a new entry"; asked the same day: "CRT's Submit only", a new
    // board "Its starting rows", and for no name "Write "Unknown""). The Submit dialog asked for it until then.
    //
    // For a draft of a PUBLISHED board:
    //
    //   - the credits it sends with its work are those the published board does not have (CreditOrder.NotPublished) - the rows
    //     "Board: Credits" frames as sent with the submission;
    //   - the first change CreditSuggestions finds worth a credit adds ONE such row in the very same step (Follow, which
    //     DraftSession runs after every edit, so one Undo takes both back): "What" and "Which part of it" the kinds of work done,
    //     within their limits (Offer, CreditLimits), the name and contact CRT remembers, the main contribution empty - required,
    //     so Submit waits for it (BoardRowRequired);
    //   - while its "What" and "Which part of it" are as this rule wrote them, they follow the work as it grows; and the row goes
    //     again when no work is left and nothing was written in it;
    //   - a row of the contributor's own is a credit sent too, and none is added beside it. None is added back by the edit that
    //     took the sent credits away either: that edit says the contributor wants none (below), and the credit comes back only
    //     when asked for ("Add a credit for my changes") or once that wish ends. Before Submit (Ensure) it is put back only
    //     where nothing said none is wanted - a credit taken away by an edit that left no work, say, and work done again since.
    //
    // A NEW board has no published board: its starting rows (NewBoardCredits) are its credits, every one of them sent, and
    // nothing is added here. A draft whose work deserves no credit (the credits alone, the KiCad data, a calibration, an order -
    // CreditSuggestions) needs none.
    //
    // *** A CONTRIBUTOR MAY WANT NO CREDIT AT ALL (owner request, 2026-11-24: "I want a possibility for the user to opt-out of
    // the credit, so if the user does not want to be creditted, then this is fine ... I do not want to have a new line
    // generated, as there could come too many lines then with minor changes"; and "A credit should still default to be
    // mandatory"). *** The credit is still added with the first change and still required; taking it away - "No credit for my
    // changes" on it, "Remove", or its row deleted in the table - says none is wanted, and then none is added again, at the next
    // change or before Submit (NoCreditAfter, FollowWork; the wish kept per draft in its marker - DraftSession.NoCreditForWork).
    // The submission is then described by its changes (SubmissionSummaryText). "Add a credit for my changes" ends it.
    //
    // PURE: boards in, a board out - the very same object when nothing changes.
    // ###########################################################################################
    public static class SubmissionCredit
    {
        private const string More = ", ...";

        // ###########################################################################################
        // The board after an edit, `before` to `after`, with its credit for the work following it (above). `published` is the
        // board the draft is judged against, null for a new board: `after` as it is.
        //
        // *** THE BOARD AS PUBLISHED NOW - THE DRAFT'S BASE ONLY WHEN THAT CANNOT BE READ (code reviews, 2026-11-16). *** What a
        // submission changes, as Submit's summary (SentWith) and the server's review measure it. Judged against the base, a
        // credit an earlier submission of this very draft sent, published since, still counted as sent - the next work got no
        // credit of its own - and with only the credits judged against the board now, the work of that published submission
        // still counted as work: every later edit added a credit again for work already published.
        //
        // `noCredit`: the contributor wants none for this work (NoCreditAfter) - none is added, though the work needs one.
        // ###########################################################################################
        public static BoardData Follow(BoardData? published, BoardData before, BoardData after, string? name, string? contact, bool noCredit = false)
        {
            ArgumentNullException.ThrowIfNull(before);
            ArgumentNullException.ThrowIfNull(after);

            if (published is null || ReferenceEquals(before, after))
            {
                return after;
            }

            IReadOnlyList<CreditEntry> creditsPublished = published.Credits;
            IReadOnlyList<CreditSuggestion> workAfter = SubmissionCredit.WorkOf(published, after);
            IReadOnlyList<CreditSuggestion> workBefore = SubmissionCredit.WorkOf(published, before);
            IReadOnlySet<CreditEntry> sentAfter = CreditOrder.NotPublished(after.Credits, creditsPublished);

            // The row this rule wrote and nobody retyped: a credit sent whose "What" and "Which part of it" are what it offered for
            // the board as it was before this edit.
            CreditEntry? written = workBefore.Count == 0
                ? null
                : SubmissionCredit.WrittenRow(after.Credits, sentAfter, SubmissionCredit.Offer(workBefore, name, contact));

            if (workAfter.Count == 0)
            {
                return written is not null && SubmissionCredit.NothingWrittenIn(written, name, contact)
                    ? BoardPartEdits.WithCredits(after, after.Credits.Where(credit => !ReferenceEquals(credit, written)).ToList())
                    : after;
            }

            CreditEntry offer = SubmissionCredit.Offer(workAfter, name, contact);

            if (written is not null)
            {
                return SubmissionCredit.SameWhat(written, offer)
                    ? after
                    : BoardPartEdits.WithCredits(after, after.Credits.Select(credit => ReferenceEquals(credit, written) ? SubmissionCredit.WithWhatOf(credit, offer) : credit).ToList());
            }

            if (noCredit || sentAfter.Count > 0 || CreditOrder.NotPublished(before.Credits, creditsPublished).Count > 0)
            {
                return after;
            }

            return BoardPartEdits.WithCredits(after, [.. after.Credits, offer]);
        }

        // ###########################################################################################
        // What a draft session runs after each edit (DraftSessions): whether the contributor wants no credit after it
        // (NoCreditAfter), and the board with its credit following the work by that (Follow). `noCredit`: what was wanted before.
        // ###########################################################################################
        public static DraftFollowUp FollowWork(BoardData? published, BoardData before, BoardData after, string? name, string? contact, bool noCredit)
        {
            bool none = SubmissionCredit.NoCreditAfter(published, before, after, noCredit);

            return new DraftFollowUp(SubmissionCredit.Follow(published, before, after, name, contact, none), none);
        }

        // ###########################################################################################
        // WHETHER THE CONTRIBUTOR WANTS NO CREDIT FOR THE WORK after an edit, `before` to `after` (owner request, 2026-11-24) -
        // `noCredit` what was wanted before it. A credit is wanted by default; the wish for none is said by the credits sent
        // being taken away, and lasts only as long as there is something it is about:
        //
        //   - an edit leaving no credit sent where there was one, while there is work: none is wanted;
        //   - a credit sent again - "Add a credit for my changes", or one added and typed: one is wanted again;
        //   - no work left, or work begun anew from none (the changes undone or published, and new ones since): the wish was
        //     about work that is gone, and the new work is credited as any is.
        //
        // Never for a new board (`published` null): its starting rows are its credits, every one of them sent.
        //
        // *** KEPT IN THE DRAFT'S HISTORY (code review, 2026-11-24). *** Each step holds the wish with its content
        // (DraftSession), so an undo of "Add a credit for my changes" takes the credit away AND puts the wish back - left at "one
        // wanted", the next change added the credit again unasked. A credit sent still outweighs it everywhere (SendsNoCredit,
        // Follow and Ensure add nothing beside a credit sent).
        // ###########################################################################################
        public static bool NoCreditAfter(BoardData? published, BoardData before, BoardData after, bool noCredit)
        {
            ArgumentNullException.ThrowIfNull(before);
            ArgumentNullException.ThrowIfNull(after);

            if (published is null ||
                SubmissionCredit.WorkOf(published, after).Count == 0 ||
                SubmissionCredit.WorkOf(published, before).Count == 0 ||
                CreditOrder.NotPublished(after.Credits, published.Credits).Count > 0)
            {
                return false;
            }

            return noCredit || CreditOrder.NotPublished(before.Credits, published.Credits).Count > 0;
        }

        // ###########################################################################################
        // Whether `board`'s work goes without a credit because the contributor wants none (`noCredit`, the draft's) - what
        // "Board: Credits" says, with "Add a credit for my changes" beside it. False with no work, or a credit sent.
        // ###########################################################################################
        public static bool SendsNoCredit(BoardData? published, BoardData board, bool noCredit)
        {
            ArgumentNullException.ThrowIfNull(board);

            return noCredit && published is not null &&
                   SubmissionCredit.WorkOf(published, board).Count > 0 &&
                   CreditOrder.NotPublished(board.Credits, published.Credits).Count == 0;
        }

        // Whether `board` holds work worth a credit against `published` (CreditSuggestions) - null, a new board, none: its starting
        // rows are its credits. What "No credit for my changes" is offered by: with no such work there is no credit for the
        // contributor's changes to go without (code review, 2026-11-24).
        public static bool HasWorkWorthACredit(BoardData? published, BoardData board)
        {
            ArgumentNullException.ThrowIfNull(board);

            return published is not null && SubmissionCredit.WorkOf(published, board).Count > 0;
        }

        // ###########################################################################################
        // `board` with its credit for the work, when the work needs one and it has none - what Submit does first, so a credit
        // taken away is back for the contributor to fill in. The same board otherwise, and while the contributor wants no credit
        // (`noCredit` - NoCreditAfter). `published` as for Follow.
        // ###########################################################################################
        public static BoardData Ensure(BoardData? published, BoardData board, string? name, string? contact, bool noCredit = false)
        {
            ArgumentNullException.ThrowIfNull(board);

            if (published is null || noCredit)
            {
                return board;
            }

            IReadOnlyList<CreditSuggestion> work = SubmissionCredit.WorkOf(published, board);

            return work.Count == 0 || CreditOrder.NotPublished(board.Credits, published.Credits).Count > 0
                ? board
                : BoardPartEdits.WithCredits(board, [.. board.Credits, SubmissionCredit.Offer(work, name, contact)]);
        }

        // ###########################################################################################
        // The credit offered for `work`: "What" its categories, "Which part of it" its parts - each once, in the order found, joined
        // by ", " - within their limits (CreditLimits), and `name` with `contact`. Categories too many for a line are "Board data",
        // which covers them; parts too many end in ", ...". No main contribution and no date: the contributor writes the one, the
        // server sets the other as it publishes (CreditOrder).
        // ###########################################################################################
        public static CreditEntry Offer(IReadOnlyList<CreditSuggestion> work, string? name, string? contact)
        {
            ArgumentNullException.ThrowIfNull(work);

            List<string> categories = SubmissionCredit.Each(work.Select(suggestion => suggestion.Category));
            string category = string.Join(", ", categories);

            if (category.Length > CreditLimits.Category)
            {
                category = categories.FirstOrDefault(listed => string.Equals(listed, CreditSuggestions.BoardData, StringComparison.OrdinalIgnoreCase))
                    ?? CreditSuggestions.BoardData;
            }

            return new CreditEntry
            {
                Category = category,
                SubCategory = SubmissionCredit.Fit(SubmissionCredit.Each(work.SelectMany(suggestion => suggestion.SubCategory.Split(", "))), CreditLimits.SubCategory),
                NameOrHandle = SubmissionCredit.WithinLimit(name, CreditLimits.NameOrHandle),
                Contact = SubmissionCredit.WithinLimit(contact, CreditLimits.Contact),
            };
        }

        // ###########################################################################################
        // A remembered name or contact as the credit takes it - left out when longer than its limit (code review, 2026-11-16):
        // one remembered before the limits (the Submit dialog's free box) made the credit CRT added an error the contributor
        // never typed. Never cut: a cut contact is another address. Left out, a name is asked for as any empty one is.
        // ###########################################################################################
        private static string WithinLimit(string? remembered, int limit)
        {
            string value = remembered?.Trim() ?? string.Empty;

            return value.Length <= limit ? value : string.Empty;
        }

        // ###########################################################################################
        // What the credits sent with the work say it was - their main contributions, one after the other: what the submission's
        // summary is made of (SubmissionSummaryText). `sent`: the credits known to be sent (KnownSent).
        // ###########################################################################################
        public static string MainContributionsOf(IEnumerable<CreditEntry> sent) =>
            string.Join(
                "\n",
                sent
                    .Select(credit => RichNote.NormalisedMarkup(credit.MainContribution))
                    .Where(contribution => contribution.Length > 0));

        // The credits `board` sends with its work - those `published` does not have, in the board's order; every one of a new
        // board's (`published` null).
        public static IReadOnlyList<CreditEntry> SentWith(BoardData board, BoardData? published)
        {
            ArgumentNullException.ThrowIfNull(board);

            IReadOnlySet<CreditEntry> sent = CreditOrder.NotPublished(board.Credits, published?.Credits);

            return board.Credits.Where(sent.Contains).ToList();
        }

        // ###########################################################################################
        // The credits a Submit sends, as far as can be told - what its summary is made of and whose name is remembered: those the
        // published board as read has not; a new board's (`isNewBoard`, `published` null) every one. A published board that could
        // not be read says nothing of which are sent - every credit would be, other people's words in the summary and another
        // person's name remembered (code review, 2026-11-16) - so none: the summary is then the changes in words.
        // ###########################################################################################
        public static IReadOnlyList<CreditEntry> KnownSent(bool isNewBoard, BoardData board, BoardData? published) =>
            isNewBoard || published is not null ? SubmissionCredit.SentWith(board, published) : [];

        // ###########################################################################################
        // The credits whose name and contact CRT remembers after a Submit, for the next draft's credit: those sent, once the
        // server accepted them - none for a send cancelled, failed or refused - and only when they name ONE person (code review,
        // 2026-11-16). Nothing says which credit is the contributor's own: a new board's starting rows credit whoever drew each
        // schematic image, and a credit added for somebody else stood first - remembered, that person's name and contact were
        // the next draft's credit. So a new board's credits teach nothing, and a published board's only when every one sent
        // names the same person, as the work's credit alone does ("Unknown", nobody, aside).
        // ###########################################################################################
        public static IReadOnlyList<CreditEntry> ToRemember(SubmissionResult? result, bool isNewBoard, IReadOnlyList<CreditEntry> sent)
        {
            ArgumentNullException.ThrowIfNull(sent);

            if (result is not { IsAccepted: true } || isNewBoard)
            {
                return [];
            }

            List<CreditEntry> named = sent
                .Where(credit => !string.IsNullOrWhiteSpace(credit.NameOrHandle) &&
                                 !string.Equals(credit.NameOrHandle.Trim(), CreditSuggestions.UnknownName, StringComparison.OrdinalIgnoreCase))
                .ToList();

            int people = named
                .Select(credit => string.Join('\u001F', credit.NameOrHandle.Trim(), credit.Contact?.Trim() ?? string.Empty).ToUpperInvariant())
                .Distinct(StringComparer.Ordinal)
                .Count();

            return people == 1 ? named : [];
        }

        // ---- what the work is ---------------------------------------------------------------------

        // CreditSuggestions.For worked out once per board: an edit asks for the board before it, which the edit before asked for
        // as the board after it.
        private static readonly ConditionalWeakTable<BoardData, WorkRead> Work = new();

        private sealed record WorkRead(BoardData Published, IReadOnlyList<CreditSuggestion> Suggestions);

        private static IReadOnlyList<CreditSuggestion> WorkOf(BoardData published, BoardData board)
        {
            if (SubmissionCredit.Work.TryGetValue(board, out WorkRead? read) && ReferenceEquals(read.Published, published))
            {
                return read.Suggestions;
            }

            List<CreditSuggestion> suggestions = CreditSuggestions.For(published, board);
            SubmissionCredit.Work.AddOrUpdate(board, new WorkRead(published, suggestions));

            return suggestions;
        }

        private static CreditEntry? WrittenRow(IReadOnlyList<CreditEntry> credits, IReadOnlySet<CreditEntry> sent, CreditEntry offered) =>
            credits.FirstOrDefault(credit => sent.Contains(credit) && SubmissionCredit.SameWhat(credit, offered));

        private static bool SameWhat(CreditEntry first, CreditEntry second) =>
            string.Equals(first.Category?.Trim(), second.Category?.Trim(), StringComparison.OrdinalIgnoreCase) &&
            string.Equals(first.SubCategory?.Trim(), second.SubCategory?.Trim(), StringComparison.OrdinalIgnoreCase);

        // Nothing but what this rule put in it: no main contribution, the name and contact as remembered - as Offer took them.
        private static bool NothingWrittenIn(CreditEntry credit, string? name, string? contact) =>
            RichNote.NormalisedMarkup(credit.MainContribution).Length == 0 &&
            string.Equals(credit.NameOrHandle?.Trim() ?? string.Empty, SubmissionCredit.WithinLimit(name, CreditLimits.NameOrHandle), StringComparison.OrdinalIgnoreCase) &&
            string.Equals(credit.Contact?.Trim() ?? string.Empty, SubmissionCredit.WithinLimit(contact, CreditLimits.Contact), StringComparison.OrdinalIgnoreCase);

        private static CreditEntry WithWhatOf(CreditEntry credit, CreditEntry offer) => new()
        {
            Category = offer.Category,
            SubCategory = offer.SubCategory,
            NameOrHandle = credit.NameOrHandle,
            Contact = credit.Contact,
            MainContribution = credit.MainContribution,
            Date = credit.Date,
        };

        // Each value once, trimmed, in the order found - capitals no difference.
        private static List<string> Each(IEnumerable<string?> values) =>
            values
                .Select(value => value?.Trim() ?? string.Empty)
                .Where(value => value.Length > 0)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();

        // As many of `parts` as a line of `limit` holds, ", ..." after them when some are left out.
        private static string Fit(IReadOnlyList<string> parts, int limit)
        {
            string all = string.Join(", ", parts);

            if (all.Length <= limit)
            {
                return all;
            }

            var taken = new List<string>();

            foreach (string part in parts)
            {
                if (string.Join(", ", [.. taken, part]).Length + SubmissionCredit.More.Length > limit)
                {
                    break;
                }

                taken.Add(part);
            }

            return taken.Count == 0 ? string.Empty : string.Join(", ", taken) + SubmissionCredit.More;
        }
    }
}
