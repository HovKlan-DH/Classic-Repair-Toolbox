using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using Handlers.DataHandling;

namespace Handlers.MaintainerHandling
{
    // ###########################################################################################
    // How the "Systems" screen reads (owner request, 2026-09-27: "Right-side will show data per
    // system - e.g. who has contributed to it and who is set as maintainer etc.").
    //
    // Pure, so the words are tested - the rule ReviewQueueDisplay and ProductionDisplay follow.
    //
    // *** A SUBMISSION'S STATE IS SAID IN CRT's OWN WORDS. *** The server sends the word its
    // contributor is told, and it is put into words by CRT.Data's SubmissionReceiptPresenter - the
    // very function CRT's "My submissions" uses - so the maintainer and the contributor describe one
    // submission identically ("Published to the BETA source", "Taken back out of BETA - waiting for
    // review again"). CLAUDE.md's "vocabulary the user reads" rule, kept by construction.
    // ###########################################################################################
    public static class SystemsDisplay
    {
        // "Commodore / C64 / 250407" - or the id whole when the parts are missing.
        public static string Name(SystemOverviewEntry system)
        {
            ArgumentNullException.ThrowIfNull(system);

            string name = string.Join(
                " / ",
                new[] { system.Manufacturer, system.Hardware, system.Board }.Where(part => !string.IsNullOrWhiteSpace(part)));

            return name.Length > 0
                ? name
                : string.IsNullOrWhiteSpace(system.SystemId) ? "(unknown system)" : system.SystemId;
        }

        // ###########################################################################################
        // Where the system's data is, in one phrase - and NOTHING in the ordinary case (owner
        // request, 2026-10-03: "There is no need to always see 'In BETA and stable source', as this
        // will be default for all - only show where there is something odd/off"):
        //
        //   ""                                        - in BETA and the stable source, as every
        //                                               system should be; or published where the
        //                                               server has no stable tree to look in
        //   "BETA ahead of the stable source"         - published to BETA, waiting for Beta > Stable
        //   "In BETA - not in the stable source yet"  - never promoted
        //   "In the stable source - not in BETA"      - BETA lacks what the stable source holds
        //   "Not published yet"                       - a new system whose first submission is still
        //                                               in review, or was turned down
        //
        // And, after any of those, what is off about CRT's DROP-DOWN LISTS (owner request,
        // 2026-10-04: "I do not expect there should be cases where something can only be listed in
        // stable? If so, it must be flagged") - see ListingParts.
        //
        // The stable source is only claimed when the server looked (InProduction not null).
        // ###########################################################################################
        public static string Where(SystemOverviewEntry system) =>
            SystemsDisplay.Joined(SystemsDisplay.WhereParts(system));

        // ###########################################################################################
        // Where, in pieces, with the part that still has to be DONE marked (owner request,
        // 2026-09-27: "for the important info, then do bold that ... 'nobody assigned' is important,
        // as that should be done - the same for 'not in production yet'"). Marked: "not in
        // production yet" and "BETA ahead of production" (both wait for "Publish to production"),
        // and "not in BETA" (production holds a board BETA lacks - something to put right). Not
        // marked: "Not published yet", which a turned-down new system reads too, with nothing to do.
        // ###########################################################################################
        public static IReadOnlyList<StatusPart> WhereParts(SystemOverviewEntry system)
        {
            ArgumentNullException.ThrowIfNull(system);

            var parts = new List<StatusPart>(SystemsDisplay.TreeParts(system));
            IReadOnlyList<StatusPart> listing = SystemsDisplay.ListingParts(system);

            if (parts.Count > 0 && listing.Count > 0)
                parts.Add(new(" - ", false));

            parts.AddRange(listing);
            return parts;
        }

        // ###########################################################################################
        // WHAT IS OFF ABOUT THE DROP-DOWN LISTS (2026-10-04), both marked as still to be DONE:
        //
        //   "listed only in the stable source's drop-down list"  - the stable list names it and
        //       BETA's does not: the next promotion would not keep it, and BETA's CRT users cannot
        //       reach it. Not said for a board BETA does not hold at all - "not in BETA" says that.
        //   "missing from the stable source's drop-down list"    - the stable source holds the board
        //       and BETA's list names it, but the stable list does not: stable CRT users cannot
        //       reach a board that is there.
        //
        // Nothing when either list is unknown (null) - never "not listed" without looking. A new
        // system in BETA only, waiting for its first promotion, is listed in BETA alone, as it
        // should be, and says nothing.
        // ###########################################################################################
        public static IReadOnlyList<StatusPart> ListingParts(SystemOverviewEntry system)
        {
            ArgumentNullException.ThrowIfNull(system);

            if (system.ListedInStable == true && system.ListedInBeta == false && !(system.InProduction == true && !system.InBeta))
                return [new("listed only in the stable source's drop-down list", true)];

            if (system.ListedInStable == false && system.ListedInBeta == true && system.InProduction == true)
                return [new("missing from the stable source's drop-down list", true)];

            return [];
        }

        // Where the board's data is - the trees. See WhereParts.
        private static IReadOnlyList<StatusPart> TreeParts(SystemOverviewEntry system)
        {
            if (!system.InBeta)
            {
                return system.InProduction == true
                    ? [new("In the stable source - ", false), new("not in BETA", true)]
                    : [new("Not published yet", false)];
            }

            if (system.IsAwaitingProduction)
            {
                return system.InProduction == true
                    ? [new("BETA ahead of the stable source", true)]
                    : [new("In BETA - ", false), new("not in the stable source yet", true)];
            }

            // The ordinary case says nothing - the list and the panel show only what is off.
            return system.InProduction == false
                ? [new("In BETA - ", false), new("not in the stable source yet", true)]
                : [];
        }

        // ###########################################################################################
        // The grey line under a system in the list: where its data is when that is anything but the
        // ordinary (WhereParts), how many maintain it, and - only when it is the case - that it is
        // closed to contributions. No view count any more (owner request, 2026-10-03: "remove the
        // 'no views in 30 days' from the data in the left-side list") - the views are the
        // system's Statistics view.
        // ###########################################################################################
        public static string ListLine(SystemOverviewEntry system) =>
            SystemsDisplay.Joined(SystemsDisplay.ListLineParts(system));

        // The same line in pieces, the parts still to be done marked - see WhereParts. "Nobody
        // assigned" is one: a system with no maintainer sends everything to the administrator.
        public static IReadOnlyList<StatusPart> ListLineParts(SystemOverviewEntry system)
        {
            ArgumentNullException.ThrowIfNull(system);

            var parts = new List<StatusPart>(SystemsDisplay.WhereParts(system));

            if (parts.Count > 0)
                parts.Add(new(" - ", false));

            parts.Add(new(MaintainerAssignmentDisplay.CountPhrase(system.MaintainerCount), system.MaintainerCount == 0));

            if (!system.IsAccepting)
                parts.Add(new(" - closed to contributions", false));

            return parts;
        }

        // ###########################################################################################
        // BOARD VIEWS (owner request, 2026-09-27: usage statistics "as systems statics in the
        // 'Systems' menu"). A view is a board on screen in CRT for ten seconds, counted every time
        // (CRT.Data's BoardViewContract); views from CRTs downloading BETA data - mostly maintainers
        // checking their own work - are counted apart, so they cannot make a board look used.
        // ###########################################################################################

        // The section's heading - or, with nothing counted at all, the whole section.
        public static string ViewsHeading(BoardViewStatistics views)
        {
            ArgumentNullException.ThrowIfNull(views);

            return SystemsDisplay.HasNoViews(views)
                ? "No views of this board in CRT counted in the last 12 months."
                : "Views in CRT";
        }

        // The Statistics view when the server sent no counts at all - an older server, or counts it
        // could not read. Not "no views", which would be a claim about the board.
        public const string NoViewCountsLine = "The server sent no view counts for this board.";

        public static bool HasNoViews(BoardViewStatistics views) =>
            views.Last365Days <= 0 && views.FromBetaLast30Days <= 0;

        // "[12] in the last 7 days - [48] in 30 days - [310] in 12 months", the numbers bold.
        public static IReadOnlyList<ReviewNoteRun> ViewCountRuns(BoardViewStatistics views)
        {
            ArgumentNullException.ThrowIfNull(views);

            var runs = new List<ReviewNoteRun>();

            SystemsDisplay.AddCount(runs, views.Last7Days, " in the last 7 days - ");
            SystemsDisplay.AddCount(runs, views.Last30Days, " in 30 days - ");
            SystemsDisplay.AddCount(runs, views.Last365Days, " in 12 months");

            return runs;
        }

        // "Most views in the last 12 months: Denmark (120), Germany (60)." - null with no country known.
        public static string? ViewCountriesLine(BoardViewStatistics views)
        {
            ArgumentNullException.ThrowIfNull(views);

            if (views.TopCountries.Count == 0)
                return null;

            return "Most views in the last 12 months: " + string.Join(
                ", ",
                views.TopCountries.Select(country =>
                    $"{country.CountryName} ({country.Views.ToString(CultureInfo.InvariantCulture)})")) + ".";
        }

        // "Not counted above: [3] from CRTs downloading BETA data in the last 30 days." - null with none.
        public static IReadOnlyList<ReviewNoteRun>? ViewBetaRuns(BoardViewStatistics views)
        {
            ArgumentNullException.ThrowIfNull(views);

            if (views.FromBetaLast30Days <= 0)
                return null;

            var runs = new List<ReviewNoteRun> { new("Not counted above: ", IsCount: false) };
            SystemsDisplay.AddCount(runs, views.FromBetaLast30Days, " from CRTs downloading BETA data in the last 30 days.");
            return runs;
        }

        // What a view is, under the numbers.
        public static string ViewExplanation =>
            $"A view is this board on screen in CRT for at least " +
            $"{BoardViewRules.MinimumTimeOnScreen.TotalSeconds.ToString(CultureInfo.InvariantCulture)} seconds, counted every time.";

        private static void AddCount(List<ReviewNoteRun> runs, int count, string after)
        {
            runs.Add(new ReviewNoteRun("[", IsCount: false));
            runs.Add(new ReviewNoteRun(Math.Max(0, count).ToString(CultureInfo.InvariantCulture), IsCount: true));
            runs.Add(new ReviewNoteRun("]" + after, IsCount: false));
        }

        private static string Joined(IEnumerable<StatusPart> parts) => string.Concat(parts.Select(part => part.Text));

        // The maintainers section's heading - and, with nobody assigned, where its submissions go.
        public static string MaintainersHeading(int count) =>
            count == 0
                ? "Nobody maintains this system - its submissions go to the administrator."
                : count == 1 ? "Maintainer" : $"Maintainers ({count.ToString(CultureInfo.InvariantCulture)})";

        // ###########################################################################################
        // ONE EVENT OF THE SYSTEM'S HISTORY (owner request, 2026-09-27: "I would like to see the date,
        // newest first, to understand what has happened to a system"): what happened; a grey line
        // under it says by whom and anything more. The History view puts the date in a column of its
        // own and gathers a submission's events onto one card (SystemHistoryDisplay, 2026-10-04) -
        // these words are for every other event. A decision is said in CRT's own state words
        // (DescribeState), the ones "My submissions" uses - so "merged" reads "Published to the BETA
        // source" here too.
        // ###########################################################################################
        public static string HistoryWhat(SystemHistoryEntry entry)
        {
            ArgumentNullException.ThrowIfNull(entry);

            string number = entry.SubmissionId is long id ? $"#{id.ToString(CultureInfo.InvariantCulture)}" : "A submission";
            string detail = entry.Detail?.Trim() ?? string.Empty;

            string what = entry.Event switch
            {
                SystemHistoryEvents.Sent => $"{number} sent",
                SystemHistoryEvents.Decided => $"{number} - {SubmissionReceiptPresenter.DescribeState(detail)}",
                SystemHistoryEvents.Amended => $"{number} changed by a maintainer",
                SystemHistoryEvents.PublishedToProduction => "Published to the stable source",
                SystemHistoryEvents.FoundInProduction => "Found already in the stable source (copied there outside CRT)",
                SystemHistoryEvents.PushedBack => "Pushed back from BETA to the queue",
                SystemHistoryEvents.RejectedFromBeta => "Rejected in BETA and taken out of it",
                SystemHistoryEvents.MaintainerAdded => $"{SystemsDisplay.Or(detail, "Somebody")} made a maintainer",
                SystemHistoryEvents.MaintainerRemoved => $"{SystemsDisplay.Or(detail, "Somebody")} removed as a maintainer",
                SystemHistoryEvents.Invited => $"{SystemsDisplay.Or(detail, "Somebody")} invited to be a maintainer",
                SystemHistoryEvents.InvitationWithdrawn => $"Invitation to {SystemsDisplay.Or(detail, "somebody")} withdrawn",
                SystemHistoryEvents.InvitationAccepted => $"{SystemsDisplay.Or(detail, "Somebody")} accepted the invitation and became a maintainer",
                SystemHistoryEvents.Placed => "Placed in the drop-down lists",
                SystemHistoryEvents.DraftDiscarded => DraftDiscardWording.HistoryWhat(number),
                SystemHistoryEvents.Deleted => "Deleted from the BETA and stable data and the database",
                _ => entry.Event
            };

            return what;
        }

        // The grey line: who, and what the line itself did not say - a submission's description, a
        // promotion's or push-back's counts, the names a placement gave. Empty when there is nothing.
        public static string HistoryFooter(SystemHistoryEntry entry)
        {
            ArgumentNullException.ThrowIfNull(entry);

            string who = entry.Who?.Trim() ?? string.Empty;
            string detail = entry.Detail?.Trim() ?? string.Empty;

            string? more = entry.Event switch
            {
                SystemHistoryEvents.Sent => detail.Length > 0 ? detail : "(no description given)",
                SystemHistoryEvents.Amended or SystemHistoryEvents.PublishedToProduction or SystemHistoryEvents.FoundInProduction or
                    SystemHistoryEvents.PushedBack or SystemHistoryEvents.RejectedFromBeta or
                    SystemHistoryEvents.Placed or SystemHistoryEvents.Deleted => detail.Length > 0 ? detail : null,
                _ => null
            };

            // The pool rows name the person in the line itself, and an accepted invitation IS its
            // person - "by" would name them twice.
            string? by = who.Length == 0 || entry.Event == SystemHistoryEvents.InvitationAccepted
                ? null
                : entry.Event == SystemHistoryEvents.Sent ? $"from {who}" : $"by {who}";

            return string.Join(" - ", new[] { by, more }.Where(part => !string.IsNullOrWhiteSpace(part)));
        }

        // What a maintainer told the contributor with a decision - null for anything else.
        public static string? HistoryNote(SystemHistoryEntry entry)
        {
            ArgumentNullException.ThrowIfNull(entry);

            return string.IsNullOrWhiteSpace(entry.Note) ? null : $"Told the contributor: {entry.Note.Trim()}";
        }

        private static string Or(string text, string fallback) => text.Length > 0 ? text : fallback;

        // ###########################################################################################
        // An invitation nobody has accepted yet (2026-09-27), under the system's maintainers on the
        // administrator's screen: "Invited: bo@example.com", and under it when it was sent and how
        // long its code works - the one thing that decides whether to send it again.
        // ###########################################################################################
        public static string InvitationLine(MaintainerInvitationEntry invitation)
        {
            ArgumentNullException.ThrowIfNull(invitation);

            return $"Invited: {invitation.Email}";
        }

        public static string InvitationFooter(MaintainerInvitationEntry invitation)
        {
            ArgumentNullException.ThrowIfNull(invitation);

            return $"Sent {SubmissionReceiptPresenter.FormatDate(invitation.InvitedUtc)} - " +
                   $"the code works until {SubmissionReceiptPresenter.FormatDate(invitation.ExpiresUtc)}, " +
                   "and they become a maintainer when they use it";
        }

        public static string MaintainerLine(PoolMaintainerEntry maintainer)
        {
            ArgumentNullException.ThrowIfNull(maintainer);

            return SystemsDisplay.NameAndAddress(maintainer.DisplayName, maintainer.Email);
        }

        public static string ContributorsHeading(int count) =>
            count == 0
                ? "Nobody has contributed to this system through CRT yet."
                : count == 1 ? "Contributor" : $"Contributors ({count.ToString(CultureInfo.InvariantCulture)})";

        // "Anna (anna@example.com)" - or the address alone for a contributor with no account. For a
        // system the account does not maintain (2026-10-05) the server sends no address, so a
        // contributor with no account has nothing to be named by - said as that, not "(no address)",
        // which would claim they gave none.
        public static string ContributorName(SystemContributorEntry contributor, bool addressesHidden = false)
        {
            ArgumentNullException.ThrowIfNull(contributor);

            if (string.IsNullOrWhiteSpace(contributor.Name))
            {
                if (!string.IsNullOrWhiteSpace(contributor.Email))
                    return contributor.Email.Trim();

                return addressesHidden ? "A contributor without an account" : "(no address)";
            }

            return SystemsDisplay.NameAndAddress(contributor.Name, contributor.Email);
        }

        // ###########################################################################################
        // Said above the Contributor and Maintainer views of a system the account does not maintain
        // (owner request, 2026-10-05), so the missing addresses read as a rule rather than a fault.
        // ###########################################################################################
        public const string AddressesHiddenLine = "Email addresses are shown only for the systems you maintain.";

        // ###########################################################################################
        // How a contributor's submissions to this system went, and when they last sent one:
        // "3 accepted, 1 waiting, 1 rejected - last sent 2026-September-25". Only the counts that are
        // not zero - "0 changes requested, 0 rejected" on every line would drown the one that is not.
        // ###########################################################################################
        public static string ContributorRecord(SystemContributorEntry contributor)
        {
            ArgumentNullException.ThrowIfNull(contributor);

            var counts = new List<string>();

            void Add(int count, string what)
            {
                if (count > 0)
                    counts.Add($"{count.ToString(CultureInfo.InvariantCulture)} {what}");
            }

            Add(contributor.Accepted, "accepted");
            Add(contributor.Waiting, "waiting");
            Add(contributor.ChangesRequested, "sent back for changes");
            Add(contributor.Rejected, "rejected");

            var parts = new List<string>();

            if (counts.Count > 0)
                parts.Add(string.Join(", ", counts));

            if (contributor.LastSubmittedUtc is DateTimeOffset last)
                parts.Add($"last sent {SubmissionReceiptPresenter.FormatDate(last)}");

            string line = string.Join(" - ", parts);
            return line.Length == 0 ? line : char.ToUpperInvariant(line[0]) + line[1..];
        }

        // What the contributor wrote about it, or that they wrote nothing.
        public static string SubmissionTitle(SystemSubmissionEntry submission)
        {
            ArgumentNullException.ThrowIfNull(submission);

            return string.IsNullOrWhiteSpace(submission.Summary) ? "(no description given)" : submission.Summary.Trim();
        }

        // ###########################################################################################
        // What the contributor was TOLD, when anything was - a maintainer's reason for a rejection or
        // a change request, why it was pushed back out of BETA, or that a newer submission replaced
        // it. Null when nothing was said. Labelled by who reads it, not who wrote it: the replacement
        // note is the server's, not a maintainer's.
        // ###########################################################################################
        public static string? SubmissionComment(SystemSubmissionEntry submission)
        {
            ArgumentNullException.ThrowIfNull(submission);

            return string.IsNullOrWhiteSpace(submission.DecisionComment)
                ? null
                : $"Told the contributor: {submission.DecisionComment.Trim()}";
        }

        private static string NameAndAddress(string? name, string? email)
        {
            string shown = string.IsNullOrWhiteSpace(name) ? "(no name)" : name.Trim();

            return string.IsNullOrWhiteSpace(email) ? shown : $"{shown} ({email.Trim()})";
        }
    }

    // ###########################################################################################
    // One piece of a status line, and whether it is something still to be DONE - which the screen
    // shows in bold (owner request, 2026-09-27).
    // ###########################################################################################
    public sealed record StatusPart(string Text, bool IsToDo);
}
