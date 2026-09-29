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
    // submission identically ("Published to BETA source", "Taken back out of BETA - waiting for
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
        // Where the system's data is, in one phrase:
        //
        //   "In BETA and production"           - the ordinary case
        //   "BETA ahead of production"         - published to BETA, waiting for the BETA screen
        //   "In BETA - not in production yet"  - never promoted
        //   "Not published yet"                - a new system whose first submission is still in
        //                                        review, or was turned down
        //
        // Production is only claimed when the server looked (InProduction not null); without a
        // production tree it says "Published", which is all that is known.
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

            if (!system.InBeta)
            {
                return system.InProduction == true
                    ? [new("In production - ", false), new("not in BETA", true)]
                    : [new("Not published yet", false)];
            }

            if (system.IsAwaitingProduction)
            {
                return system.InProduction == true
                    ? [new("BETA ahead of production", true)]
                    : [new("In BETA - ", false), new("not in production yet", true)];
            }

            return system.InProduction switch
            {
                true => [new("In BETA and production", false)],
                false => [new("In BETA - ", false), new("not in production yet", true)],
                _ => [new("Published", false)]
            };
        }

        // ###########################################################################################
        // The grey line under a system in the list: where its data is, how many maintain it, and -
        // only when it is the case - that it is closed to contributions.
        // ###########################################################################################
        public static string ListLine(SystemOverviewEntry system) =>
            SystemsDisplay.Joined(SystemsDisplay.ListLineParts(system));

        // The same line in pieces, the parts still to be done marked - see WhereParts. "Nobody
        // assigned" is one: a system with no maintainer sends everything to the administrator.
        public static IReadOnlyList<StatusPart> ListLineParts(SystemOverviewEntry system)
        {
            ArgumentNullException.ThrowIfNull(system);

            var parts = new List<StatusPart>(SystemsDisplay.WhereParts(system))
            {
                new(" - ", false),
                new(MaintainerAssignmentDisplay.CountPhrase(system.MaintainerCount), system.MaintainerCount == 0)
            };

            if (!system.IsAccepting)
                parts.Add(new(" - closed to contributions", false));

            // How often CRT users look at it (2026-09-27) - absent from an older server.
            if (system.ViewsLast30Days is int views)
                parts.Add(new(" - " + SystemsDisplay.ViewsPhrase(views), false));

            return parts;
        }

        // ###########################################################################################
        // BOARD VIEWS (owner request, 2026-09-27: usage statistics "as systems statics in the
        // 'Systems' menu"). A view is a board on screen in CRT for ten seconds, counted every time
        // (CRT.Data's BoardViewContract); views from CRTs downloading BETA data - mostly maintainers
        // checking their own work - are counted apart, so they cannot make a board look used.
        // ###########################################################################################

        // The list's phrase: "48 views in 30 days", "1 view in 30 days", "no views in 30 days".
        public static string ViewsPhrase(int views) =>
            views switch
            {
                <= 0 => "no views in 30 days",
                1 => "1 view in 30 days",
                _ => $"{views.ToString(CultureInfo.InvariantCulture)} views in 30 days"
            };

        // The section's heading - or, with nothing counted at all, the whole section.
        public static string ViewsHeading(BoardViewStatistics views)
        {
            ArgumentNullException.ThrowIfNull(views);

            return SystemsDisplay.HasNoViews(views)
                ? "No views of this board in CRT counted in the last 12 months."
                : "Views in CRT";
        }

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

        // ###########################################################################################
        // The revisions, under the system's name on the right: "BETA revision 2026-September-20 -
        // production revision 2026-May-14, published there 2026-May-14". Null for a shipped board
        // nothing has been published through yet - it has no revision on record to name.
        // ###########################################################################################
        public static string? Revisions(SystemOverviewEntry system)
        {
            ArgumentNullException.ThrowIfNull(system);

            var parts = new List<string>();

            if (!string.IsNullOrWhiteSpace(system.BetaRevision))
                parts.Add($"BETA revision {system.BetaRevision.Trim()}");

            if (!string.IsNullOrWhiteSpace(system.ProductionRevision))
            {
                string production = $"production revision {system.ProductionRevision.Trim()}";

                if (system.ProductionPublishedUtc is DateTimeOffset published)
                    production += $", published there {SubmissionReceiptPresenter.FormatDate(published)}";

                parts.Add(production);
            }

            if (parts.Count == 0)
                return null;

            string line = string.Join(" - ", parts);
            return char.ToUpperInvariant(line[0]) + line[1..];
        }

        // The maintainers section's heading - and, with nobody assigned, where its submissions go.
        public static string MaintainersHeading(int count) =>
            count == 0
                ? "Nobody maintains this system - its submissions go to the administrator."
                : count == 1 ? "Maintainer" : $"Maintainers ({count.ToString(CultureInfo.InvariantCulture)})";

        // "Anna (anna@example.com)".
        // ###########################################################################################
        // THE SYSTEM'S HISTORY (owner request, 2026-09-27: "I would like to see the date, newest
        // first, to understand what has happened to a system"). One line per event, the DATE FIRST
        // so the column reads down as a timeline, then what happened; a grey line under it says by
        // whom and anything more. A decision is said in CRT's own state words (DescribeState), the
        // ones "My submissions" uses - so "merged" reads "Published to BETA source" here too.
        // ###########################################################################################
        public static string HistoryLine(SystemHistoryEntry entry)
        {
            ArgumentNullException.ThrowIfNull(entry);

            string number = entry.SubmissionId is long id ? $"#{id.ToString(CultureInfo.InvariantCulture)}" : "A submission";
            string detail = entry.Detail?.Trim() ?? string.Empty;

            string what = entry.Event switch
            {
                SystemHistoryEvents.Sent => $"{number} sent",
                SystemHistoryEvents.Decided => $"{number} - {SubmissionReceiptPresenter.DescribeState(detail)}",
                SystemHistoryEvents.Amended => $"{number} changed by a maintainer",
                SystemHistoryEvents.PublishedToProduction => "Published to production",
                SystemHistoryEvents.FoundInProduction => "Found already in production (copied there outside CRT)",
                SystemHistoryEvents.PushedBack => "Pushed back from BETA to the queue",
                SystemHistoryEvents.RejectedFromBeta => "Rejected in BETA and taken out of it",
                SystemHistoryEvents.MaintainerAdded => $"{SystemsDisplay.Or(detail, "Somebody")} made a maintainer",
                SystemHistoryEvents.MaintainerRemoved => $"{SystemsDisplay.Or(detail, "Somebody")} removed as a maintainer",
                SystemHistoryEvents.Invited => $"{SystemsDisplay.Or(detail, "Somebody")} invited to be a maintainer",
                SystemHistoryEvents.InvitationWithdrawn => $"Invitation to {SystemsDisplay.Or(detail, "somebody")} withdrawn",
                SystemHistoryEvents.InvitationAccepted => $"{SystemsDisplay.Or(detail, "Somebody")} accepted the invitation and became a maintainer",
                SystemHistoryEvents.Placed => "Placed in the drop-down lists",
                SystemHistoryEvents.DraftDiscarded => DraftDiscardWording.HistoryWhat(number),
                _ => entry.Event
            };

            return $"{SubmissionReceiptPresenter.FormatDate(entry.AtUtc)} - {what}";
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
                    SystemHistoryEvents.Placed => detail.Length > 0 ? detail : null,
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

        public static string HistoryHeading(int count) =>
            count == 0 ? "Nothing has happened to this system yet" : "History";

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

        // "Anna (anna@example.com)" - or the address alone for a contributor with no account.
        public static string ContributorName(SystemContributorEntry contributor)
        {
            ArgumentNullException.ThrowIfNull(contributor);

            if (string.IsNullOrWhiteSpace(contributor.Name))
                return string.IsNullOrWhiteSpace(contributor.Email) ? "(no address)" : contributor.Email.Trim();

            return SystemsDisplay.NameAndAddress(contributor.Name, contributor.Email);
        }

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

        public static string SubmissionsHeading(int count) =>
            count == 0 ? "No submissions yet." : "Recent submissions";

        // What the contributor wrote about it, or that they wrote nothing.
        public static string SubmissionTitle(SystemSubmissionEntry submission)
        {
            ArgumentNullException.ThrowIfNull(submission);

            return string.IsNullOrWhiteSpace(submission.Summary) ? "(no description given)" : submission.Summary.Trim();
        }

        // ###########################################################################################
        // The grey line under it: "#12 - Published to BETA source - sent 2026-September-25 -
        // anna@example.com". The state in CRT's words (see the header).
        // ###########################################################################################
        public static string SubmissionFooter(SystemSubmissionEntry submission)
        {
            ArgumentNullException.ThrowIfNull(submission);

            var parts = new List<string>
            {
                $"#{submission.Id.ToString(CultureInfo.InvariantCulture)}",
                SubmissionReceiptPresenter.DescribeState(submission.State)
            };

            if (submission.CreatedUtc != default)
                parts.Add($"sent {SubmissionReceiptPresenter.FormatDate(submission.CreatedUtc)}");

            if (!string.IsNullOrWhiteSpace(submission.ContactEmail))
                parts.Add(submission.ContactEmail.Trim());

            return string.Join(" - ", parts);
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
