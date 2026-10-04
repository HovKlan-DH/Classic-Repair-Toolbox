using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using Handlers.DataHandling;

namespace Handlers.MaintainerHandling
{
    // ###########################################################################################
    // How the BETA screen reads - the "Production" window until the four screens replaced the
    // windows, 2026-09-27 (owner request, 2026-09-25: BETA first, then production "only after he
    // has checked that the data looks correct in BETA").
    //
    // Pure, so the words are tested - the same rule ReviewQueueDisplay and
    // MaintainerAssignmentDisplay follow. One rule here is a decision rather than formatting: the
    // "Publish to production" button follows the SERVER's canPublish AND the maintainer's own tick in
    // "I have checked this in CRT with the BETA data" - never either alone (CanPress).
    // ###########################################################################################
    public static class ProductionDisplay
    {
        // "Commodore / C64 / 250407  -  BETA 2026-September-25, production 2026-May-14" - the
        // heading over the plan, where it has the width to wrap.
        public static string SystemLine(ProductionSystemRow system) =>
            $"{ProductionDisplay.SystemName(system)}  -  {ProductionDisplay.SystemStatus(system)}";

        // ###########################################################################################
        // The same, on TWO lines, for the list on the left. One line was cut off at the list's
        // width in the first render - "...never p" - which hid the one fact the row is there to
        // give: how far production is behind.
        // ###########################################################################################
        public static string ListEntry(ProductionSystemRow system) =>
            $"{ProductionDisplay.SystemName(system)}\n{ProductionDisplay.SystemStatus(system)}";

        // ###########################################################################################
        // The grey line under a system in the BETA screen's list (2026-09-27): where BETA and
        // production stand - and, when this account has already approved and it waits for the OTHER
        // approver, that too, since the row is dimmed for it (the queue's own wording).
        // ###########################################################################################
        public static string ListFooter(ProductionSystemRow system)
        {
            ArgumentNullException.ThrowIfNull(system);

            string status = ProductionDisplay.SystemStatus(system);
            status = char.ToUpperInvariant(status[0]) + status[1..];

            return system.AwaitsYou == false ? $"{status} - with the other approver" : status;
        }

        public static string SystemName(ProductionSystemRow system)
        {
            ArgumentNullException.ThrowIfNull(system);

            string name = string.Join(
                " / ",
                new[] { system.Manufacturer, system.Hardware, system.Board }
                    .Where(part => !string.IsNullOrWhiteSpace(part)));

            return name.Length > 0
                ? name
                : string.IsNullOrWhiteSpace(system.SystemId) ? "(unknown system)" : system.SystemId;
        }

        // "BETA 2026-September-25, production 2026-May-14"
        public static string SystemStatus(ProductionSystemRow system)
        {
            ArgumentNullException.ThrowIfNull(system);

            string beta = string.IsNullOrWhiteSpace(system.BetaRevision) ? "BETA ahead" : $"BETA {system.BetaRevision}";

            string production = string.IsNullOrWhiteSpace(system.ProductionRevision)
                ? "never published to the stable source"
                : $"stable {system.ProductionRevision}";

            return $"{beta}, {production}";
        }

        // ###########################################################################################
        // "2 files to add, 1 to replace, 1,180 already the same". Nothing to copy is said as such,
        // because it is a real outcome (production was brought level by hand) and publishing then
        // only records it.
        // ###########################################################################################
        public static string PlanSummary(ProductionPlanView plan)
        {
            ArgumentNullException.ThrowIfNull(plan);

            int added = plan.Files.Count(file => file.Change == PromotionChange.Added);
            int replaced = plan.Files.Count(file => file.Change == PromotionChange.Replaced);
            int removed = plan.Removals?.Files.Count ?? 0;

            string unchanged = plan.UnchangedCount.ToString("N0", CultureInfo.InvariantCulture);

            // Removals are said only when there are any: most promotions remove nothing, and a
            // standing "0 to remove" would be read past on the one that does.
            string removing = removed == 0 ? string.Empty : $", {ProductionDisplay.Count(removed, "file", "to remove")}";

            if (added == 0 && replaced == 0)
                return $"Production already has every file ({unchanged}){removing}. Publishing only records it.";

            return $"{ProductionDisplay.Count(added, "file", "to add")}, " +
                $"{ProductionDisplay.Count(replaced, "file", "to replace")}{removing}, {unchanged} already the same";
        }

        // ###########################################################################################
        // *** WHOSE WORK THIS PROMOTION CARRIES (owner request, 2026-09-27). ***
        //
        // The panel showed the file copy list and nothing else, which is the mechanics of a copy
        // rather than anything a maintainer can judge - "I am not sure if the shown information in
        // the right-side panel is any helpful". A board is hundreds of files and a path cannot say
        // whether the data is right; meanwhile the one thing the button actually does - push named
        // people's accepted work to every CRT user - was nowhere on the screen.
        //
        // So the headline names the count, and each line names the contributor and their own
        // description of what they sent. These are the server's `carrying` facts, the same set it
        // emails once the promotion succeeds.
        // ###########################################################################################
        public static string CarryingHeadline(IReadOnlyList<CarriedSubmission>? carrying)
        {
            int count = carrying?.Count ?? 0;

            // Said even when it is none, because that IS the answer for a board brought level by
            // hand, and silence would read as "the question was not asked".
            return count == 0
                ? "No contributions are waiting to go out with this."
                : count == 1
                    ? "1 contribution goes out to everyone with this:"
                    : $"{count.ToString("N0", CultureInfo.InvariantCulture)} contributions go out to everyone with this:";
        }

        // ###########################################################################################
        // One carried submission: who sent it, what they called it, and when it was accepted into
        // BETA. The comment is the contributor's own - the same text the review queue lists it by,
        // so a maintainer recognises the submission they approved.
        // ###########################################################################################
        public static string CarryingLine(CarriedSubmission submission, DateTimeOffset now)
        {
            ArgumentNullException.ThrowIfNull(submission);

            string who = string.IsNullOrWhiteSpace(submission.ContactEmail)
                ? "(no contact address)"
                : submission.ContactEmail.Trim();

            string comment = string.IsNullOrWhiteSpace(submission.Comment)
                ? "(no description)"
                : submission.Comment.Trim();

            // "accepted 3 days ago". An absent timestamp says nothing rather than guessing - the
            // same rule ReviewQueueDisplay.Waiting follows about a missing date.
            string when = ProductionDisplay.AcceptedAgo(submission.DecidedUtc, now);

            return when.Length == 0
                ? $"{who} - {comment}"
                : $"{who} - {comment} ({when})";
        }

        // ###########################################################################################
        // How long ago a submission was accepted into BETA. TRUNCATED to the coarser unit downwards,
        // as ReviewQueueDisplay.Waiting is: a promotion three weeks behind should not read as "a
        // month", which would make the delay sound worse than it is, nor 29 days as "4 weeks".
        // ###########################################################################################
        public static string AcceptedAgo(DateTimeOffset? decidedUtc, DateTimeOffset now)
        {
            if (decidedUtc is null)
                return string.Empty;

            TimeSpan since = now - decidedUtc.Value;

            // A clock difference between client and server can give a small negative; reading
            // "accepted -3 minutes ago" looks broken.
            if (since < TimeSpan.Zero || since.TotalMinutes < 1)
                return "accepted just now";

            if (since.TotalHours < 1)
                return ProductionDisplay.Ago((int)since.TotalMinutes, "minute");

            if (since.TotalDays < 1)
                return ProductionDisplay.Ago((int)since.TotalHours, "hour");

            return ProductionDisplay.Ago((int)since.TotalDays, "day");
        }

        private static string Ago(int count, string unit) =>
            count == 1
                ? $"accepted 1 {unit} ago"
                : $"accepted {count.ToString(CultureInfo.InvariantCulture)} {unit}s ago";

        // ###########################################################################################
        // *** PUSHING A BOARD BACK TO THE QUEUE (owner decision, 2026-09-27). *** The words have to
        // carry the one thing that makes this different from a per-submission action: a rollback is
        // PER SYSTEM, so it takes back EVERY submission merged since the last promotion. Saying
        // only "this is rolled back" would let a maintainer discard two other contributors' accepted
        // work believing they were returning one.
        //
        // The two kinds say different things because they ARE different operations: a restore puts
        // production's board back, while a system never promoted has nothing to go back to and
        // leaves BETA entirely.
        // ###########################################################################################
        public static string RollBackHeadline(BetaRollbackPlanView? plan) =>
            plan?.Kind == BetaRollbackKind.RemoveFromBeta
                ? "Remove this board from BETA and push it back to the queue?"
                : "Roll this board back to what the stable source has, and push it back to the queue?";

        public static string RollBackExplanation(BetaRollbackPlanView? plan)
        {
            if (plan is null)
                return string.Empty;

            string what = ProductionDisplay.WhatBetaGetsBack(plan);

            int returning = plan.Returning.Count;

            // The sentence that stops a maintainer discarding someone else's work unknowingly.
            string who = returning == 0
                ? "No submission goes back to the queue."
                : returning == 1
                    ? "The submission below goes back to the queue for review."
                    : $"ALL {returning.ToString(CultureInfo.InvariantCulture)} submissions below go back to the queue - a board cannot be " +
                        "rolled back one contribution at a time.";

            return $"{what} {who}";
        }

        // What happens to BETA's data - the same for a push-back and a rejection.
        private static string WhatBetaGetsBack(BetaRollbackPlanView plan) =>
            plan.Kind == BetaRollbackKind.RemoveFromBeta

                // Nothing of it has ever been published, so there is no earlier state to return to.
                ? "Nothing of this system is in the stable source, so its data is removed from BETA entirely."
                : $"BETA goes back to the data the stable source already has: " +
                    $"{ProductionDisplay.Count(plan.Restored.Count, "file", "restored")}, " +
                    $"{ProductionDisplay.Count(plan.Removed.Count, "file", "removed")}.";

        // ###########################################################################################
        // *** REJECTING FROM BETA (owner request, 2026-09-28: "a direct 'Reject' button also - just
        // like the normal queue. Then there is no need to push it back and then reject it"). *** The
        // same rollback, so BETA's data moves exactly as for a push-back - but the submissions are
        // rejected and do not come back to the queue. The same care over names: a rejection takes
        // out EVERY submission merged since the last promotion, so the sentence says ALL of them.
        // ###########################################################################################
        public static string RejectHeadline(BetaRollbackPlanView? plan) =>
            plan?.Kind == BetaRollbackKind.RemoveFromBeta
                ? "Reject this system and remove it from BETA?"
                : "Reject this, and roll the board back to what the stable source has?";

        public static string RejectExplanation(BetaRollbackPlanView? plan)
        {
            if (plan is null)
                return string.Empty;

            int rejected = plan.Returning.Count;

            string who = rejected == 0
                ? "No submission is rejected."
                : rejected == 1
                    ? "The submission below is rejected: its contributor is told why, and it does not come back to the queue."
                    : $"ALL {rejected.ToString(CultureInfo.InvariantCulture)} submissions below are rejected - a board cannot be " +
                        "taken out of BETA one contribution at a time. Each contributor is told why.";

            return $"{ProductionDisplay.WhatBetaGetsBack(plan)} {who}";
        }

        public static string RejectConfirmButton(BetaRollbackPlanView? plan) =>
            plan is not null && plan.Returning.Count > 1
                ? $"Reject all {plan.Returning.Count.ToString(CultureInfo.InvariantCulture)}"
                : "Reject";

        public static string RejectingWait(string systemId) =>
            $"Rejecting {systemId}. BETA's data is being put back as the stable source has it - please wait until it is done.";

        // What a finished rejection says, as RolledBack does for a push-back.
        public static string Rejected(BetaRollbackResult result)
        {
            ArgumentNullException.ThrowIfNull(result);

            return $"Rejected: {ProductionDisplay.FilesMoved(result)}. " +
                $"{ProductionDisplay.Count(result.SubmissionsReturned, "submission", "rejected")}, " +
                "and the contributors have been told.";
        }

        // ###########################################################################################
        // *** ASKED TO REJECT, BUT PUSHED BACK. *** A server older than "Reject" ignores the request's
        // flag and pushes back - the submissions are in the queue again, NOT rejected. Said as it is,
        // with what to do, rather than claiming a rejection that did not happen.
        // ###########################################################################################
        public static string PushedBackInsteadOfRejected(BetaRollbackResult result)
        {
            ArgumentNullException.ThrowIfNull(result);

            return $"The server is older than this CRT and PUSHED {result.SystemId} BACK instead of rejecting it: " +
                $"{ProductionDisplay.FilesMoved(result)}, " +
                $"{ProductionDisplay.Count(result.SubmissionsReturned, "submission", "back in the queue")}. " +
                "Reject it from " + MaintainerScreenWording.ContributorQueueQuoted + ".";
        }

        // ###########################################################################################
        // *** THE SHARED FILES, SAID SEPARATELY (code review, 2026-09-27). *** A rollback puts back
        // the shared files its submissions changed, and those reach EVERY board citing them - a
        // different kind of consequence from this board's own files, so it gets its own sentence
        // rather than disappearing into a count. Null when there are none. A shared file the
        // submissions ADDED is never removed (BetaRollbackPlan), so there is nothing else to say.
        // ###########################################################################################
        public static string? RollBackSharedFiles(BetaRollbackPlanView? plan)
        {
            int restored = plan?.SharedRestored?.Count ?? 0;

            if (restored == 0)
                return null;

            string count = restored == 1
                ? "1 goes back to the stable source's version"
                : $"{restored.ToString(CultureInfo.InvariantCulture)} go back to the stable source's version";

            return $"Shared files, used by every board that cites them: {count}.";
        }

        // Every shared path the confirmation lists under that sentence.
        public static IReadOnlyList<string> RollBackSharedPaths(BetaRollbackPlanView? plan) =>
            [.. plan?.SharedRestored ?? []];

        // The button's own words, so it says which of the two it does rather than a generic verb.
        public static string RollBackConfirmButton(BetaRollbackPlanView? plan) =>
            plan?.Kind == BetaRollbackKind.RemoveFromBeta
                ? "Remove from BETA"
                : "Roll back";

        // ###########################################################################################
        // The "please wait" over the whole window while a system is pushed back or published
        // (owner request, 2026-09-27). Says what is happening, and that it takes a moment - a copy
        // of a whole board, then the lists read again.
        // ###########################################################################################
        public static string PushingBackWait(string systemId) =>
            $"Pushing {systemId} back to the queue. BETA's data is being put back as the stable source has it - please wait until it is done.";

        public static string PublishingWait(string systemId) =>
            $"Publishing {systemId} to the stable source. Its files are being copied from BETA - please wait until it is done.";

        // ###########################################################################################
        // What a finished rollback says. Named counts rather than "done", because the maintainer
        // has just changed what every BETA user downloads and should see how much moved.
        // ###########################################################################################
        public static string RolledBack(BetaRollbackResult result)
        {
            ArgumentNullException.ThrowIfNull(result);

            return $"Pushed back: {ProductionDisplay.FilesMoved(result)}. " +
                $"{ProductionDisplay.Count(result.SubmissionsReturned, "submission", "back in the queue")}, " +
                "and the contributors have been told.";
        }

        // How much of BETA moved - a removal counts only what left, "0 files restored" being noise.
        private static string FilesMoved(BetaRollbackResult result) =>
            result.Kind == BetaRollbackKind.RemoveFromBeta
                ? ProductionDisplay.Count(result.FilesRemoved, "file", "removed from BETA")
                : $"{ProductionDisplay.Count(result.FilesRestored, "file", "restored")}, " +
                    $"{ProductionDisplay.Count(result.FilesRemoved, "file", "removed")}";

        // ###########################################################################################
        // May the button be pressed? The server must say yes AND the maintainer must have ticked the
        // box. The tick is the human half of "only after he has checked it in BETA"; the server's
        // half is the content hash the request carries back.
        // ###########################################################################################
        public static bool CanPress(ProductionPlanView? plan, bool checkedInBeta) =>
            plan is not null && plan.CanPublish && checkedInBeta;

        private static string Count(int count, string noun, string what) =>
            count == 1
                ? $"1 {noun} {what}"
                : $"{count.ToString("N0", CultureInfo.InvariantCulture)} {noun}s {what}";
    }
}
