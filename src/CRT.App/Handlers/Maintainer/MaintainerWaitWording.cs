using System;
using System.Globalization;
using CRT;
using Handlers.DataHandling;

namespace Handlers.MaintainerHandling
{
    // ###########################################################################################
    // What the "please wait" overlay says while the maintainer waits for the server, and what they
    // are told when that wait ran into the two-minute limit (owner decision, 2026-09-28: "I want
    // this method everywhere in the entire project where there is a Wait ... it must be solid in
    // validating if it did finish").
    //
    // Pure and tested, like every presenter behind the Maintainer tab. Each sentence names WHAT is happening -
    // "Working..." was the line that started this, and said nothing.
    //
    // The after-timeout sentences are built by WaitWording, which the rest of CRT uses too, so the
    // limit is described the same way on every screen. Each takes what a fresh look at the
    // server found: finished, not (yet), or null when that look failed too.
    // ###########################################################################################
    public static class MaintainerWaitWording
    {
        public const string SigningIn = "Signing in...";

        public const string SigningOut = "Signing out...";

        public const string ReadingQueue = "Reading the review queue...";

        public const string ReadingSubmissionAgain = "Reading the submission again...";

        public const string AskingForResetCode = "Asking for a password reset code...";

        public const string SettingPassword = "Setting your new password...";

        public const string AcceptingInvitation = "Creating your account from the invitation...";

        // The "Your account" window (2026-10-03).
        public const string SavingName = "Saving your name...";
        public const string SendingEmailCode = "Mailing a code to the new address...";
        public const string ChangingEmail = "Changing your email address...";
        public const string ChangingPassword = "Changing your password...";

        public const string CheckingUnusedFiles = "Checking every file against every workbook...";

        // Rebuilding dataChecksums.json for both trees (2026-10-01): the server hashes every file
        // in each one, so this is seconds rather than instant.
        public const string RebuildingManifests = "Hashing every data file and rebuilding the manifests...";

        public const string ReadingListing = "Reading CRT's drop-down lists...";

        // Account > Maintainers (2026-10-04): every account, for the "choose somebody" list.
        public const string ReadingAccounts = "Reading the accounts...";

        // Account > Order of systems (2026-10-04): the new order written into both sources' lists.
        public const string SavingSystemOrder = "Saving the order of the drop-down lists in BETA and the stable source...";

        public static string OpeningSubmission(string systemId) => $"Opening the submission to {systemId}...";

        public static string ReadingSystem(string systemId) => $"Reading {systemId}...";

        // A system's Board data and Files views (2026-10-03).
        public static string ReadingSystemTable(string systemId) => $"Reading {systemId}'s board from BETA...";

        public static string ReadingSystemFiles(string systemId) => $"Listing {systemId}'s files in BETA...";

        // The file tree (2026-09-28): working out a submission's, and fetching one file to open.
        public static string ReadingSubmissionFiles(string systemId) =>
            $"Working out {systemId}'s files after approving - every file in its BETA folder is checked...";

        public static string OpeningFile(string fileName) => $"Fetching {fileName} to open it...";

        public static string ReadingProductionPlan(string systemId) =>
            $"Working out what publishing {systemId} to the stable source would copy...";

        public static string ReadingRollbackPlan(string systemId) =>
            $"Working out what pushing {systemId} back to the queue would do...";

        public static string RemovingUnusedFiles(int count) =>
            count == 1 ? "Removing 1 unused file..." : string.Create(CultureInfo.InvariantCulture, $"Removing {count} unused files...");

        public static string SavingPlacement(string systemId) =>
            $"Saving where {systemId} goes in CRT's drop-down lists...";

        public static string AddingMaintainer(string who, string systemId) => $"Adding {who} as a maintainer of {systemId}...";

        public static string RemovingMaintainer(string who, string systemId) => $"Removing {who} as a maintainer of {systemId}...";

        public static string Inviting(string email, string systemId) => $"Inviting {email} to maintain {systemId}...";

        public static string WithdrawingInvitation(string email) => $"Withdrawing the invitation to {email}...";

        // Deleting a system (2026-10-03): the plan reads every workbook in both trees.
        public static string ReadingDeletionPlan(string systemId) => $"Working out what deleting {systemId} would remove...";

        public static string DeletingSystem(string systemId) => $"Deleting {systemId} from both data sources and the database...";

        public const string ReadingSystems = "Reading the systems...";

        // Account > Reset contribution data and Account > API usage (2026-10-04).
        public const string ReadingResetCounts = "Counting what a reset would delete...";

        public const string ResettingData = "Deleting every submission, account, maintainer and the history...";

        public const string ReadingApiUsage = "Reading which CRT versions call which route...";

        // ---- After the limit: things with nothing to check --------------------------------------

        // A reset code is mailed or it is not; there is nothing to read back.
        public static string ResetCodeNoAnswer =>
            $"The server did not answer within {WaitWording.Limit}. A code may still arrive - wait a few minutes before asking again.";

        public static string PasswordNoAnswer =>
            $"The server did not answer within {WaitWording.Limit}. Your new password may have been set - try signing in with it, and ask for a new code only if that fails.";

        // The "Your account" window's two changes with nothing to read back: a code is mailed or it
        // is not, and a password cannot be asked for (2026-10-03).
        public static string EmailCodeNoAnswer =>
            $"The server did not answer within {WaitWording.Limit}. A code may still arrive at the new address - wait a few minutes before asking again.";

        public static string NewPasswordNoAnswer =>
            $"The server did not answer within {WaitWording.Limit}. Your password may have been changed - if CRT asks you to sign in, try the new one first.";

        // ###########################################################################################
        // The window's other two changes, judged by reading the account again (2026-10-03): `now`
        // is what the server holds, null when that could not be read.
        // ###########################################################################################
        public static string NameAfterTimeout(string wanted, AccountAnswer? now)
        {
            if (now is null)
                return WaitWording.AfterTimeout(null, string.Empty, string.Empty);

            return string.Equals(now.DisplayName, wanted, StringComparison.Ordinal)
                ? WaitWording.AfterTimeout(true, $"your name is now {wanted}.", string.Empty)
                : WaitWording.AfterTimeout(false, string.Empty, $"your name is still {now.DisplayName}.");
        }

        // `before`: the address when the code was sent back. Any other address now means the code
        // went through - the code alone says which address it was for.
        public static string EmailAfterTimeout(string before, AccountAnswer? now)
        {
            if (now is null)
                return WaitWording.AfterTimeout(null, string.Empty, string.Empty);

            return string.Equals(now.Email, before, StringComparison.Ordinal)
                ? WaitWording.AfterTimeout(false, string.Empty, $"your address is still {before}.")
                : WaitWording.AfterTimeout(true, $"your email address is now {now.Email}.", string.Empty);
        }

        public static string InvitationNoAnswer =>
            $"The server did not answer within {WaitWording.Limit}. Your account may have been created - try signing in with the address the invitation was sent to and the password you chose - it is filled in - before using the code again.";

        // ---- After the limit: things read back from the server ----------------------------------

        // ###########################################################################################
        // A decision that ran into the limit, after reading the submission again. `stateBefore` is
        // what it was when the button was pressed, `stateNow` what the server says now (null when
        // that could not be read). Finished means it moved to what this decision makes it.
        //
        // *** "STILL WAITING" IS ONLY SAID WHEN NOTHING MOVED. *** Somebody else can decide the same
        // submission meanwhile; it then reads what it became, in CRT's own words for states
        // (SubmissionReceiptPresenter.DescribeState), so both applications name it alike.
        // ###########################################################################################
        public static string DecisionAfterTimeout(ReviewDecisionKind kind, string? stateBefore, string? stateNow)
        {
            if (stateNow is null)
                return WaitWording.AfterTimeout(null, string.Empty, string.Empty);

            bool moved = !string.Equals(stateBefore, stateNow, StringComparison.Ordinal);

            string? done = (kind, stateNow) switch
            {
                (ReviewDecisionKind.Approve, "merged") => "the submission is published to BETA.",
                (ReviewDecisionKind.Approve, "approved") => "your approval is recorded.",
                (ReviewDecisionKind.Reject, "rejected") => "the submission is rejected and the contributor has been told why.",
                (ReviewDecisionKind.RequestChanges, "changes_requested") => "the submission is returned to the contributor for changes.",
                _ => null
            };

            if (moved && done is not null)
                return WaitWording.AfterTimeout(true, done, string.Empty);

            return WaitWording.AfterTimeout(
                false,
                string.Empty,
                moved
                    ? $"the submission now reads \"{SubmissionReceiptPresenter.DescribeState(stateNow)}\""
                    : "the submission is still waiting");
        }

        // A save of the table: finished when the submission's version moved past the one it was
        // opened at. Null when the version could not be read.
        public static string SaveAfterTimeout(int versionBefore, int? versionNow) =>
            WaitWording.AfterTimeout(
                versionNow is int now ? now > versionBefore : null,
                "your changes are saved",
                "your changes are not saved yet");

        // ###########################################################################################
        // A change sent from a system's table (2026-10-03): finished when the queue now holds a
        // submission of it - this system, with the description that was sent, waiting for review.
        // The number is named, since that is where it waits. Null when the queue could not be read.
        // ###########################################################################################
        // ###########################################################################################
        // A change published from a system's table, after no answer: what the system's submissions
        // say about it (SystemSections.FindSent) - in BETA, made but waiting in the queue, or not
        // there at all. `read` false: the system could not be read to look.
        // ###########################################################################################
        public static string SystemEditAfterTimeout(bool read, SystemSubmissionEntry? found) =>
            WaitWording.AfterTimeout(
                read ? found is not null : null,
                found is not null && SystemSections.ReachedBeta(found)
                    ? $"your change was published to BETA as submission #{found.Id}"
                    : $"your change was saved as submission #{found?.Id}, but not published to BETA - it waits under {MaintainerScreenWording.ContributorQueueQuoted}",
                "your change is not among the system's submissions, so it was not published");

        public static string PublishAfterTimeout(string systemId, bool? stillInBeta) =>
            WaitWording.AfterTimeout(
                stillInBeta is bool waiting ? !waiting : null,
                $"{systemId} is published to the stable source",
                $"{systemId} is still waiting in BETA");

        public static string PushBackAfterTimeout(string systemId, bool? stillInBeta) =>
            WaitWording.AfterTimeout(
                stillInBeta is bool waiting ? !waiting : null,
                $"{systemId} is pushed back to the queue",
                $"{systemId} is still in BETA");

        public static string RejectAfterTimeout(string systemId, bool? stillInBeta) =>
            WaitWording.AfterTimeout(
                stillInBeta is bool waiting ? !waiting : null,
                $"{systemId} is rejected and out of BETA",
                $"{systemId} is still in BETA");

        public static string RemovalAfterTimeout(int? stillThere) =>
            WaitWording.AfterTimeout(
                stillThere is int left ? left == 0 : null,
                "the files are removed",
                stillThere == 1 ? "1 of them is still there" : $"{stillThere} of them are still there");

        // `stillListed`: whether the systems list, read again, still holds it. A delete that finished
        // part-way leaves it listed too, and says so when Delete is pressed again.
        public static string DeleteAfterTimeout(string systemId, bool? stillListed) =>
            WaitWording.AfterTimeout(
                stillListed is bool listed ? !listed : null,
                $"{systemId} is deleted",
                $"{systemId} is still there");

        public static string PlacementAfterTimeout(string systemId, bool? saved) =>
            WaitWording.AfterTimeout(saved, $"{systemId}'s place is saved", $"{systemId}'s place is not saved yet");

        // ###########################################################################################
        // *** A REBUILD THAT TIMED OUT IS NOT "TRY AGAIN" WITHOUT A WORD (code review, 2026-10-01).
        // *** It WRITES both manifests, and the server carries on past the two minutes - it may be
        // waiting for an approval's publish lock, then hashing. Unlike the others above there is no
        // route to look the result up with, so this says so - and that pressing again is safe: a
        // rebuild only ever writes the manifest the files on disk describe, so twice is the same as
        // once.
        // ###########################################################################################
        public static string RebuildAfterTimeout =>
            $"The server did not answer within {WaitWording.Limit}, and whether the manifests were rebuilt " +
            "cannot be checked from here - it may still be working on it. Pressing the button again is safe: " +
            "it rebuilds them once more from the files as they are.";

        // Adding, removing, inviting, withdrawing - `done` read back from the system's detail.
        public static string MaintainersAfterTimeout(bool? done, string doneClause, string notDoneClause) =>
            WaitWording.AfterTimeout(done, doneClause, notDoneClause);
    }
}
