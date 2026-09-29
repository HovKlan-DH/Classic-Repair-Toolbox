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
    // The after-timeout sentences are built by WaitWording (CRT.UI), which CRT uses too, so the
    // limit is described the same way in both applications. Each takes what a fresh look at the
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

        public const string CheckingUnusedFiles = "Checking every file against every workbook...";

        public const string ReadingListing = "Reading CRT's drop-down lists...";

        public static string OpeningSubmission(string systemId) => $"Opening the submission to {systemId}...";

        public static string ReadingSystem(string systemId) => $"Reading {systemId}...";

        // The file tree (2026-09-28): working out a submission's, and fetching one file to open.
        public static string ReadingSubmissionFiles(string systemId) =>
            $"Working out {systemId}'s files after approving - every file in its BETA folder is checked...";

        public static string OpeningFile(string fileName) => $"Fetching {fileName} to open it...";

        public static string ReadingProductionPlan(string systemId) =>
            $"Working out what publishing {systemId} to production would copy...";

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

        // ---- After the limit: things with nothing to check --------------------------------------

        // A reset code is mailed or it is not; there is nothing to read back.
        public static string ResetCodeNoAnswer =>
            $"The server did not answer within {WaitWording.Limit}. A code may still arrive - wait a few minutes before asking again.";

        public static string PasswordNoAnswer =>
            $"The server did not answer within {WaitWording.Limit}. Your new password may have been set - try signing in with it, and ask for a new code only if that fails.";

        public static string InvitationNoAnswer =>
            $"The server did not answer within {WaitWording.Limit}. Your account may have been created - try signing in with the address and password you chose before using the code again.";

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

        public static string PublishAfterTimeout(string systemId, bool? stillInBeta) =>
            WaitWording.AfterTimeout(
                stillInBeta is bool waiting ? !waiting : null,
                $"{systemId} is published to production",
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

        public static string PlacementAfterTimeout(string systemId, bool? saved) =>
            WaitWording.AfterTimeout(saved, $"{systemId}'s place is saved", $"{systemId}'s place is not saved yet");

        // Adding, removing, inviting, withdrawing - `done` read back from the system's detail.
        public static string MaintainersAfterTimeout(bool? done, string doneClause, string notDoneClause) =>
            WaitWording.AfterTimeout(done, doneClause, notDoneClause);
    }
}
