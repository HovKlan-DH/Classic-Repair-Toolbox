using System.Security.Cryptography;
using System.Text;

namespace CRT.Server.Handlers.Submissions
{
    // ###########################################################################################
    // The pure half of deleting a board (owner request, 2026-10-03) - which submissions are still
    // in play, what the administrator confirmed, and the sentences a refusal is given in. Tested
    // without a tree or a database; BoardDeletionFlow and BoardDeletionFiles do the I/O.
    // ###########################################################################################
    public static class BoardDeletionRules
    {
        // The two trees, as the administrator reads them - the words the Account screen's other
        // tools use ("Both the BETA and the stable data are rebuilt").
        public const string BetaTreeName = "BETA";
        public const string ProductionTreeName = "stable";

        // ###########################################################################################
        // *** STILL IN PLAY: the submissions whose contributors are mailed (owner decision,
        // 2026-10-03: "Delete them and mail the contributors"). *** Waiting for review (pending, or
        // pending again after a push-back - "returned"), approved once and waiting for the second
        // approval, or merged into BETA and not yet in the stable source. Takes the CONTRIBUTOR-
        // FACING state (ProductionPromotionRules.ContributorFacingState), so a merged submission
        // already promoted - "published" - is finished work, not open work.
        // ###########################################################################################
        public static bool IsOpen(string? contributorFacingState) =>
            contributorFacingState is SubmissionState.Pending
                or SubmissionState.Approved
                or SubmissionState.Merged
                or ProductionPromotionRules.ReturnedState;

        // ###########################################################################################
        // *** WHAT THE ADMINISTRATOR WAS SHOWN, AS ONE VALUE. *** Every file of the board's folder
        // in each tree, whether each drop-down list names it, whether it has a database record, and
        // every submission with its state. The delete is refused when the value it computes under
        // the lock differs - so a publish, a promotion, a new submission or a file copied in by hand
        // since the confirmation opened is never deleted unseen.
        //
        // Paths compare ORDINALLY, as the server's case-sensitive trees do; the order given does not
        // matter. Lower-case hex SHA-256.
        // ###########################################################################################
        public static string Fingerprint(
            IEnumerable<string> betaFiles,
            IEnumerable<string> productionFiles,
            bool listedInBeta,
            bool listedInProduction,
            bool hasRecord,
            IEnumerable<(long Id, string State)> submissions)
        {
            ArgumentNullException.ThrowIfNull(betaFiles);
            ArgumentNullException.ThrowIfNull(productionFiles);
            ArgumentNullException.ThrowIfNull(submissions);

            var text = new StringBuilder();

            foreach (string path in betaFiles.Order(StringComparer.Ordinal))
                text.Append("beta\n").Append(path).Append('\n');

            foreach (string path in productionFiles.Order(StringComparer.Ordinal))
                text.Append("stable\n").Append(path).Append('\n');

            text.Append("listed\n").Append(listedInBeta ? '1' : '0').Append(listedInProduction ? '1' : '0').Append('\n');
            text.Append("record\n").Append(hasRecord ? '1' : '0').Append('\n');

            foreach ((long id, string state) in submissions.OrderBy(submission => submission.Id))
                text.Append("submission\n").Append(id).Append('\n').Append(state).Append('\n');

            return Convert.ToHexStringLower(SHA256.HashData(Encoding.UTF8.GetBytes(text.ToString())));
        }

        // ---- Refusals, one sentence each ---------------------------------------------------------

        public const string NoBoardNamedMessage = "No board was named.";

        public const string ChangedMessage =
            "This board changed after it was shown to you - something was published, sent or copied in since. " +
            "Nothing was deleted. Press Delete again to see what it holds now.";

        public const string ReasonRequiredMessage =
            "Say why the board is being deleted - the contributors of its open submissions are told this, and nothing else.";

        public static string NoSuchBoardMessage(string boardId) =>
            $"There is no board [{boardId}] - nothing of it is in the BETA data, the stable data or the database.";

        // In the CRT 4.0 tree (2026-10-25) there is no older main file in the tree and nothing
        // another board uses lies in a board's folder - bytes are stored once for every board
        // naming them, and those another board lists stay (BoardDeletionFiles) - so the Excel
        // tree's "an older main Excel data file lists it" and "another board uses a file in its
        // folder" refusals are gone.
        public static string UnreadableMasterMessage(string treeName, string masterName, string why) =>
            $"[{masterName}] in the {treeName} data could not be read ({why.Trim().TrimEnd('.')}), so whether it lists this " +
            "board is not known. Nothing can be deleted until it can be read.";

        public static string IncompleteTreeMessage(string treeName, IEnumerable<string> problems) =>
            $"The {treeName} data could not be read completely, so whether another board uses a file this board lists " +
            $"is not known. Nothing can be deleted until it can: {string.Join(" ", problems ?? [])}".TrimEnd();

        // The link named from the tree's root, never as the server's own path (code review, 2026-10-13).
        public static string LinkMessage(string treeName, string root, string link) =>
            $"There is a symbolic link at [{TreeWriteAccess.DisplayName(root, link)}] in the {treeName} data, and deleting through it could remove files " +
            "outside the data. Nothing can be deleted until the link is removed on the server.";

        public static string MissingTreeMessage(string treeName) =>
            $"The {treeName} data folder could not be found, so what it holds of this board is not known.";

        public static string UnreadableFolderMessage(string treeName, string why) =>
            $"This board's folder in the {treeName} data could not be read completely ({why.Trim().TrimEnd('.')}). " +
            "Nothing can be deleted until it can.";

        // ---- After a refusal part-way ------------------------------------------------------------

        public static string ListingNotRemovedMessage(string treeName, string failure, bool somethingChanged) =>
            $"The board's row could not be taken out of the {treeName} data's drop-down lists: {failure.Trim().TrimEnd('.')}. " +
            (somethingChanged
                ? "The stable data's row is already gone; press Delete again to finish."
                : "Nothing was changed.");

        public static string FilesLeftMessage(int left) =>
            $"{left} file(s) could not be removed - see the server log. The board is out of the drop-down lists and " +
            "everything else that could go has gone; press Delete again to finish. Its database record and submissions are kept until then.";

        public const string NotRecordedMessage =
            "The board's files and drop-down rows ARE gone, but its database record could not be deleted, so its " +
            "submissions are still there and nobody has been told. Press Delete again to finish - nothing is removed twice.";
    }
}
