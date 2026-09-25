using CRT;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // WHERE A CONTRIBUTOR'S SUBMISSION RECEIPTS LIVE (NewContributeStrategy.md Phase 4, task 6).
    //
    // See SubmissionReceipt's own header for WHY these exist at all. This class is only the
    // storage half: one JSON file beside the user's settings, holding the receipts written by
    // this machine.
    //
    // *** IT LIVES IN AppData, NOT IN THE DATA TREE AND NOT IN A DRAFT. *** Two reasons, and both
    // are load-bearing:
    //   - a receipt carries a CAPABILITY TOKEN, which authorises reading that submission. The
    //     synced Data tree is uploaded from and downloaded to; a draft folder is the very thing
    //     the submit path walks and sends. A token in either would eventually be transmitted
    //     somewhere it has no business being.
    //   - receipts outlive the draft they came from. The whole point is to ask about a submission
    //     after the work is done with, so tying the file's lifetime to a draft that gets discarded
    //     would delete the receipt exactly when it is still wanted.
    //
    // The Load/LoadFrom split is the same test seam UserSettings, DataManager and WorklogManager
    // all use, for the same reason: no test may touch the user's real AppData folder, so LoadFrom
    // takes an explicit path and Load resolves the real one. NEVER call Load() from a test.
    //
    // FAILURES ARE SOFT THROUGHOUT. A receipts file that will not parse costs a contributor the
    // ability to check status in-app; it must never cost them the ability to run the application
    // or to submit again. Every failure logs and degrades to an empty list.
    // ###########################################################################################
    public static class SubmissionReceiptStore
    {
        public const string ReceiptsFileName = "submissions.json";

        private static string _path = string.Empty;
        private static List<SubmissionReceipt> _receipts = new();

        private static readonly JsonSerializerOptions JsonOptions = new()
        {
            WriteIndented = true
        };

        // ###########################################################################################
        // Everything this machine has sent, newest first. A copy, so a caller iterating the list
        // cannot be disturbed by a save happening underneath it.
        // ###########################################################################################
        public static IReadOnlyList<SubmissionReceipt> All =>
            SubmissionReceiptPresenter.InDisplayOrder(_receipts);

        // ###########################################################################################
        // Resolves the real AppData location and loads. Called once at startup.
        // ###########################################################################################
        public static void Load()
        {
            try
            {
                string appData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
                string directory = Path.Combine(appData, AppConfig.AppFolderName);

                Directory.CreateDirectory(directory);

                LoadFrom(Path.Combine(directory, ReceiptsFileName));
            }
            catch (Exception ex)
            {
                Logger.Warning($"Failed to load submission receipts: [{ex.Message}] - starting empty");
                _receipts = new List<SubmissionReceipt>();
            }
        }

        // ###########################################################################################
        // Loads from an explicit path and makes that path the save target - the test seam.
        // ###########################################################################################
        internal static void LoadFrom(string receiptsFilePath)
        {
            _path = receiptsFilePath ?? string.Empty;
            _receipts = new List<SubmissionReceipt>();

            if (string.IsNullOrWhiteSpace(_path) || !File.Exists(_path))
                return;

            try
            {
                string json = File.ReadAllText(_path);

                List<SubmissionReceipt>? loaded =
                    JsonSerializer.Deserialize<List<SubmissionReceipt>>(json, JsonOptions);

                // A receipt with no token can never be used to ask the server anything, so it is
                // dropped rather than shown as a row that silently fails on every refresh.
                _receipts = loaded?
                    .Where(receipt => receipt is not null && !string.IsNullOrWhiteSpace(receipt.UploadToken))
                    .ToList() ?? new List<SubmissionReceipt>();
            }
            catch (Exception ex) when (ex is IOException or JsonException or UnauthorizedAccessException)
            {
                Logger.Warning($"Failed to read submission receipts [{_path}]: [{ex.Message}] - starting empty");
                _receipts = new List<SubmissionReceipt>();
            }
        }

        // ###########################################################################################
        // Records a submission that has just been sent.
        //
        // Called at the moment the server accepts the manifest and hands back a token - NOT at
        // finalise. A submission that was created and then failed mid-upload still exists on the
        // server and still has a state worth asking about, and it is exactly the case a
        // contributor most wants to look up afterwards.
        //
        // Re-recording the same submission id REPLACES the earlier receipt rather than adding a
        // second: the id is the server's own, so two rows for one id could only ever disagree.
        // ###########################################################################################
        public static void Record(SubmissionReceipt receipt)
        {
            ArgumentNullException.ThrowIfNull(receipt);

            if (string.IsNullOrWhiteSpace(receipt.UploadToken))
            {
                // Nothing to record: without the token the row could never be used.
                Logger.Warning("Refusing to record a submission receipt with no token.");
                return;
            }

            _receipts.RemoveAll(existing => existing.SubmissionId == receipt.SubmissionId);
            _receipts.Add(receipt);

            Save();
        }

        // ###########################################################################################
        // Writes back what the server last said about one submission. Does nothing when the id is
        // unknown, so a reply about a receipt deleted meanwhile cannot resurrect it.
        // ###########################################################################################
        public static void UpdateState(
            long submissionId,
            string state,
            string reviewerComment,
            DateTimeOffset checkedUtc,

            // When the REVIEWER decided, as the server reports it. Optional and trailing so every
            // existing caller and test is unaffected; a caller that omits it leaves whatever was
            // already recorded, which is correct - "the server did not tell us" must never erase a
            // date it told us last time.
            DateTimeOffset? decidedUtc = null)
        {
            int index = _receipts.FindIndex(receipt => receipt.SubmissionId == submissionId);
            if (index < 0)
                return;

            SubmissionReceipt existing = _receipts[index];

            _receipts[index] = new SubmissionReceipt
            {
                SubmissionId = existing.SubmissionId,
                UploadToken = existing.UploadToken,
                SystemId = existing.SystemId,
                Summary = existing.Summary,
                SentUtc = existing.SentUtc,
                LastKnownState = state ?? string.Empty,
                LastCheckedUtc = checkedUtc,
                ReviewerComment = reviewerComment ?? string.Empty,

                // *** CARRIED THROUGH UNCHANGED, and that is what makes the badge work. *** A
                // refresh re-reads the same comment from the server every time, so resetting this
                // here would make an already-read comment unread again on every check - a badge
                // that reappears for no reason is one the contributor learns to ignore.
                //
                // It is not cleared when the comment CHANGES either: HasUnreadComment compares the
                // two texts, so a new comment is unread by construction without anything having to
                // remember to reset a flag.
                AcknowledgedComment = existing.AcknowledgedComment ?? string.Empty,

                // Carried for exactly the same reason, and with the same consequence: a refresh
                // re-reads the same STATE every time, so resetting this here would re-badge a
                // decision the contributor has already seen on every launch. A state that has
                // genuinely moved is unread by construction, because HasUnreadDecision compares
                // the two values rather than trusting a flag.
                AcknowledgedState = existing.AcknowledgedState ?? string.Empty,

                // Kept when the caller passes nothing, so a refresh that cannot read a decision
                // date does not wipe one recorded earlier.
                DecidedUtc = decidedUtc ?? existing.DecidedUtc
            };

            Save();
        }

        // ###########################################################################################
        // Records that the contributor has READ the comment currently on this receipt.
        //
        // *** IT STORES THE TEXT IT ACKNOWLEDGED, NOT "true". *** Acknowledging the sentence that
        // is on screen right now means a LATER, different sentence is still unread - which is the
        // case that matters, since a second round of review feedback would otherwise arrive
        // pre-dismissed and never be seen.
        //
        // Deliberately takes no comment argument: it reads the stored one, so a caller cannot
        // acknowledge text that was never actually shown. Doing nothing for an unknown id matches
        // every other method here - the receipt may have been forgotten from another window while
        // this one was open.
        //
        // *** IT ACKNOWLEDGES THE DECIDED STATE TOO SINCE 2026-09-23. *** The badge now counts a
        // decision as news even when the reviewer wrote nothing (see
        // SubmissionReceiptPresenter.HasUnreadDecision), so marking a row read has to clear both
        // or a silently-approved submission would stay badged forever with no way to dismiss it.
        // ###########################################################################################
        public static void AcknowledgeComment(long submissionId)
        {
            int index = _receipts.FindIndex(receipt => receipt.SubmissionId == submissionId);
            if (index < 0)
                return;

            SubmissionReceipt existing = _receipts[index];

            // *** THE EMPTY-COMMENT EARLY RETURN IS GONE, and removing it is the fix. *** It used
            // to bail out here whenever the reviewer had written nothing, which is precisely the
            // publish-with-no-comment case the badge now reports - so the row could be shown,
            // read, and still come back badged on the next launch.
            //
            // Nothing to acknowledge AT ALL is still a needless save, so that case is kept.
            if (!SubmissionReceiptPresenter.HasUnreadNews(existing))
                return;

            _receipts[index] = new SubmissionReceipt
            {
                SubmissionId = existing.SubmissionId,
                UploadToken = existing.UploadToken,
                SystemId = existing.SystemId,
                Summary = existing.Summary,
                SentUtc = existing.SentUtc,
                LastKnownState = existing.LastKnownState,
                LastCheckedUtc = existing.LastCheckedUtc,
                ReviewerComment = existing.ReviewerComment,
                AcknowledgedComment = existing.ReviewerComment,

                // The state as it reads RIGHT NOW, for the same reason the comment is stored as
                // text: a later decision differs from this one and is therefore unread again,
                // with nothing having to remember to reset anything.
                AcknowledgedState = existing.LastKnownState ?? string.Empty,

                // *** CARRIED, like every other field here. *** This method rebuilds the whole
                // record, so a field left out is a field ERASED - marking a comment as read would
                // silently drop the date it was written, and the row would lose its "Replied ..."
                // line the moment the contributor acknowledged it.
                DecidedUtc = existing.DecidedUtc
            };

            Save();
        }

        // ###########################################################################################
        // How many receipts carry a comment the contributor has not yet marked as read.
        //
        // Exposed here rather than made the caller's job so the Drafts tab does not have to know
        // how a receipt is shaped - it asks a question and gets a number for its badge.
        // ###########################################################################################
        public static int UnreadCommentCount() =>
            SubmissionReceiptPresenter.UnreadCommentCount(_receipts);

        // ###########################################################################################
        // Forgets one receipt - the user's own "remove this from my list".
        //
        // THIS DELETES THE ONLY COPY OF THE TOKEN, and with it any further ability to ask about
        // that submission from this machine. It does not withdraw or cancel anything: the
        // submission itself is untouched and will still be reviewed. The confirmation dialog says
        // exactly that, because "remove" next to a queued contribution reads as cancelling it.
        // ###########################################################################################
        public static void Forget(long submissionId)
        {
            if (_receipts.RemoveAll(receipt => receipt.SubmissionId == submissionId) > 0)
                Save();
        }

        private static void Save()
        {
            if (string.IsNullOrWhiteSpace(_path))
                return;

            AtomicJsonFile.Write(_path, _receipts, JsonOptions, "submission receipts");
        }
    }
}
