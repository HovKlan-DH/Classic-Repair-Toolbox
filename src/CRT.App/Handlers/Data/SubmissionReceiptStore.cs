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

        // ###########################################################################################
        // *** EVERY READ AND WRITE HOLDS THIS (code review, 2026-09-29). *** The Submit dialog records
        // and finalises its receipt from the THREAD POOL - its upload runs under Task.Run - while the
        // launch status check, the Drafts tab and the discard reporter use the same list on the UI
        // thread. Unguarded, two writers at once lost receipts, a reader threw "Collection was
        // modified", and two saves raced on the same temporary file. Reads hand back copies, so a
        // caller iterating one is never disturbed by a write. Save runs inside the lock, so the file
        // is always written from a list no other thread is changing.
        // ###########################################################################################
        private static readonly object Gate = new();

        private static readonly JsonSerializerOptions JsonOptions = new()
        {
            WriteIndented = true
        };

        // ###########################################################################################
        // Everything this machine has sent, newest first. A copy, so a caller iterating the list
        // cannot be disturbed by a save happening underneath it.
        // ###########################################################################################
        public static IReadOnlyList<SubmissionReceipt> All
        {
            get
            {
                lock (Gate)
                    return SubmissionReceiptPresenter.InDisplayOrder(_receipts.ToList());
            }
        }

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

                lock (Gate)
                    _receipts = new List<SubmissionReceipt>();
            }
        }

        // ###########################################################################################
        // Loads from an explicit path and makes that path the save target - the test seam.
        // ###########################################################################################
        internal static void LoadFrom(string receiptsFilePath)
        {
            lock (Gate)
                LoadFromLocked(receiptsFilePath);
        }

        private static void LoadFromLocked(string receiptsFilePath)
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

            lock (Gate)
            {
                _receipts.RemoveAll(existing => existing.SubmissionId == receipt.SubmissionId);
                _receipts.Add(receipt);

                Save();
            }
        }

        // ###########################################################################################
        // THE STATE THE SERVER CONFIRMED WHEN THE SUBMISSION WAS FINALISED (owner report,
        // 2026-09-27). The receipt is written before any byte is uploaded (SubmitDraftWindow says
        // why), so until now it knew no state at all until the next launch's status check - and the
        // Drafts tab's badge read "Not checked yet" straight after a successful send. The finalise
        // answer already carries the state ("pending", or "rejected" when the automatic checks
        // refused it), so it is recorded here.
        //
        // A refusal is recorded as SEEN: the send window has just shown its reasons, and badging it
        // as unread news would repeat what the contributor is looking at. An answer with no state
        // (an older server) changes nothing, rather than stamping "checked" onto an unknown state.
        // ###########################################################################################
        public static void RecordFinalised(SubmissionResult? result, DateTimeOffset checkedUtc)
        {
            if (result is null || string.IsNullOrWhiteSpace(result.State))
                return;

            // One lock across both, so no reader sees the new state not yet acknowledged. (The
            // lock is re-entrant for this thread, so the two calls take it again freely.)
            lock (Gate)
            {
                UpdateState(result.SubmissionId, result.State.Trim(), string.Empty, checkedUtc);

                if (!result.IsAccepted)
                    AcknowledgeComment(result.SubmissionId);
            }
        }

        // ###########################################################################################
        // Writes back what the server last said about one submission. Does nothing when the id is
        // unknown, so a reply about a receipt deleted meanwhile cannot resurrect it.
        // ###########################################################################################
        public static void UpdateState(
            long submissionId,
            string state,
            string maintainerComment,
            DateTimeOffset checkedUtc,

            // When the MAINTAINER decided, as the server reports it. Optional and trailing so every
            // existing caller and test is unaffected; a caller that omits it leaves whatever was
            // already recorded, which is correct - "the server did not tell us" must never erase a
            // date it told us last time.
            DateTimeOffset? decidedUtc = null,

            // Whether a maintainer changed it (2026-09-25). Optional and trailing for the same reason
            // as decidedUtc, and kept when omitted for the same reason too.
            bool? amendedByMaintainer = null)
        {
            lock (Gate)
            {
                int index = _receipts.FindIndex(receipt => receipt.SubmissionId == submissionId);
                if (index < 0)
                    return;

                SubmissionReceipt existing = _receipts[index];

                // *** EVERYTHING NOT NAMED HERE IS CARRIED, by `with` - and that is what makes the
                // badges work. *** AcknowledgedComment and AcknowledgedState stay as they were: a
                // refresh re-reads the same comment and state from the server every time, so resetting
                // them here would make an already-read decision unread again on every check. A comment
                // or state that genuinely CHANGED is unread by construction, because HasUnreadComment /
                // HasUnreadDecision compare texts rather than trusting a flag. SourceNoticeDismissed and
                // the discard fields are carried too: a refresh must not bring back a closed notice or
                // forget a discard.
                _receipts[index] = existing with
                {
                    LastKnownState = state ?? string.Empty,
                    LastCheckedUtc = checkedUtc,
                    MaintainerComment = maintainerComment ?? string.Empty,

                    // Kept when the caller passes nothing, so a refresh that cannot read a decision
                    // date does not wipe one recorded earlier.
                    DecidedUtc = decidedUtc ?? existing.DecidedUtc,
                    AmendedByMaintainer = amendedByMaintainer ?? existing.AmendedByMaintainer,

                    // The server answered, so it knows the submission again.
                    NotFoundUtc = null
                };

                Save();
            }
        }

        // ###########################################################################################
        // How stale a receipt's LastCheckedUtc may get before an unchanged answer writes it (code
        // review, 2026-10-04). The minute check asked about every open receipt and saved the whole
        // file for each one, only to move this date - which nothing shows, and which gates nothing
        // finer than a once-a-day or once-a-week re-check (SubmissionReceiptPresenter.IsStillOpen).
        // ###########################################################################################
        public static readonly TimeSpan CheckedSaveInterval = TimeSpan.FromHours(12);

        // ###########################################################################################
        // The server answered with nothing new: the date it was checked is written only when the
        // stored one is missing or older than CheckedSaveInterval, so an unchanged answer costs no
        // write at all most of the time. An answer that clears a "not found" is news, and is
        // written by UpdateState instead.
        // ###########################################################################################
        public static void NoteChecked(long submissionId, DateTimeOffset checkedUtc)
        {
            lock (Gate)
            {
                int index = _receipts.FindIndex(receipt => receipt.SubmissionId == submissionId);
                if (index < 0)
                    return;

                SubmissionReceipt existing = _receipts[index];

                if (existing.LastCheckedUtc is DateTimeOffset last && checkedUtc - last < CheckedSaveInterval)
                    return;

                _receipts[index] = existing with { LastCheckedUtc = checkedUtc };
                Save();
            }
        }

        // ###########################################################################################
        // The server answered that it does not know this submission (HTTP 404) - deleted with its
        // system, for one. Its state is kept as it was; it is then asked about only once per
        // SubmissionReceiptPresenter.NotFoundRecheckInterval (code review, 2026-10-04: it was asked
        // every minute for ever).
        // ###########################################################################################
        public static void NoteNotFound(long submissionId, DateTimeOffset checkedUtc)
        {
            lock (Gate)
            {
                int index = _receipts.FindIndex(receipt => receipt.SubmissionId == submissionId);
                if (index < 0)
                    return;

                SubmissionReceipt existing = _receipts[index];

                _receipts[index] = existing with
                {
                    LastCheckedUtc = checkedUtc,
                    NotFoundUtc = existing.NotFoundUtc ?? checkedUtc
                };

                Save();
            }
        }

        // ###########################################################################################
        // Records that the contributor closed the "now in the online source - switch back from
        // BETA" notice for these submissions (SubmissionReceiptPresenter.NeedingSourceSwitchNotice),
        // so it is not shown for them again. One save for them all.
        // ###########################################################################################
        public static void DismissSourceNotice(IEnumerable<long> submissionIds) =>
            SubmissionReceiptStore.Change(
                submissionIds,
                receipt => !receipt.SourceNoticeDismissed,
                receipt => receipt with { SourceNoticeDismissed = true });

        // The same for the "now in the BETA source - tick BETA to try it" notice (2026-10-03,
        // SubmissionReceiptPresenter.NeedingBetaTryNotice).
        public static void DismissBetaNotice(IEnumerable<long> submissionIds) =>
            SubmissionReceiptStore.Change(
                submissionIds,
                receipt => !receipt.BetaNoticeDismissed,
                receipt => receipt with { BetaNoticeDismissed = true });

        // Applies `change` to each of these receipts that `needs` it, with one save for them all.
        private static void Change(
            IEnumerable<long> submissionIds,
            Func<SubmissionReceipt, bool> needs,
            Func<SubmissionReceipt, SubmissionReceipt> change)
        {
            var ids = new HashSet<long>(submissionIds ?? []);

            lock (Gate)
            {
                bool changed = false;

                for (int index = 0; index < _receipts.Count; index++)
                {
                    SubmissionReceipt existing = _receipts[index];
                    if (!ids.Contains(existing.SubmissionId) || !needs(existing))
                        continue;

                    _receipts[index] = change(existing);
                    changed = true;
                }

                if (changed)
                    Save();
            }
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
        // decision as news even when the maintainer wrote nothing (see
        // SubmissionReceiptPresenter.HasUnreadDecision), so marking a row read has to clear both
        // or a silently-approved submission would stay badged forever with no way to dismiss it.
        // ###########################################################################################
        public static void AcknowledgeComment(long submissionId)
        {
            lock (Gate)
            {
                int index = _receipts.FindIndex(receipt => receipt.SubmissionId == submissionId);
                if (index < 0)
                    return;

                SubmissionReceipt existing = _receipts[index];

                // *** THE EMPTY-COMMENT EARLY RETURN IS GONE, and removing it is the fix. *** It used
                // to bail out here whenever the maintainer had written nothing, which is precisely
                // the publish-with-no-comment case the badge now reports - so the row could be shown,
                // read, and still come back badged on the next launch.
                //
                // Nothing to acknowledge AT ALL is still a needless save, so that case is kept.
                if (!SubmissionReceiptPresenter.HasUnreadNews(existing))
                    return;

                // The comment and the state as they read RIGHT NOW, stored as text: a later comment or
                // decision differs from these and is therefore unread again, with nothing having to
                // remember to reset anything. Every other field is carried by `with`.
                _receipts[index] = existing with
                {
                    AcknowledgedComment = existing.MaintainerComment,
                    AcknowledgedState = existing.LastKnownState ?? string.Empty
                };

                Save();
            }
        }

        // ###########################################################################################
        // THE CONTRIBUTOR DISCARDED THEIR DRAFT (owner request, 2026-09-28) - see CRT.Data's
        // DraftDiscardContract. Marks these receipts discarded now and not yet reported; the reporter
        // (DraftDiscardReporter) tells the server and marks each one reported once an answer
        // finishes it. One save for them all. An id already marked keeps its first time.
        // ###########################################################################################
        public static void MarkDraftDiscarded(IEnumerable<long> submissionIds, DateTimeOffset nowUtc)
        {
            HashSet<long> ids = [.. submissionIds];

            lock (Gate)
            {
                bool changed = false;

                for (int index = 0; index < _receipts.Count; index++)
                {
                    SubmissionReceipt existing = _receipts[index];

                    if (!ids.Contains(existing.SubmissionId) || existing.DraftDiscardedUtc is not null)
                        continue;

                    _receipts[index] = existing with { DraftDiscardedUtc = nowUtc, DraftDiscardReported = false };
                    changed = true;
                }

                if (changed)
                    Save();
            }
        }

        // The discards the server has not been told about yet.
        public static IReadOnlyList<SubmissionReceipt> PendingDraftDiscardNotices
        {
            get
            {
                lock (Gate)
                    return _receipts.Where(receipt => receipt.DraftDiscardedUtc is not null && !receipt.DraftDiscardReported).ToList();
            }
        }

        // The server has the notice (or can never take it) - it is not sent again.
        public static void MarkDraftDiscardReported(long submissionId)
        {
            lock (Gate)
            {
                int index = _receipts.FindIndex(receipt => receipt.SubmissionId == submissionId);

                if (index < 0 || _receipts[index].DraftDiscardReported)
                    return;

                SubmissionReceipt existing = _receipts[index];
                _receipts[index] = existing with
                {
                    DraftDiscardedUtc = existing.DraftDiscardedUtc ?? DateTimeOffset.UtcNow,
                    DraftDiscardReported = true
                };

                Save();
            }
        }

        // ###########################################################################################
        // How many receipts carry a comment the contributor has not yet marked as read.
        //
        // Exposed here rather than made the caller's job so the Drafts tab does not have to know
        // how a receipt is shaped - it asks a question and gets a number for its badge.
        // ###########################################################################################
        public static int UnreadCommentCount()
        {
            lock (Gate)
                return SubmissionReceiptPresenter.UnreadCommentCount(_receipts);
        }

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
            lock (Gate)
            {
                if (_receipts.RemoveAll(receipt => receipt.SubmissionId == submissionId) > 0)
                    Save();
            }
        }

        // Always called with Gate held - see Gate.
        private static void Save()
        {
            if (string.IsNullOrWhiteSpace(_path))
                return;

            AtomicJsonFile.Write(_path, _receipts, JsonOptions, "submission receipts");
        }
    }
}
