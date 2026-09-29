using Handlers.DataHandling;

namespace CRT.Server.Handlers.Submissions
{
    // ###########################################################################################
    // The "talk to the submissions database" seam, the same idea as IAccountStore: the flow logic
    // depends on this interface, a MySQL implementation sits behind it as an untested I/O
    // boundary, and an in-memory fake stands in under test.
    //
    // These methods are deliberately dumb - insert a row, mark a file uploaded, read a submission
    // back. Not one of them decides anything. Every rule about what may be submitted, when an
    // upload is complete, and whether a submission may be queued lives in SubmissionFlows and
    // SubmissionValidator, both pure and both fully tested.
    // ###########################################################################################
    public interface ISubmissionStore
    {
        // ###########################################################################################
        // Creates a submission in the 'uploading' state and records every file the manifest lists.
        //
        // Returns the new submission's id, which the client uses for the upload and finalise steps.
        // The id is server-issued rather than client-chosen, so one contributor cannot address
        // another's in-flight submission by guessing.
        // ###########################################################################################
        Task<long> CreateAsync(NewSubmission submission, CancellationToken cancellationToken = default);

        Task<SubmissionRecord?> FindAsync(long submissionId, CancellationToken cancellationToken = default);

        // Every distinct hash this submission needs, with whether its bytes have arrived.
        Task<IReadOnlyList<SubmissionFileRecord>> GetFilesAsync(long submissionId, CancellationToken cancellationToken = default);

        // Marks every file carrying this hash as uploaded. By hash rather than by path, because one
        // blob can legitimately appear at several paths in the same system - a shared image
        // referenced from two boards - and uploading it once must satisfy all of them.
        Task MarkUploadedAsync(long submissionId, string sha256, CancellationToken cancellationToken = default);

        Task SetStateAsync(long submissionId, string state, DateTimeOffset whenUtc, CancellationToken cancellationToken = default);

        // The manifest's rows, stored whole - see 0002_submission_files.sql for why they are JSON
        // rather than eleven tables.
        Task SavePayloadAsync(long submissionId, SubmissionManifest manifest, CancellationToken cancellationToken = default);

        Task<SubmissionManifest?> LoadPayloadAsync(long submissionId, CancellationToken cancellationToken = default);

        // Findings are STORED rather than recomputed: a rejected submission's reasons are shown to
        // the contributor days later, and re-running validation would answer differently once the
        // base revision has moved on.
        Task SaveFindingsAsync(long submissionId, IReadOnlyList<ValidationFinding> findings, CancellationToken cancellationToken = default);

        Task<IReadOnlyList<ValidationFinding>> GetFindingsAsync(long submissionId, CancellationToken cancellationToken = default);

        // Submissions whose upload window has passed and were never finalised. Their blobs must be
        // collected or the disk fills quietly - named as a trap in NewContributeStrategy.md.
        Task<IReadOnlyList<long>> GetExpiredUploadsAsync(DateTimeOffset now, CancellationToken cancellationToken = default);

        // What a contributor sees in "my submissions".
        Task<IReadOnlyList<SubmissionRecord>> GetForAccountAsync(long accountId, int limit, CancellationToken cancellationToken = default);

        // ###########################################################################################
        // The review QUEUE: submissions waiting for somebody to decide about them (Phase 5,
        // task 2).
        //
        // WHAT COUNTS AS WAITING IS A DECISION, NOT A DETAIL. Only 'pending' qualifies. The
        // transport states are not review states - an 'uploading' row is a contribution still
        // arriving and a maintainer acting on it would be deciding about a half-delivered
        // submission - and every other state has already been decided. `SubmissionState.Pending`
        // is the one state that means "arrived intact and is somebody's to decide", which is
        // exactly what a queue is.
        //
        // Newest LAST (ascending by id), deliberately: a queue is worked from the front, and the
        // oldest submission is the one that has been waiting longest. A contributor whose work
        // sits behind a steady trickle of newer ones is how a contribution quietly never gets
        // reviewed.
        //
        // The limit is a guard against a queue that has grown unexpectedly large dragging the
        // whole list into memory, not a paging mechanism - paging is worth building when there is
        // a queue big enough to need it.
        // ###########################################################################################
        Task<IReadOnlyList<SubmissionRecord>> GetQueueAsync(int limit, CancellationToken cancellationToken = default);

        // ###########################################################################################
        // One system's submissions still 'pending', oldest first - the ones a newly queued
        // submission from the same contributor may replace (SubmissionReplacementRules). By system
        // rather than through GetQueueAsync, whose limit could leave an older one out.
        // ###########################################################################################
        Task<IReadOnlyList<SubmissionRecord>> GetPendingForSystemAsync(string systemId, CancellationToken cancellationToken = default);

        // ###########################################################################################
        // One system's submissions for the "Systems" screen (owner request, 2026-09-27), NEWEST
        // first, at most `limit`, each with whether a maintainer decided it (decided_by is set).
        // Leaves out the two that never arrived - 'uploading' and 'abandoned' - which were never
        // anybody's to review.
        // ###########################################################################################
        Task<IReadOnlyList<SystemSubmissionRecord>> GetSubmissionsForSystemAsync(
            string systemId,
            int limit,
            CancellationToken cancellationToken = default);

        // ###########################################################################################
        // Every submission from one contributor - `accountId` when they were signed in, otherwise
        // `contactEmail` (trimmed, any case) among the submissions sent WITHOUT an account - with
        // its state and whether a maintainer decided it. For ContributorHistory (2026-09-26).
        // ###########################################################################################
        Task<IReadOnlyList<ContributorSubmission>> GetContributorSubmissionsAsync(
            long? accountId,
            string? contactEmail,
            CancellationToken cancellationToken = default);

        // ###########################################################################################
        // Withdraws a submission a newer one from the same contributor replaces, with `comment`
        // for the contributor - ONLY while nobody has worked on it: still 'pending' and never
        // amended. Returns whether it was withdrawn.
        //
        // *** BOTH ARE CHECKED INSIDE THE TRANSACTION THAT WITHDRAWS IT. *** A maintainer may be
        // saving table edits to it at this moment: AmendAsync takes the submission row's lock
        // first, so this waits for it and then sees its amendment - and refuses. Checked outside,
        // the edits could land on a submission that is withdrawn a moment later.
        // ###########################################################################################
        Task<bool> WithdrawReplacedAsync(
            long submissionId,
            string comment,
            DateTimeOffset whenUtc,
            CancellationToken cancellationToken = default);

        // ###########################################################################################
        // Records that a system has been published: its new revision and the content hash of the
        // published tree (Phase 5, task 6).
        //
        // THIS IS WHAT `systems.current_revision` AND `content_hash` WERE ADDED FOR in migration
        // 0004, and until now nothing wrote them. current_revision is the base a contributor's
        // next submission is diffed against, so a publish that fails to record it leaves every
        // subsequent submission re-basing against a revision that no longer describes the tree.
        //
        // It is a SEPARATE call from SetStateAsync, deliberately. The submission's state and the
        // system's published revision are two different facts about two different rows, and a
        // publish updates both - but a system can also be published WITHOUT a submission behind it
        // (the project owner correcting their own data), which a combined call could not express.
        // ###########################################################################################
        Task SetSystemPublishedAsync(
            string systemId,
            string revision,
            string contentHash,
            DateTimeOffset publishedUtc,
            CancellationToken cancellationToken = default);

        // ###########################################################################################
        // Records a maintainer's DECISION on a submission (Phase 5, task 5).
        //
        // *** SEPARATE FROM SetStateAsync BECAUSE A DECISION IS NOT JUST A STATE. *** It carries
        // WHO decided and WHY, and `submissions` has had `decided_by` and `decision_comment`
        // columns since 0001_initial.sql with nothing ever writing them. Recording the state alone
        // would leave an audit trail that says a submission was rejected and cannot say by whom or
        // for what reason - and for a contributor the comment is the ENTIRE feedback channel,
        // since contributing needs no account and there is no other way to reach them.
        //
        // The comment is required by the endpoints for a rejection and a change request
        // (ReviewDecisionRules.IsUsableReason); it is optional for an approval, where the summary
        // and the published tree already say what happened.
        // ###########################################################################################
        Task SetDecisionAsync(
            long submissionId,
            string state,
            long decidedByAccountId,
            string? comment,
            DateTimeOffset decidedUtc,
            CancellationToken cancellationToken = default);

        // ###########################################################################################
        // What one address has submitted since a moment - when, and how many bytes each asked to
        // upload - for SubmissionRateLimitPolicy (security review, 2026-09-25).
        // ###########################################################################################
        Task<IReadOnlyList<RecentSubmission>> GetRecentSubmissionsFromAddressAsync(
            string ipAddress,
            DateTimeOffset since,
            CancellationToken cancellationToken = default);

        // ###########################################################################################
        // Whether a system is open to contributions: true or false from its `systems` row, or null
        // when it has no row yet (a new system, or a shipped one nobody has submitted to).
        //
        // `is_accepting` existed from the first migration and nothing read it, so there was no way
        // to close a board that was being flooded. Setting it to 0 by hand now does.
        // ###########################################################################################
        Task<bool?> IsSystemAcceptingAsync(string systemId, CancellationToken cancellationToken = default);

        // ###########################################################################################
        // Every blob hash a submission that still NEEDS its bytes refers to: uploading, pending,
        // approved (not yet published) and merged. What the blob collector must keep.
        //
        // MERGED is kept on purpose. Its blobs are already copied into the tree, but they are also
        // what lets the NEXT submission to that board skip re-uploading every file it did not
        // change - dropping them would turn every later typo fix into a full re-upload.
        // ###########################################################################################
        Task<IReadOnlySet<string>> GetLiveBlobHashesAsync(CancellationToken cancellationToken = default);

        // ###########################################################################################
        // Deletes the file list and stored rows of submissions that ENDED without publishing -
        // rejected, abandoned, withdrawn, returned for changes - and were decided before a moment.
        // Returns how many submissions were cleared.
        //
        // The submissions row and its findings are KEPT: they are the audit trail, and what the
        // contributor's "My submissions" shows. Only the bulk - a copy of the whole board's rows,
        // and one row per file - goes, which is what an anonymous sender could otherwise pile up.
        // ###########################################################################################
        Task<int> DeleteRetiredPayloadsAsync(DateTimeOffset decidedBefore, CancellationToken cancellationToken = default);

        // ###########################################################################################
        // Every `systems` row, for the administrator's maintainer overview (Phase 6 roles).
        //
        // Only systems that have received a submission or been published have a row; a shipped
        // board nobody has touched has none. MaintainerAssignmentFlows unions this with the boards
        // found in the data tree, so a maintainer can be assigned to a board BEFORE its first
        // submission arrives - which is the ordinary order of events.
        // ###########################################################################################
        Task<IReadOnlyList<SystemRecord>> ListSystemsAsync(CancellationToken cancellationToken = default);

        // ###########################################################################################
        // Makes sure a `systems` row exists, leaving an existing one completely alone - the same
        // INSERT IGNORE CreateAsync performs for a submission. Needed before a maintainer can be
        // assigned to a shipped board that has never been submitted to: the pool table's foreign
        // key requires the row.
        // ###########################################################################################
        Task EnsureSystemAsync(
            string systemId,
            string manufacturer,
            string hardware,
            string board,
            string origin,
            DateTimeOffset createdUtc,
            CancellationToken cancellationToken = default);

        // One `systems` row, or null.
        Task<SystemRecord?> FindSystemAsync(string systemId, CancellationToken cancellationToken = default);

        // ###########################################################################################
        // Where a NEW system goes in the drop-down lists (2026-09-27, migration 0011): what a
        // maintainer placed it as in the Systems screen, or null while nobody has. SetPlacementAsync
        // answers false when the system has no row at all.
        // ###########################################################################################
        Task<SystemPlacement?> GetPlacementAsync(string systemId, CancellationToken cancellationToken = default);

        Task<bool> SetPlacementAsync(
            string systemId,
            SystemPlacement placement,
            long setByAccountId,
            DateTimeOffset setUtc,
            CancellationToken cancellationToken = default);

        // ###########################################################################################
        // Records that a system's BETA state has been copied to Production (2026-09-25): the BETA
        // revision and content hash it had, and when. A separate call from SetSystemPublishedAsync
        // because it is a separate fact about a separate tree - and the comparison between the two
        // is exactly what "waiting for production" means.
        // ###########################################################################################
        Task SetSystemInProductionAsync(
            string systemId,
            string? revision,
            string? contentHash,
            DateTimeOffset publishedUtc,
            CancellationToken cancellationToken = default);

        // ###########################################################################################
        // Records a BETA rollback (owner decision, 2026-09-27), in ONE TRANSACTION:
        //
        //   - each returning submission goes back to `pending`, if it is still `merged`, with the
        //     maintainer who pushed it back, when, and why (the contributor's only feedback) - or to
        //     `rejected`, when `reject` says the rollback was Beta > Prod's "Reject" (2026-09-28);
        //   - each one's APPROVALS are cleared. Left in place, ApprovalRules re-read them on the
        //     next review: a shared-file submission approved by both roles would be republished by
        //     ONE approval, and an administrator who approved an ordinary one was answered "you have
        //     already approved" and could never approve it again (code review, 2026-09-27);
        //   - the system's BETA revision and content hash follow the tree - production's for a
        //     restore, none for a system removed from BETA. Every other writer of those columns is a
        //     publish, moving them FORWARDS through SetSystemPublishedAsync.
        //
        // *** ONE TRANSACTION, because the tree has already been rewritten when this runs. *** As
        // separate calls a failure halfway left some submissions pending and some merged, the ones
        // flipped never mailed, and content_hash naming data BETA no longer holds (code review,
        // 2026-09-27). Now it is all or nothing, and "nothing" repairs itself: the system still
        // reads as ahead of production, so pushing back again finds a tree already level, changes
        // no file, and records it.
        // ###########################################################################################
        Task RecordRollbackAsync(
            string systemId,
            IReadOnlyList<long> returningSubmissionIds,
            long decidedByAccountId,
            string comment,
            string? betaRevision,
            string? betaContentHash,
            DateTimeOffset decidedUtc,
            bool reject = false,
            CancellationToken cancellationToken = default);

        // ###########################################################################################
        // The MERGED submissions to a system decided in (after, upTo] - the ones a production
        // promotion has just carried out to everyone, whose contributors are told so. `after` null
        // means "since the beginning": the first promotion of a system carries everything merged.
        // ###########################################################################################
        Task<IReadOnlyList<SubmissionRecord>> GetMergedSubmissionsAsync(
            string systemId,
            DateTimeOffset? decidedAfter,
            DateTimeOffset decidedUpTo,
            CancellationToken cancellationToken = default);

        // ###########################################################################################
        // The contributor discarded their own draft in CRT after sending this submission (owner
        // request, 2026-09-28; migration 0014) - see CRT.Data's DraftDiscardContract and
        // DraftDiscardFlow. Records it once: true when this call recorded it, false when it was
        // already recorded (the FIRST time is kept).
        // ###########################################################################################
        Task<bool> RecordDraftDiscardedAsync(
            long submissionId,
            DateTimeOffset discardedUtc,
            CancellationToken cancellationToken = default);

        // When each of `submissionIds` had its draft discarded - only those that did appear. One
        // query for a whole queue or list, so showing the mark costs one read, not one per row.
        Task<IReadOnlyDictionary<long, DateTimeOffset>> GetDraftDiscardsAsync(
            IReadOnlyCollection<long> submissionIds,
            CancellationToken cancellationToken = default);

        // ###########################################################################################
        // THE "BETA > PROD" LIST'S TWO FACTS, FOR EVERY WAITING SYSTEM AT ONCE (code review,
        // 2026-09-29). The list is read every minute by every open Maintainer tab, and asked three
        // queries PER waiting system (approvals, merged submissions, their discards) only to set
        // two booleans. These answer for the whole list in one query each.
        //
        // The production approvals given for each system's BETA state, keyed by system id (a system
        // with none is absent).
        // ###########################################################################################
        Task<IReadOnlyDictionary<string, IReadOnlyList<GivenApproval>>> GetProductionApprovalsForAsync(
            IReadOnlyCollection<(string SystemId, string BetaContentHash)> states,
            CancellationToken cancellationToken = default);

        // The systems among `windows` with a submission a promotion would carry - merged in
        // (DecidedAfter, decidedUpTo], GetMergedSubmissionsAsync's bounds - whose contributor has
        // discarded their own draft (migration 0014).
        Task<IReadOnlySet<string>> GetSystemsCarryingDiscardedDraftsAsync(
            IReadOnlyCollection<(string SystemId, DateTimeOffset? DecidedAfter)> windows,
            DateTimeOffset decidedUpTo,
            CancellationToken cancellationToken = default);

        // When a BETA rollback last RETURNED each of `submissionIds` to the queue (migration 0015,
        // written by RecordRollbackAsync) - only those it did. What makes a submission read as
        // "returned" (ProductionPromotionRules.ContributorFacingState); one query for a list.
        Task<IReadOnlyDictionary<long, DateTimeOffset>> GetBetaReturnsAsync(
            IReadOnlyCollection<long> submissionIds,
            CancellationToken cancellationToken = default);

        // ###########################################################################################
        // Approvals already given (migration 0008) - the first of the two a shared-file change
        // needs, remembered until the second arrives. See ApprovalRules. Adding a role that is
        // already there changes nothing: the first approval in a role is the one kept.
        // ###########################################################################################
        Task<IReadOnlyList<GivenApproval>> GetApprovalsAsync(long submissionId, CancellationToken cancellationToken = default);

        Task AddApprovalAsync(
            long submissionId,
            ApproverRole role,
            long accountId,
            string accountLabel,
            DateTimeOffset approvedUtc,
            CancellationToken cancellationToken = default);

        // The same for publishing a system to production, per BETA content hash: an approval of
        // one BETA state never carries over to the next.
        Task<IReadOnlyList<GivenApproval>> GetProductionApprovalsAsync(
            string systemId,
            string betaContentHash,
            CancellationToken cancellationToken = default);

        Task AddProductionApprovalAsync(
            string systemId,
            string betaContentHash,
            ApproverRole role,
            long accountId,
            string accountLabel,
            DateTimeOffset approvedUtc,
            CancellationToken cancellationToken = default);

        // ###########################################################################################
        // A MAINTAINER'S AMENDMENT (migration 0009), in ONE transaction: the submission's current rows
        // and files become `amended`'s; what they replace is kept in submission_amendments (the
        // first row holds the contributor's original); the approvals already given are cleared -
        // they were given to other content - and an 'approved' submission goes back to 'pending';
        // and whether it changes shared files is stored afresh. Every file in `amended` must
        // already be in the blob store: the caller imports first. The new version is 1 for the
        // first amendment.
        //
        // *** THE VERSION AND THE STATE ARE CHECKED INSIDE THE TRANSACTION (code review,
        // 2026-09-25). *** Nothing is changed unless the submission is still amendable
        // (SubmissionState.CanBeAmended) and its latest amendment is still `expectedVersion` - the
        // one the maintainer opened - both read under the row locks the write holds. The caller's
        // own checks run earlier and outside it, so two amendments at once both passed them and
        // the second silently overwrote the first.
        // ###########################################################################################
        Task<AmendStoreResult> AmendAsync(
            long submissionId,
            int expectedVersion,
            SubmissionManifest amended,
            bool touchesSharedFiles,
            long? accountId,
            string accountLabel,
            DateTimeOffset amendedUtc,
            CancellationToken cancellationToken = default);

        // The latest amendment, or null for a submission nobody has amended.
        Task<SubmissionAmendment?> GetLatestAmendmentAsync(long submissionId, CancellationToken cancellationToken = default);

        // Whether the submission REPLACES a shared file, as the tree stands now - raised when a
        // published copy has moved since it arrived, lowered when it no longer replaces one (or was
        // flagged under the older add-or-change rule). See ApprovePublishFlow.TouchesSharedFilesNow.
        Task SetTouchesSharedFilesAsync(long submissionId, bool touchesSharedFiles, CancellationToken cancellationToken = default);
    }

    // ###########################################################################################
    // A `systems` row.
    //
    // CurrentRevision and ContentHash describe what is in BETA - every publish writes them. The
    // three Production* fields (migration 0007) describe what was last copied to Production, so
    // "BETA is ahead" is a comparison of two columns rather than of two trees on disk. See
    // ProductionPromotionRules.
    // ###########################################################################################
    public sealed record SystemRecord(
        string SystemId,
        string Manufacturer,
        string Hardware,
        string Board,
        string? CurrentRevision,
        bool IsAccepting,
        string? ContentHash = null,
        string? ProductionRevision = null,
        string? ProductionContentHash = null,
        DateTimeOffset? ProductionPublishedUtc = null);

    // ###########################################################################################
    // The states a submission passes through. Strings rather than an enum because they are stored
    // as text and read in a log line; the constants stop them being mistyped.
    // ###########################################################################################
    // What ISubmissionStore.AmendAsync did. Version is the new amendment's when Amended, and the
    // latest one found when VersionChanged.
    public sealed record AmendStoreResult(AmendStoreOutcome Outcome, int Version)
    {
        public bool IsAmended => this.Outcome == AmendStoreOutcome.Amended;
    }

    public enum AmendStoreOutcome
    {
        Amended,

        // Another amendment was stored after the maintainer opened this one.
        VersionChanged,

        // Decided (or never finished uploading) - SubmissionState.CanBeAmended is false.
        NotAmendable
    }

    public static class SubmissionState
    {
        // Created, files being uploaded. Not yet visible to a maintainer.
        public const string Uploading = "uploading";

        // Complete, validated, queued for review.
        public const string Pending = "pending";

        // Failed automated validation and never queued.
        public const string Rejected = "rejected";

        // The upload window passed without being finalised.
        public const string Abandoned = "abandoned";

        // ---- Review states (Phase 5) -------------------------------------------------------
        //
        // These were in 0001_initial.sql's CHECK constraint from the start but had no constant
        // here, because Phase 4 only ever wrote the transport states above. Adding them as
        // constants rather than string literals at the call site matters more than usual: the
        // database CHECK rejects an unknown value, so a typo is a publish that throws at the very
        // last step, AFTER the tree has already been written.

        // Published into the data tree. The terminal success state.
        public const string Merged = "merged";

        // Returned to the contributor with a comment, as an editable draft.
        public const string ChangesRequested = "changes_requested";

        // Accepted by a maintainer but not yet published. Since 2026-09-25 this is what the FIRST of
        // the two approvals a shared-file change needs leaves behind - it stays in the review
        // queue until the second approval publishes it. See ApprovalRules.
        public const string Approved = "approved";

        // Taken back by the contributor.
        public const string Withdrawn = "withdrawn";

        // Still undecided, so a maintainer may change it: waiting for review, or for the second of
        // two approvals. The one rule AmendSubmissionFlow and both stores check it by.
        public static bool CanBeAmended(string? state) => state is Pending or Approved;

        // Every state the database's CHECK constraint allows. SubmissionCollectionStatesTests reads
        // the constraint out of the migrations and holds this list to it.
        public static readonly IReadOnlyList<string> All =
        [
            Uploading, Abandoned, Pending, ChangesRequested, Approved, Rejected, Withdrawn, Merged
        ];
    }

    // ###########################################################################################
    // Which submissions still NEED their stored bytes, and which have ended without publishing
    // (security review, 2026-09-25) - the two halves the collectors work from.
    //
    // *** COMPLEMENTS, AND IT MATTERS WHICH SIDE A NEW STATE LANDS ON. *** A state in neither list
    // keeps its blobs for ever; a state in both has them swept while a maintainer may still need
    // them. SubmissionCollectionStatesTests fails if the two stop covering SubmissionState.All
    // exactly once between them - which is the prompt, when a state is added, to decide here.
    //
    // FOUR EACH, because MySqlSubmissionStore binds them as four parameters.
    // ###########################################################################################
    public static class SubmissionCollectionStates
    {
        public static readonly IReadOnlyList<string> Live =
        [
            SubmissionState.Uploading, SubmissionState.Pending, SubmissionState.Approved, SubmissionState.Merged
        ];

        public static readonly IReadOnlyList<string> Retired =
        [
            SubmissionState.Rejected, SubmissionState.Abandoned, SubmissionState.Withdrawn, SubmissionState.ChangesRequested
        ];
    }

    // ###########################################################################################
    // Manufacturer/Hardware/Board travel WITH the submission, not just inside SystemId.
    //
    // They are needed to create the `systems` row a first submission for a new system implies -
    // that table stores the three parts separately "so the maintainer app can list by manufacturer
    // without parsing" (0001_initial.sql). Splitting SystemId back apart in the store would be
    // that parsing, in the one place the schema says to avoid it.
    // ###########################################################################################
    public sealed record NewSubmission(
        string SystemId,
        string Manufacturer,
        string Hardware,
        string Board,
        long? AccountId,
        string? ContactEmail,
        string? CreatedIp,
        string UploadTokenHash,
        string BaseRevision,
        string Summary,
        int FormatVersion,
        IReadOnlyList<SubmissionFile> Files,
        DateTimeOffset CreatedUtc,
        DateTimeOffset ExpiresUtc,

        // How many bytes the server asked this submission to upload - the files it did not
        // already hold - so the per-address budget can count what was really requested.
        long BytesToUpload = 0,

        // Whether it adds or changes a file under "Shared files" / "Generic shared files" -
        // decided at creation by SubmissionSharedFiles, and what routes it to the administrator
        // rather than to the system's maintainers. See ReviewAuthority.
        bool TouchesSharedFiles = false);

    public sealed record SubmissionRecord(
        long Id,
        string SystemId,
        long? AccountId,
        string? ContactEmail,
        string? UploadTokenHash,
        string BaseRevision,
        string State,
        string? Summary,
        int FormatVersion,
        DateTimeOffset CreatedUtc,
        DateTimeOffset? ExpiresUtc,
        DateTimeOffset? DecidedUtc,

        // ###########################################################################################
        // What the maintainer said, for a submission that was rejected or returned for changes.
        //
        // *** THIS IS THE CONTRIBUTOR'S ONLY FEEDBACK. *** Contributing needs no account, so there
        // is no inbox, no thread and no history - the contact email and this sentence are the whole
        // channel. It reaches them through the contributor's own status endpoint, whose
        // `maintainerComment` field was reserved from the start for exactly this.
        //
        // TRAILING and OPTIONAL so every existing construction keeps working; a record built
        // without it simply has none, which is correct for anything not yet decided.
        // ###########################################################################################
        string? DecisionComment = null,

        // Shared files belong to no system, so a submission changing one is the administrator's
        // to decide whichever board it names - see ReviewAuthority. Stored on the row so the
        // queue filters without loading the payload.
        bool TouchesSharedFiles = false);

    // One of a contributor's submissions, as ContributorHistory counts it. DecidedByMaintainer
    // tells a maintainer's rejection from the automatic checks' - decided_by is set only by a person.
    public sealed record ContributorSubmission(long Id, string State, bool DecidedByMaintainer);

    // One submission to a system, with whether a maintainer decided it - GetSubmissionsForSystemAsync.
    // DecidedByAccountId: the maintainer who decided it (2026-09-27, for the system's history) -
    // null when none did, or from a store that does not say.
    public sealed record SystemSubmissionRecord(SubmissionRecord Submission, bool DecidedByMaintainer, long? DecidedByAccountId = null);

    // One maintainer's amendment: its version (1, 2, ...), who made it, and when.
    public sealed record SubmissionAmendment(int Version, string By, DateTimeOffset AtUtc);

    public sealed record SubmissionFileRecord(
        long Id,
        string Path,
        string Sha256,
        long SizeBytes,
        bool IsUploaded);
}
