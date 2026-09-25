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
        // arriving and a reviewer acting on it would be deciding about a half-delivered
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
        // (the maintainer correcting their own data), which a combined call could not express.
        // ###########################################################################################
        Task SetSystemPublishedAsync(
            string systemId,
            string revision,
            string contentHash,
            DateTimeOffset publishedUtc,
            CancellationToken cancellationToken = default);

        // ###########################################################################################
        // Records a reviewer's DECISION on a submission (Phase 5, task 5).
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
    }

    // ###########################################################################################
    // The states a submission passes through. Strings rather than an enum because they are stored
    // as text and read in a log line; the constants stop them being mistyped.
    // ###########################################################################################
    public static class SubmissionState
    {
        // Created, files being uploaded. Not yet visible to a reviewer.
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

        // Accepted by a reviewer but not yet published.
        public const string Approved = "approved";

        // Taken back by the contributor.
        public const string Withdrawn = "withdrawn";
    }

    // ###########################################################################################
    // Manufacturer/Hardware/Board travel WITH the submission, not just inside SystemId.
    //
    // They are needed to create the `systems` row a first submission for a new system implies -
    // that table stores the three parts separately "so the review app can list by manufacturer
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
        DateTimeOffset ExpiresUtc);

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
        // What the reviewer said, for a submission that was rejected or returned for changes.
        //
        // *** THIS IS THE CONTRIBUTOR'S ONLY FEEDBACK. *** Contributing needs no account, so there
        // is no inbox, no thread and no history - the contact email and this sentence are the whole
        // channel. It reaches them through the contributor's own status endpoint, whose
        // `reviewerComment` field was reserved from the start for exactly this.
        //
        // TRAILING and OPTIONAL so every existing construction keeps working; a record built
        // without it simply has none, which is correct for anything not yet decided.
        // ###########################################################################################
        string? DecisionComment = null);

    public sealed record SubmissionFileRecord(
        long Id,
        string Path,
        string Sha256,
        long SizeBytes,
        bool IsUploaded);
}
