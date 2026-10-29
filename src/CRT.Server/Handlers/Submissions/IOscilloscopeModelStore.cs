namespace CRT.Server.Handlers.Submissions
{
    // ###########################################################################################
    // Where the oscilloscope models sent in from CRT wait for the administrator (migration 0024,
    // OscilloscopeModelFlows). MySqlOscilloscopeModelStore is the database; the tests use a fake.
    // ###########################################################################################
    public interface IOscilloscopeModelStore
    {
        // Stores a model sent in, as waiting - and marks every model of the same name the same
        // contributor sent that is still pending as replaced (OscilloscopeModelFlows' rule; one in
        // BETA never is). The same contributor is the same account - or, without one, the same typed
        // address AND the same internet connection (SenderAddress): an address alone is no proof. Its id.
        Task<long> AddAsync(NewOscilloscopeModelSubmission submission, CancellationToken cancellationToken = default);

        // How many models `senderAddress` (an internet connection, OscilloscopeModelFlows.ConnectionOf)
        // sent since `sinceUtc` - the daily limit's count.
        Task<int> CountFromAddressSinceAsync(string senderAddress, DateTimeOffset sinceUtc, CancellationToken cancellationToken = default);

        // Every model waiting - pending, or in BETA only - oldest first.
        Task<IReadOnlyList<OscilloscopeModelSubmissionRecord>> ListPendingAsync(CancellationToken cancellationToken = default);

        // How many are waiting - the Maintainer tab's badge (the review queue answer), every minute.
        Task<int> CountWaitingAsync(CancellationToken cancellationToken = default);

        Task<OscilloscopeModelSubmissionRecord?> FindAsync(long id, CancellationToken cancellationToken = default);

        // A publish that wrote BETA's main file and not the stable one's: the model is in BETA, and
        // waits there for Publish again - whether it was pending or replaced meanwhile (what was
        // written IS in BETA). False when it was neither.
        Task<bool> MarkInBetaAsync(long id, CancellationToken cancellationToken = default);

        // Another model of this name was written into BETA: every model of the name in BETA but
        // `exceptId` is no longer what BETA holds, and is marked replaced - finished, its older
        // commands would be written back over the newer ones. The records replaced, as they were.
        //
        // NAMES ARE MATCHED BY OscilloscopeModelRules.SameName, never by the database's collation
        // (code review, 2026-10-29: utf8mb4_unicode_ci ignores accents and takes "ss" for "ß", so it
        // took two models the main file keeps apart for one) - AddAsync's replacement too.
        Task<IReadOnlyList<OscilloscopeModelSubmissionRecord>> ReplaceInBetaOfNameAsync(string vendor, string seriesOrModel, long exceptId, CancellationToken cancellationToken = default);

        // The decision - only on a model still waiting (pending or in BETA); false when it was not
        // (decided meanwhile). `evenIfReplaced`: a publish under way records its model published even
        // when a newer copy marked it replaced meanwhile - what was written IS published.
        Task<bool> DecideAsync(long id, string state, long decidedByAccountId, string? note, DateTimeOffset whenUtc, bool evenIfReplaced = false, CancellationToken cancellationToken = default);
    }

    public static class OscilloscopeModelStates
    {
        public const string Pending = "pending";

        // Written into BETA's main file by a publish the stable source could not take (code review,
        // 2026-10-29): listed until Publish again finishes it; never rejected, never replaced.
        public const string InBeta = "in_beta";

        public const string Published = "published";

        public const string Rejected = "rejected";

        public const string Replaced = "replaced";

        // Still in the administrator's list.
        public static bool IsWaiting(string state) => state is Pending or InBeta;
    }

    public sealed record NewOscilloscopeModelSubmission(
        DateTimeOffset CreatedUtc,
        string Vendor,
        string SeriesOrModel,
        string ValuesJson,
        string? ContactEmail,
        string? ContactName,
        long? AccountId,
        string? Message,
        bool FullTestSuitePassed,
        string? ClientVersion,
        string? SenderAddress);

    public sealed record OscilloscopeModelSubmissionRecord(
        long Id,
        DateTimeOffset CreatedUtc,
        string Vendor,
        string SeriesOrModel,
        string ValuesJson,
        string? ContactEmail,
        string? ContactName,
        long? AccountId,
        string? Message,
        bool FullTestSuitePassed,
        string? ClientVersion,
        string State);
}
