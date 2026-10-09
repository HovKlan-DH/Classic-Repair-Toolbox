using CRT.Server.Handlers.Accounts;

namespace CRT.Server.Handlers.Submissions
{
    // ###########################################################################################
    // The database half of resetting the contribution data (owner request, 2026-10-04) - see
    // DataResetFlow for what goes and what stays, and why.
    //
    // ONE STORE, ONE TRANSACTION, rather than a method on each of the three stores the tables
    // belong to: a reset that stopped half way - submissions gone, accounts still there - would be a
    // state no flow was written for. MySqlDataResetStore is the real one, an untested I/O boundary
    // like the other MySql* stores; the tests' fake keeps counts.
    // ###########################################################################################
    public interface IDataResetStore
    {
        // What a reset would delete now - read, nothing written.
        Task<DataResetCounts> CountAsync(CancellationToken cancellationToken = default);

        // ###########################################################################################
        // Deletes it all in one transaction and writes `record` - made from the counts read INSIDE
        // that transaction, so it names exactly what went - as the first row of the new history, in
        // the same transaction: the history never shows a reset that did not happen, and a reset
        // never happens without one. Answers what was there before (what was deleted) and the ids of
        // the submissions deleted, whose partial uploads the flow then clears.
        //
        // *** WHAT WAS SHOWN IS WHAT IS DELETED (code review, 2026-10-04). *** The tables the
        // fingerprint holds are LOCKED before they are counted, so nothing can be added to them until
        // the transaction ends, and `isAsShown` is asked about those counts: false rolls the
        // transaction back with nothing deleted, and answers null. The flow's own early count, on
        // another connection, could not stop a submission committed between it and the delete.
        // ###########################################################################################
        Task<DataResetResult?> ResetAsync(
            Func<DataResetCounts, bool> isAsShown,
            Func<DataResetCounts, AuditEntry> makeRecord,
            CancellationToken cancellationToken = default);
    }

    // ###########################################################################################
    // What a reset would delete, or did.
    //
    //   LastSubmissionId / LastAccountId - the highest id of each, for the fingerprint: a count alone
    //                    stays the same when one row goes and another arrives.
    //   Accounts       - the accounts that go: every one that is not an administrator.
    //   Administrators - the accounts that stay.
    //   ProductionApprovals - approvals given for a BETA state that is waiting for the stable
    //                    source. They go too; they are a maintainer's decision, like any other.
    // ###########################################################################################
    public sealed record DataResetCounts(
        int Submissions,
        long LastSubmissionId,
        int Accounts,
        long LastAccountId,
        int Administrators,
        int Maintainers,
        int Invitations,
        int ProductionApprovals,
        int Boards,
        int HistoryEntries,
        int BoardViews,
        int ApiUsageRows);

    public sealed record DataResetResult(DataResetCounts Deleted, IReadOnlyList<long> SubmissionIds);
}
