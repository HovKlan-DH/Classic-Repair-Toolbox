using CRT.Server.Handlers.Accounts;
using CRT.Server.Handlers.Submissions;

namespace CRT.Server.Tests.Fakes
{
    // ###########################################################################################
    // An in-memory IDataResetStore: it holds the counts it is given and answers with them, and a
    // reset empties them - keeping the administrators - and remembers the history row it was asked
    // to write. Failure, when set, is thrown by ResetAsync with nothing changed, as the transaction
    // would roll back. ArrivesBeforeReset runs inside ResetAsync before its counts are read - a row
    // another request committed between the flow's early count and the transaction - and AfterReset
    // once the reset is committed.
    // ###########################################################################################
    public sealed class FakeDataResetStore : IDataResetStore
    {
        public FakeDataResetStore(DataResetCounts counts, IReadOnlyList<long>? submissionIds = null)
        {
            this.Counts = counts;
            this.SubmissionIds = submissionIds ?? [];
        }

        public DataResetCounts Counts { get; set; }

        public IReadOnlyList<long> SubmissionIds { get; set; }

        public Exception? Failure { get; set; }

        public Action? ArrivesBeforeReset { get; set; }

        public Action? AfterReset { get; set; }

        public int Resets { get; private set; }

        public AuditEntry? Record { get; private set; }

        public Task<DataResetCounts> CountAsync(CancellationToken cancellationToken = default) =>
            Task.FromResult(this.Counts);

        public Task<DataResetResult?> ResetAsync(
            Func<DataResetCounts, bool> isAsShown,
            Func<DataResetCounts, AuditEntry> makeRecord,
            CancellationToken cancellationToken = default)
        {
            if (this.Failure is not null)
                throw this.Failure;

            this.ArrivesBeforeReset?.Invoke();

            DataResetCounts before = this.Counts;
            IReadOnlyList<long> ids = this.SubmissionIds;

            // Rolled back with nothing deleted, as the real transaction is.
            if (!isAsShown(before))
                return Task.FromResult<DataResetResult?>(null);

            this.Record = makeRecord(before);
            this.Resets++;

            // Everything gone but the administrators; the new history is the one row just written.
            this.Counts = new DataResetCounts(0, 0, 0, 0, before.Administrators, 0, 0, 0, 0, 1, 0, 0);
            this.SubmissionIds = [];

            this.AfterReset?.Invoke();

            return Task.FromResult<DataResetResult?>(new DataResetResult(before, ids));
        }
    }
}
