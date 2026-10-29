using CRT.Server.Handlers.Submissions;
using Handlers.DataHandling;

namespace CRT.Server.Tests.Fakes
{
    // ###########################################################################################
    // IOscilloscopeModelStore in memory - MySqlOscilloscopeModelStore's rules as the flows rely on
    // them: an added model waits, an earlier one of the same name from the same contributor still
    // pending is replaced (one in BETA never is; without an account, the same contributor is the same
    // address AND connection), a decision lands only on a model still waiting. Names are compared by
    // OscilloscopeModelRules.SameName, as the MySQL store compares them (never by its collation).
    // ###########################################################################################
    public sealed class FakeOscilloscopeModelStore : IOscilloscopeModelStore
    {
        private long thisNextId = 1;

        public List<(OscilloscopeModelSubmissionRecord Record, string? SenderAddress, string? Note, long? DecidedBy)> Rows { get; } = [];

        public OscilloscopeModelSubmissionRecord this[long id] => this.Rows.Single(row => row.Record.Id == id).Record;

        public string? NoteOf(long id) => this.Rows.Single(row => row.Record.Id == id).Note;

        public Task<long> AddAsync(NewOscilloscopeModelSubmission submission, CancellationToken cancellationToken = default)
        {
            for (int index = 0; index < this.Rows.Count; index++)
            {
                OscilloscopeModelSubmissionRecord record = this.Rows[index].Record;

                bool sameContributor = submission.AccountId is long account
                    ? record.AccountId == account
                    : record.AccountId is null &&
                      string.Equals(record.ContactEmail, submission.ContactEmail, StringComparison.OrdinalIgnoreCase) &&
                      submission.SenderAddress is not null &&
                      string.Equals(this.Rows[index].SenderAddress, submission.SenderAddress, StringComparison.Ordinal);

                if (record.State == OscilloscopeModelStates.Pending && sameContributor &&
                    OscilloscopeModelRules.SameName(record.Vendor, record.SeriesOrModel, submission.Vendor, submission.SeriesOrModel))
                {
                    this.Rows[index] = this.Rows[index] with { Record = record with { State = OscilloscopeModelStates.Replaced } };
                }
            }

            long id = this.thisNextId++;

            this.Rows.Add((
                new OscilloscopeModelSubmissionRecord(
                    id,
                    submission.CreatedUtc,
                    submission.Vendor,
                    submission.SeriesOrModel,
                    submission.ValuesJson,
                    submission.ContactEmail,
                    submission.ContactName,
                    submission.AccountId,
                    submission.Message,
                    submission.FullTestSuitePassed,
                    submission.ClientVersion,
                    OscilloscopeModelStates.Pending),
                submission.SenderAddress,
                null,
                null));

            return Task.FromResult(id);
        }

        // Awaited while counting - a test's way to have requests count at the same time.
        public Func<Task>? WhileCounting { get; set; }

        public async Task<int> CountFromAddressSinceAsync(string senderAddress, DateTimeOffset sinceUtc, CancellationToken cancellationToken = default)
        {
            if (this.WhileCounting is { } meanwhile)
                await meanwhile();

            return this.Rows.Count(row => row.SenderAddress == senderAddress && row.Record.CreatedUtc > sinceUtc);
        }

        public Task<IReadOnlyList<OscilloscopeModelSubmissionRecord>> ListPendingAsync(CancellationToken cancellationToken = default) =>
            Task.FromResult<IReadOnlyList<OscilloscopeModelSubmissionRecord>>(this.Rows
                .Select(row => row.Record)
                .Where(record => OscilloscopeModelStates.IsWaiting(record.State))
                .OrderBy(record => record.CreatedUtc)
                .ThenBy(record => record.Id)
                .ToList());

        public Task<int> CountWaitingAsync(CancellationToken cancellationToken = default) =>
            Task.FromResult(this.Rows.Count(row => OscilloscopeModelStates.IsWaiting(row.Record.State)));

        public Task<OscilloscopeModelSubmissionRecord?> FindAsync(long id, CancellationToken cancellationToken = default) =>
            Task.FromResult(this.Rows.Select(row => row.Record).FirstOrDefault(record => record.Id == id));

        // Run just before a decision lands - a test's way to have something happen meanwhile.
        public Action? BeforeDecide { get; set; }

        // Run just before a model is marked in BETA - the same, for a publish stable did not take.
        public Action? BeforeMarkInBeta { get; set; }

        public Task<bool> MarkInBetaAsync(long id, CancellationToken cancellationToken = default)
        {
            Action? meanwhile = this.BeforeMarkInBeta;
            this.BeforeMarkInBeta = null;
            meanwhile?.Invoke();

            int index = this.Rows.FindIndex(row => row.Record.Id == id);

            if (index < 0 || this.Rows[index].Record.State is not (OscilloscopeModelStates.Pending or OscilloscopeModelStates.Replaced))
                return Task.FromResult(false);

            this.Rows[index] = this.Rows[index] with { Record = this.Rows[index].Record with { State = OscilloscopeModelStates.InBeta } };

            return Task.FromResult(true);
        }

        // Set: ReplaceInBetaOfNameAsync throws it - a test's way to have the database fail there.
        public Exception? FailReplaceInBeta { get; set; }

        public Task<IReadOnlyList<OscilloscopeModelSubmissionRecord>> ReplaceInBetaOfNameAsync(string vendor, string seriesOrModel, long exceptId, CancellationToken cancellationToken = default)
        {
            if (this.FailReplaceInBeta is { } failure)
                throw failure;

            var replaced = new List<OscilloscopeModelSubmissionRecord>();

            for (int index = 0; index < this.Rows.Count; index++)
            {
                OscilloscopeModelSubmissionRecord record = this.Rows[index].Record;

                if (record.Id != exceptId && record.State == OscilloscopeModelStates.InBeta &&
                    OscilloscopeModelRules.SameName(record.Vendor, record.SeriesOrModel, vendor, seriesOrModel))
                {
                    this.Rows[index] = this.Rows[index] with { Record = record with { State = OscilloscopeModelStates.Replaced } };
                    replaced.Add(record);
                }
            }

            return Task.FromResult<IReadOnlyList<OscilloscopeModelSubmissionRecord>>(replaced);
        }

        public Task<bool> DecideAsync(long id, string state, long decidedByAccountId, string? note, DateTimeOffset whenUtc, bool evenIfReplaced = false, CancellationToken cancellationToken = default)
        {
            Action? meanwhile = this.BeforeDecide;
            this.BeforeDecide = null;
            meanwhile?.Invoke();

            int index = this.Rows.FindIndex(row => row.Record.Id == id);

            if (index < 0 ||
                !(OscilloscopeModelStates.IsWaiting(this.Rows[index].Record.State) ||
                  (evenIfReplaced && this.Rows[index].Record.State == OscilloscopeModelStates.Replaced)))
            {
                return Task.FromResult(false);
            }

            this.Rows[index] = (this.Rows[index].Record with { State = state }, this.Rows[index].SenderAddress, note, decidedByAccountId);

            return Task.FromResult(true);
        }
    }
}
