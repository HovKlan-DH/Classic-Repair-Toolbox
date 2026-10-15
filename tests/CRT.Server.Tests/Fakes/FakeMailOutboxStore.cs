using CRT.Server.Handlers.Email;

namespace CRT.Server.Tests.Fakes
{
    // ###########################################################################################
    // An in-memory IMailOutboxStore (code review, 2026-10-13) - the real store's statements as list
    // operations: an AUTO_INCREMENT id, "due" meaning next_attempt_utc reached AND words kept, and
    // giving up by creation time.
    // ###########################################################################################
    public sealed class FakeMailOutboxStore : IMailOutboxStore
    {
        private long thisNextId = 1;

        public List<OutboxMail> Mails { get; } = [];

        public bool FailAdd { get; set; }

        public Task AddAsync(OutboxMail mail, CancellationToken cancellationToken = default)
        {
            if (this.FailAdd)
                throw new InvalidOperationException("The database is not reachable.");

            this.Mails.Add(mail with { Id = this.thisNextId++ });
            return Task.CompletedTask;
        }

        public Task<IReadOnlyList<OutboxMail>> GetDueAsync(DateTimeOffset nowUtc, int limit, CancellationToken cancellationToken = default) =>
            Task.FromResult<IReadOnlyList<OutboxMail>>(this.Mails
                .Where(mail => mail.NextAttemptUtc <= nowUtc && mail.Body is not null)
                .OrderBy(mail => mail.Id)
                .Take(limit)
                .ToList());

        public Task RemoveAsync(long id, CancellationToken cancellationToken = default)
        {
            this.Mails.RemoveAll(mail => mail.Id == id);
            return Task.CompletedTask;
        }

        public Task RecordFailureAsync(long id, int attempts, DateTimeOffset nextAttemptUtc, string? error, CancellationToken cancellationToken = default)
        {
            int index = this.Mails.FindIndex(mail => mail.Id == id);

            if (index >= 0)
                this.Mails[index] = this.Mails[index] with { Attempts = attempts, NextAttemptUtc = nextAttemptUtc, LastError = error };

            return Task.CompletedTask;
        }

        public Task<int> RemoveCreatedBeforeAsync(DateTimeOffset createdBefore, CancellationToken cancellationToken = default) =>
            Task.FromResult(this.Mails.RemoveAll(mail => mail.CreatedUtc < createdBefore));

        public Task<int> CountAsync(CancellationToken cancellationToken = default) => Task.FromResult(this.Mails.Count);
    }
}
