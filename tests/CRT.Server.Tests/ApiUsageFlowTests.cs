using CRT.Server.Handlers.Accounts;
using CRT.Server.Handlers.Submissions;
using CRT.Server.Handlers.Usage;
using CRT.Server.Tests.Fakes;
using Handlers.DataHandling;
using Microsoft.Extensions.Logging.Abstractions;
using Xunit;

namespace CRT.Server.Tests
{
    // ###########################################################################################
    // Covers ApiUsageFlow - Account > "API usage" (owner request, 2026-10-04). Administrator only; the
    // window counts today as one of its days; the launches are a second opinion that never costs the
    // routes when the check-in table cannot be read.
    // ###########################################################################################
    public sealed class ApiUsageFlowTests
    {
        private static readonly DateTimeOffset Now = new(2026, 10, 4, 12, 0, 0, TimeSpan.Zero);

        private static ReviewAccess Access(bool administrator) =>
            ReviewAccess.For(new AccountRecord(1, "a@example.com", "a@example.com", "hash", "A",
                IsVerified: true, IsAdministrator: administrator, IsLocked: false, ApiUsageFlowTests.Now, null));

        private static Task<ApiUsageAnswer?> ReadAsync(ReviewAccess? access, int? days, FakeApiUsageStore store, FakeCheckInStore checkIns) =>
            ApiUsageFlow.ReadAsync(access, days, [("GET", "/api/review/queue")], store, checkIns, ApiUsageFlowTests.Now, NullLogger.Instance);

        [Fact]
        public async Task Only_an_administrator_sees_the_usage()
        {
            var store = new FakeApiUsageStore();

            Assert.Null(await ApiUsageFlowTests.ReadAsync(ApiUsageFlowTests.Access(administrator: false), null, store, new FakeCheckInStore()));
            Assert.Null(await ApiUsageFlowTests.ReadAsync(null, null, store, new FakeCheckInStore()));
            Assert.Null(store.FirstDayAsked);
        }

        // ###########################################################################################
        // Ninety days when none are asked for, today being the last of them - and the launches are
        // counted over the SAME UTC days as the route calls (code review, 2026-10-04: they were a
        // rolling window in the database's local time, so with one day asked for the two columns
        // on one screen measured different periods).
        // ###########################################################################################
        [Fact]
        public async Task The_window_is_the_days_asked_for_ending_today_for_routes_and_launches_alike()
        {
            var store = new FakeApiUsageStore();
            var checkIns = new FakeCheckInStore();

            ApiUsageAnswer? answer = await ApiUsageFlowTests.ReadAsync(ApiUsageFlowTests.Access(administrator: true), null, store, checkIns);

            Assert.Equal(90, answer!.Days);
            Assert.Equal(new DateOnly(2026, 7, 7), store.FirstDayAsked);
            Assert.Equal(new DateOnly(2026, 7, 7), checkIns.FirstDayAsked);

            await ApiUsageFlowTests.ReadAsync(ApiUsageFlowTests.Access(administrator: true), 1, store, checkIns);

            Assert.Equal(new DateOnly(2026, 10, 4), store.FirstDayAsked);
            Assert.Equal(new DateOnly(2026, 10, 4), checkIns.FirstDayAsked);
        }

        [Fact]
        public async Task A_check_in_table_that_cannot_be_read_leaves_out_the_launches_not_the_routes()
        {
            var store = new FakeApiUsageStore();
            store.Rows.Add(new ApiUsageRow("GET", "/api/review/queue", "3.0.0", 12, ApiUsageFlowTests.Now));

            var checkIns = new FakeCheckInStore { LaunchFailure = new InvalidOperationException("crt_update is missing.") };

            ApiUsageAnswer? answer = await ApiUsageFlowTests.ReadAsync(ApiUsageFlowTests.Access(administrator: true), 30, store, checkIns);

            Assert.Empty(answer!.Installations);
            Assert.Equal(12, Assert.Single(answer.Routes).Calls);
        }
    }
}
