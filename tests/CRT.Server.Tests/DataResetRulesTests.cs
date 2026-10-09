using CRT.Server.Handlers.Submissions;
using Xunit;

namespace CRT.Server.Tests
{
    // ###########################################################################################
    // Covers DataResetRules - above all, what the fingerprint a reset is held to covers. It must
    // move when somebody's work arrives or goes (a reset confirmed against old counts would delete
    // what the administrator never saw), and must NOT move for what grows by itself (every reset
    // would then be refused as "changed").
    // ###########################################################################################
    public sealed class DataResetRulesTests
    {
        private static readonly DataResetCounts Counts =
            new(Submissions: 14, LastSubmissionId: 52, Accounts: 6, LastAccountId: 9, Administrators: 1,
                Maintainers: 4, Invitations: 2, ProductionApprovals: 1, Boards: 7, HistoryEntries: 310,
                BoardViews: 1200, ApiUsageRows: 80);

        [Fact]
        public void The_same_counts_give_the_same_fingerprint()
        {
            Assert.Equal(DataResetRules.Fingerprint(DataResetRulesTests.Counts), DataResetRules.Fingerprint(DataResetRulesTests.Counts with { }));
            Assert.Equal(64, DataResetRules.Fingerprint(DataResetRulesTests.Counts).Length);
        }

        // ###########################################################################################
        // One submission gone and another arrived leaves the COUNT as it was - which is why the
        // highest id is in it too.
        // ###########################################################################################
        public static TheoryData<string, DataResetCounts> SomebodysWork() => new()
        {
            { "a submission more", DataResetRulesTests.Counts with { Submissions = 15 } },
            { "one submission swapped for another", DataResetRulesTests.Counts with { LastSubmissionId = 53 } },
            { "an account more", DataResetRulesTests.Counts with { Accounts = 7 } },
            { "one account swapped for another", DataResetRulesTests.Counts with { LastAccountId = 10 } },
            { "a maintainer more", DataResetRulesTests.Counts with { Maintainers = 5 } },
            { "an invitation more", DataResetRulesTests.Counts with { Invitations = 3 } },
            { "a production approval more", DataResetRulesTests.Counts with { ProductionApprovals = 2 } },
            { "a board record more", DataResetRulesTests.Counts with { Boards = 8 } },
        };

        [Theory]
        [MemberData(nameof(SomebodysWork))]
        public void Somebodys_work_arriving_or_going_moves_the_fingerprint(string what, DataResetCounts after)
        {
            Assert.True(
                DataResetRules.Fingerprint(DataResetRulesTests.Counts) != DataResetRules.Fingerprint(after),
                $"{what} should move the fingerprint");
        }

        public static TheoryData<string, DataResetCounts> GrowingByItself() => new()
        {
            { "a history row (a sign-in)", DataResetRulesTests.Counts with { HistoryEntries = 311 } },
            { "board views reported", DataResetRulesTests.Counts with { BoardViews = 1215 } },
            { "API usage written", DataResetRulesTests.Counts with { ApiUsageRows = 81 } },
            { "an administrator granted by hand", DataResetRulesTests.Counts with { Administrators = 2 } },
        };

        [Theory]
        [MemberData(nameof(GrowingByItself))]
        public void What_grows_by_itself_does_not_move_the_fingerprint(string what, DataResetCounts after)
        {
            Assert.True(
                DataResetRules.Fingerprint(DataResetRulesTests.Counts) == DataResetRules.Fingerprint(after),
                $"{what} should not move the fingerprint");
        }

        [Fact]
        public void The_plan_carries_every_count_and_says_why_only_when_switched_off()
        {
            var on = DataResetRules.ToPlan(DataResetRulesTests.Counts, isEnabled: true);
            var off = DataResetRules.ToPlan(DataResetRulesTests.Counts, isEnabled: false);

            Assert.Equal(DataResetRules.Fingerprint(DataResetRulesTests.Counts), on.Fingerprint);
            Assert.Equal((14, 6, 1, 4, 2, 7, 310, 1200, 80),
                (on.Submissions, on.Accounts, on.Administrators, on.Maintainers, on.Invitations, on.Boards, on.HistoryEntries, on.BoardViews, on.ApiUsageRows));
            Assert.True(on.IsEnabled);
            Assert.Null(on.NotEnabledBecause);

            Assert.False(off.IsEnabled);
            Assert.Equal(DataResetRules.NotEnabledMessage, off.NotEnabledBecause);
        }

        // The history row names every count, and the administrators kept.
        [Fact]
        public void The_history_row_names_what_went_and_what_was_kept()
        {
            string detail = DataResetRules.Detail(DataResetRulesTests.Counts);

            Assert.Contains("14 submission(s)", detail);
            Assert.Contains("6 account(s)", detail);
            Assert.Contains("4 maintainer(s)", detail);
            Assert.Contains("2 invitation(s)", detail);
            Assert.Contains("1 production approval(s)", detail);
            Assert.Contains("7 board record(s)", detail);
            Assert.Contains("1200 board view(s)", detail);
            Assert.Contains("80 API usage row(s)", detail);
            Assert.Contains("1 administrator(s) kept", detail);
        }
    }
}
