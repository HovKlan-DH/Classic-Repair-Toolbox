using CRT.Server.Handlers.Usage;
using Handlers.DataHandling;

namespace CRT.Server.Tests
{
    // ###########################################################################################
    // Covers BoardViewStatisticsRules - the Systems screen's view numbers from a system's facts.
    //
    // The windows are WHOLE UTC DAYS INCLUDING TODAY, and BETA-source views count only in their own
    // number: those are mostly maintainers checking their work, and must not make a board look used.
    // ###########################################################################################
    public sealed class BoardViewStatisticsRulesTests
    {
        // Late in the day, so "today" is not an edge by accident.
        private static readonly DateTimeOffset Now = new(2026, 9, 27, 21, 30, 0, TimeSpan.Zero);

        private static DateOnly DaysAgo(int days) => DateOnly.FromDateTime(BoardViewStatisticsRulesTests.Now.UtcDateTime.Date).AddDays(-days);

        private static BoardViewFact Fact(int daysAgo, int views, string? code = "DK", bool fromBeta = false) =>
            new(BoardViewStatisticsRulesTests.DaysAgo(daysAgo), fromBeta, code,
                code switch { "DK" => "Denmark", "DE" => "Germany", "US" => "United States", "SE" => "Sweden", "NO" => "Norway", "FI" => "Finland", _ => null },
                views);

        // Today and the six days before are "the last 7 days"; the eighth day back is not.
        [Fact]
        public void The_windows_are_whole_days_counting_today()
        {
            BoardViewStatistics stats = BoardViewStatisticsRules.Build(
                [
                    BoardViewStatisticsRulesTests.Fact(0, 1),
                    BoardViewStatisticsRulesTests.Fact(6, 2),
                    BoardViewStatisticsRulesTests.Fact(7, 4),
                    BoardViewStatisticsRulesTests.Fact(29, 8),
                    BoardViewStatisticsRulesTests.Fact(30, 16),
                    BoardViewStatisticsRulesTests.Fact(364, 32),
                    BoardViewStatisticsRulesTests.Fact(365, 64)
                ],
                BoardViewStatisticsRulesTests.Now);

            Assert.Equal(1 + 2, stats.Last7Days);
            Assert.Equal(1 + 2 + 4 + 8, stats.Last30Days);
            Assert.Equal(1 + 2 + 4 + 8 + 16 + 32, stats.Last365Days);
        }

        // The window's first moment is midnight UTC of its first day - what the list's count asks the
        // store from, so the list and the detail count the same days.
        [Fact]
        public void A_window_starts_at_midnight_utc_of_its_first_day()
        {
            Assert.Equal(new DateTimeOffset(2026, 9, 27, 0, 0, 0, TimeSpan.Zero), BoardViewStatisticsRules.WindowStart(BoardViewStatisticsRulesTests.Now, 1));
            Assert.Equal(new DateTimeOffset(2026, 8, 29, 0, 0, 0, TimeSpan.Zero), BoardViewStatisticsRules.WindowStart(BoardViewStatisticsRulesTests.Now, 30));
        }

        [Fact]
        public void Beta_source_views_count_only_in_their_own_number()
        {
            BoardViewStatistics stats = BoardViewStatisticsRules.Build(
                [
                    BoardViewStatisticsRulesTests.Fact(1, 5),
                    BoardViewStatisticsRulesTests.Fact(1, 7, fromBeta: true),
                    BoardViewStatisticsRulesTests.Fact(40, 11, fromBeta: true)
                ],
                BoardViewStatisticsRulesTests.Now);

            Assert.Equal((5, 5, 5), (stats.Last7Days, stats.Last30Days, stats.Last365Days));
            Assert.Equal(7, stats.FromBetaLast30Days);
            Assert.Equal(5, Assert.Single(stats.TopCountries).Views);
        }

        // ###########################################################################################
        // The countries: the last year, most views first, at most five, summed across days - and a
        // view with no country counts in the totals but names no country.
        // ###########################################################################################
        [Fact]
        public void The_top_countries_are_the_years_five_most_summed_across_days()
        {
            BoardViewStatistics stats = BoardViewStatisticsRules.Build(
                [
                    BoardViewStatisticsRulesTests.Fact(1, 10, "DE"),
                    BoardViewStatisticsRulesTests.Fact(100, 15, "DE"),
                    BoardViewStatisticsRulesTests.Fact(2, 20, "DK"),
                    BoardViewStatisticsRulesTests.Fact(3, 5, "US"),
                    BoardViewStatisticsRulesTests.Fact(3, 4, "SE"),
                    BoardViewStatisticsRulesTests.Fact(3, 3, "NO"),
                    BoardViewStatisticsRulesTests.Fact(3, 2, "FI"),
                    BoardViewStatisticsRulesTests.Fact(3, 50, code: null),
                    BoardViewStatisticsRulesTests.Fact(400, 99, "US")
                ],
                BoardViewStatisticsRulesTests.Now);

            Assert.Equal(
                ["Germany 25", "Denmark 20", "United States 5", "Sweden 4", "Norway 3"],
                stats.TopCountries.Select(country => $"{country.CountryName} {country.Views}"));

            Assert.Equal(10 + 15 + 20 + 5 + 4 + 3 + 2 + 50, stats.Last365Days);
        }

        [Fact]
        public void No_views_at_all_is_all_zero_and_no_countries()
        {
            BoardViewStatistics stats = BoardViewStatisticsRules.Build([], BoardViewStatisticsRulesTests.Now);

            Assert.Equal((0, 0, 0, 0), (stats.Last7Days, stats.Last30Days, stats.Last365Days, stats.FromBetaLast30Days));
            Assert.Empty(stats.TopCountries);
        }
    }
}
