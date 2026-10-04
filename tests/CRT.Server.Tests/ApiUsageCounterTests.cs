using CRT.Server.Handlers.Usage;
using Handlers.DataHandling;
using Xunit;

namespace CRT.Server.Tests
{
    // ###########################################################################################
    // Covers ApiUsageCounter - the API usage counted in memory between two writes (owner request,
    // 2026-10-04). What matters: a day's calls of one route by one version add up into one tally,
    // taking empties the counter, a failed write put back loses nothing, and a sender making up
    // versions cannot grow it without end.
    // ###########################################################################################
    public sealed class ApiUsageCounterTests
    {
        private static readonly DateTimeOffset Morning = new(2026, 10, 4, 8, 0, 0, TimeSpan.Zero);

        [Fact]
        public void Calls_of_one_route_by_one_version_on_one_day_add_up_to_one_tally_with_the_last_time()
        {
            var counter = new ApiUsageCounter();

            counter.Count("GET", "/api/review/queue", "3.0.0", ApiUsageCounterTests.Morning);
            counter.Count("GET", "/api/review/queue", "3.0.0", ApiUsageCounterTests.Morning.AddMinutes(5));
            counter.Count("GET", "/api/review/queue", "3.0.0", ApiUsageCounterTests.Morning.AddMinutes(2));

            ApiUsageTally tally = Assert.Single(counter.TakeAll());

            Assert.Equal((new DateOnly(2026, 10, 4), "GET", "/api/review/queue", "3.0.0", 3L),
                (tally.Day, tally.Method, tally.Route, tally.Version, tally.Calls));
            Assert.Equal(ApiUsageCounterTests.Morning.AddMinutes(5), tally.LastUtc);
        }

        // Another version, method, route or UTC day is a tally of its own.
        [Fact]
        public void Each_version_method_route_and_day_is_counted_apart()
        {
            var counter = new ApiUsageCounter();

            counter.Count("GET", "/api/review/queue", "3.0.0", ApiUsageCounterTests.Morning);
            counter.Count("GET", "/api/review/queue", "3.1.0", ApiUsageCounterTests.Morning);
            counter.Count("POST", "/api/review/queue", "3.0.0", ApiUsageCounterTests.Morning);
            counter.Count("GET", "/api/review/production", "3.0.0", ApiUsageCounterTests.Morning);
            counter.Count("GET", "/api/review/queue", "3.0.0", ApiUsageCounterTests.Morning.AddHours(17));

            IReadOnlyList<ApiUsageTally> tallies = counter.TakeAll();

            Assert.Equal(5, tallies.Count);
            Assert.All(tallies, tally => Assert.Equal(1, tally.Calls));
            Assert.Contains(tallies, tally => tally.Day == new DateOnly(2026, 10, 5));
        }

        [Fact]
        public void Taking_the_tallies_leaves_the_counter_empty()
        {
            var counter = new ApiUsageCounter();
            counter.Count("GET", "/api/health", ApiUsageVersion.NotCrt, ApiUsageCounterTests.Morning);

            Assert.Single(counter.TakeAll());
            Assert.Empty(counter.TakeAll());
        }

        // ###########################################################################################
        // A write the database refused puts its tallies back - added to whatever was counted since,
        // not replacing it - so the next write carries both.
        // ###########################################################################################
        [Fact]
        public void Tallies_put_back_add_to_what_was_counted_since()
        {
            var counter = new ApiUsageCounter();
            counter.Count("GET", "/api/review/queue", "3.0.0", ApiUsageCounterTests.Morning);
            counter.Count("GET", "/api/review/queue", "3.0.0", ApiUsageCounterTests.Morning);

            IReadOnlyList<ApiUsageTally> failed = counter.TakeAll();

            counter.Count("GET", "/api/review/queue", "3.0.0", ApiUsageCounterTests.Morning.AddMinutes(6));
            counter.PutBack(failed);

            ApiUsageTally tally = Assert.Single(counter.TakeAll());
            Assert.Equal(3, tally.Calls);
            Assert.Equal(ApiUsageCounterTests.Morning.AddMinutes(6), tally.LastUtc);
        }

        [Fact]
        public void Clearing_forgets_everything_not_yet_written()
        {
            var counter = new ApiUsageCounter();
            counter.Count("GET", "/api/review/queue", "3.0.0", ApiUsageCounterTests.Morning);

            counter.Clear();

            Assert.Empty(counter.TakeAll());
        }

        // ###########################################################################################
        // *** BOUNDED. *** Past MaximumKeys waiting tallies, a NEW key counts its version as
        // "(other)" - while a key already there keeps counting as itself. Reached here through
        // routes times versions, since the versions alone stop at MaximumVersionsPerDay (below).
        // ###########################################################################################
        [Fact]
        public void Past_the_key_cap_a_new_key_counts_as_other_while_known_ones_keep_counting()
        {
            var counter = new ApiUsageCounter();
            int routes = ApiUsageCounter.MaximumKeys / ApiUsageCounter.MaximumVersionsPerDay;

            for (int route = 0; route < routes; route++)
            {
                for (int version = 0; version < ApiUsageCounter.MaximumVersionsPerDay; version++)
                    counter.Count("GET", $"/api/route{route}", $"1.0.{version}", ApiUsageCounterTests.Morning);
            }

            counter.Count("GET", "/api/one-more", "1.0.5", ApiUsageCounterTests.Morning);
            counter.Count("GET", "/api/one-more", "1.0.6", ApiUsageCounterTests.Morning);
            counter.Count("GET", "/api/route0", "1.0.0", ApiUsageCounterTests.Morning);

            IReadOnlyList<ApiUsageTally> tallies = counter.TakeAll();

            Assert.Equal(ApiUsageCounter.MaximumKeys + 1, tallies.Count);
            Assert.Equal(2, Assert.Single(tallies, tally => tally.Version == ApiUsageVersion.Other).Calls);
            Assert.Equal(2, Assert.Single(tallies, tally => tally.Route == "/api/route0" && tally.Version == "1.0.0").Calls);
        }

        // ###########################################################################################
        // *** BOUNDED PER DAY, ACROSS WRITES (code review, 2026-10-04). *** The key cap emptied with
        // every write, so versions made up all day ("CRT 9.9.1", "CRT 9.9.2" ...) were up to 10,000
        // new rows every five minutes. The day's versions are remembered across writes: past
        // MaximumVersionsPerDay a new one counts as "(other)", a known one still as itself.
        // ###########################################################################################
        [Fact]
        public void Past_the_days_version_cap_a_new_version_counts_as_other_even_after_a_write()
        {
            var counter = new ApiUsageCounter();

            for (int index = 0; index < ApiUsageCounter.MaximumVersionsPerDay; index++)
            {
                counter.Count("GET", "/api/health", $"9.9.{index}", ApiUsageCounterTests.Morning);

                // A write in between, as the flusher's five minutes would bring.
                if (index % 10 == 0)
                    counter.TakeAll();
            }

            counter.TakeAll();

            counter.Count("GET", "/api/health", "9.9.1000", ApiUsageCounterTests.Morning.AddHours(1));
            counter.Count("GET", "/api/review/queue", "9.9.1001", ApiUsageCounterTests.Morning.AddHours(1));
            counter.Count("GET", "/api/health", "9.9.5", ApiUsageCounterTests.Morning.AddHours(1));

            IReadOnlyList<ApiUsageTally> tallies = counter.TakeAll();

            Assert.Equal(["(other)", "(other)", "9.9.5"], tallies.Select(tally => tally.Version).Order(StringComparer.Ordinal));
        }

        // A new UTC day has room again, and a request naming no CRT is never held to the cap.
        [Fact]
        public void A_new_day_has_room_again_and_requests_naming_no_CRT_are_never_held_to_the_cap()
        {
            var counter = new ApiUsageCounter();

            for (int index = 0; index < ApiUsageCounter.MaximumVersionsPerDay; index++)
                counter.Count("GET", "/api/health", $"9.9.{index}", ApiUsageCounterTests.Morning);

            counter.Count("GET", "/api/health", ApiUsageVersion.NotCrt, ApiUsageCounterTests.Morning);
            counter.Count("GET", "/api/health", "3.0.0", ApiUsageCounterTests.Morning.AddDays(1));

            IReadOnlyList<ApiUsageTally> tallies = counter.TakeAll();

            Assert.Contains(tallies, tally => tally.Version == ApiUsageVersion.NotCrt);
            Assert.Contains(tallies, tally => tally.Version == "3.0.0" && tally.Day == new DateOnly(2026, 10, 5));
            Assert.DoesNotContain(tallies, tally => tally.Version == ApiUsageVersion.Other);
        }
    }
}
