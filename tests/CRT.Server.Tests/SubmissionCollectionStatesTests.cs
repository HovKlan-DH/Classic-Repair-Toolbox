using System.Text.RegularExpressions;
using CRT.Server.Handlers.Submissions;

namespace CRT.Server.Tests
{
    // ###########################################################################################
    // Holds the code's lists of submission states to the DATABASE's (security review, 2026-09-25),
    // and checks the schema change that stopped a board being hijacked by a case-variant id.
    //
    // *** WHY READ THE MIGRATIONS. *** The collectors keep a submission's files when its state is
    // "live" and sweep them when it is "retired". A state added to the CHECK constraint and to
    // neither list would keep its files for ever; one added to both would have them swept while a
    // reviewer still needs them. Every test of the flows runs against an in-memory store, so nothing
    // else would notice the database growing a state the code does not know about - the same blind
    // spot migration 0004 was written to correct. The migration files are copied beside the tests.
    // ###########################################################################################
    public sealed class SubmissionCollectionStatesTests
    {
        // Every migration in order, with "--" comments removed - 0004 annotates its state list with
        // comments that themselves contain parentheses, which would end the list early.
        private static string MigrationsText() =>
            Regex.Replace(
                string.Join(
                    "\n",
                    Directory.GetFiles(Path.Combine(AppContext.BaseDirectory, "Migrations"), "*.sql")
                        .OrderBy(path => path, StringComparer.Ordinal)
                        .Select(File.ReadAllText)),
                "--[^\n]*",
                string.Empty);

        // The states the LATEST ck_submissions_state constraint allows - it is dropped and re-added
        // by later migrations, so the last definition is the one in force.
        private static IReadOnlySet<string> StatesTheDatabaseAllows()
        {
            string sql = SubmissionCollectionStatesTests.MigrationsText();

            MatchCollection definitions = Regex.Matches(
                sql,
                @"ck_submissions_state\s+CHECK\s*\(\s*state\s+IN\s*\((?<list>[^)]*)\)",
                RegexOptions.IgnoreCase);

            Assert.NotEmpty(definitions);

            return Regex.Matches(definitions[^1].Groups["list"].Value, "'(?<state>[a-z_]+)'")
                .Select(match => match.Groups["state"].Value)
                .ToHashSet(StringComparer.Ordinal);
        }

        [Fact]
        public void The_code_knows_exactly_the_states_the_database_allows()
        {
            Assert.Equal(
                SubmissionCollectionStatesTests.StatesTheDatabaseAllows().Order(StringComparer.Ordinal),
                SubmissionState.All.Order(StringComparer.Ordinal));
        }

        [Fact]
        public void Every_state_is_either_live_or_retired_and_never_both()
        {
            Assert.Empty(SubmissionCollectionStates.Live.Intersect(SubmissionCollectionStates.Retired));

            Assert.Equal(
                SubmissionState.All.Order(StringComparer.Ordinal),
                SubmissionCollectionStates.Live.Concat(SubmissionCollectionStates.Retired).Order(StringComparer.Ordinal));
        }

        // MySqlSubmissionStore binds each list as exactly four parameters.
        [Fact]
        public void Each_list_has_the_four_entries_the_store_binds()
        {
            Assert.Equal(4, SubmissionCollectionStates.Live.Count);
            Assert.Equal(4, SubmissionCollectionStates.Retired.Count);
        }

        // Anything awaiting or past a decision to publish keeps its files.
        [Fact]
        public void Pending_approved_and_merged_are_live()
        {
            Assert.Contains(SubmissionState.Pending, SubmissionCollectionStates.Live);
            Assert.Contains(SubmissionState.Approved, SubmissionCollectionStates.Live);
            Assert.Contains(SubmissionState.Merged, SubmissionCollectionStates.Live);
            Assert.Contains(SubmissionState.Uploading, SubmissionCollectionStates.Live);
        }

        // ###########################################################################################
        // *** THE SYSTEM ID IS BYTE-FOR-BYTE IN EVERY TABLE THAT HOLDS IT. *** With a case-folding
        // key, an anonymous submission for "commodore/c64/250425" created the row every later real
        // submission for that board attached to - and was then rejected. All three columns must
        // change together, because MariaDB requires a foreign key and its target to share a
        // collation.
        // ###########################################################################################
        [Theory]
        [InlineData("systems")]
        [InlineData("submissions")]
        [InlineData("maintainers")]
        public void The_system_id_is_binary_collated_in(string table)
        {
            Assert.Matches(
                new Regex(
                    $@"ALTER\s+TABLE\s+{table}\s+MODIFY\s+system_id\s+VARCHAR\(255\)\s+NOT\s+NULL\s+COLLATE\s+utf8mb4_bin",
                    RegexOptions.IgnoreCase),
                SubmissionCollectionStatesTests.MigrationsText());
        }

        // The per-address byte budget reads this column; the store writes it.
        [Fact]
        public void Submissions_record_the_bytes_they_asked_to_upload()
        {
            Assert.Matches(
                new Regex(@"ADD\s+COLUMN\s+bytes_to_upload\s+BIGINT", RegexOptions.IgnoreCase),
                SubmissionCollectionStatesTests.MigrationsText());
        }
    }
}
