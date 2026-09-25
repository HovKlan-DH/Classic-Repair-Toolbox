using CRT.Server.Handlers.Database;

namespace CRT.Server.Tests
{
    // ###########################################################################################
    // Covers the rule that decides which migrations run, in what order, and when to refuse.
    //
    // Every failure this pins down is one that would otherwise be SILENT and would leave two
    // databases claiming the same version while holding different schemas - the hardest class of
    // bug to diagnose later, because the symptom appears in code that is correct.
    //
    // No database anywhere in this file, by construction: MigrationPlan is pure, which is the
    // whole reason the ordering rules live there rather than inside the runner.
    // ###########################################################################################
    public class MigrationPlanTests
    {
        private static MigrationScript Script(int number, string fileName, string sql = "CREATE TABLE t (id INT);")
        {
            return new MigrationScript(number, fileName, sql, MigrationPlan.ComputeChecksum(sql));
        }

        private static AppliedMigration Applied(int number, string fileName, string sql = "CREATE TABLE t (id INT);")
        {
            return new AppliedMigration(
                number,
                fileName,
                MigrationPlan.ComputeChecksum(sql),
                new DateTime(2026, 9, 21, 8, 0, 0, DateTimeKind.Utc));
        }

        // -----------------------------------------------------------------------------------
        // The ordinary paths.
        // -----------------------------------------------------------------------------------

        [Fact]
        public void An_empty_database_applies_every_migration_in_number_order()
        {
            // Deliberately supplied out of order: the plan must sort, not trust the input. A
            // directory walk returns whatever order the filesystem gives, which differs between
            // Windows and Linux.
            var available = new[]
            {
                Script(3, "0003_sessions.sql"),
                Script(1, "0001_initial.sql"),
                Script(2, "0002_accounts.sql")
            };

            MigrationPlanResult plan = MigrationPlan.Build(available, Array.Empty<AppliedMigration>());

            Assert.True(plan.CanProceed);
            Assert.Equal(new[] { 1, 2, 3 }, plan.ToApply.Select(script => script.Number));
        }

        [Fact]
        public void A_fully_migrated_database_applies_nothing()
        {
            var available = new[] { Script(1, "0001_initial.sql") };
            var applied = new[] { Applied(1, "0001_initial.sql") };

            MigrationPlanResult plan = MigrationPlan.Build(available, applied);

            Assert.True(plan.CanProceed);
            Assert.Empty(plan.ToApply);
        }

        [Fact]
        public void Only_the_unapplied_migrations_are_planned()
        {
            var available = new[]
            {
                Script(1, "0001_initial.sql"),
                Script(2, "0002_accounts.sql"),
                Script(3, "0003_sessions.sql")
            };

            var applied = new[] { Applied(1, "0001_initial.sql"), Applied(2, "0002_accounts.sql") };

            MigrationPlanResult plan = MigrationPlan.Build(available, applied);

            Assert.True(plan.CanProceed);
            Assert.Equal(new[] { 3 }, plan.ToApply.Select(script => script.Number));
        }

        [Fact]
        public void A_migration_is_identified_by_its_number_not_its_filename()
        {
            // Renaming a migration file must not re-run it. The number is the identity; the name
            // is a human label that may be improved.
            var available = new[] { Script(1, "0001_initial_schema.sql") };
            var applied = new[] { Applied(1, "0001_initial.sql") };

            MigrationPlanResult plan = MigrationPlan.Build(available, applied);

            Assert.True(plan.CanProceed);
            Assert.Empty(plan.ToApply);
        }

        // -----------------------------------------------------------------------------------
        // The refusals. Each of these is a silent-divergence bug if it is allowed through.
        // -----------------------------------------------------------------------------------

        [Fact]
        public void An_applied_migration_whose_content_changed_fails_the_whole_run()
        {
            // The single most common way two databases end up disagreeing while both claim to be
            // at the same version: someone fixes a typo in a migration that has already run
            // somewhere. The database it already ran against cannot be changed retroactively.
            var available = new[] { Script(1, "0001_initial.sql", "CREATE TABLE t (id BIGINT);") };
            var applied = new[] { Applied(1, "0001_initial.sql", "CREATE TABLE t (id INT);") };

            MigrationPlanResult plan = MigrationPlan.Build(available, applied);

            Assert.False(plan.CanProceed);
            Assert.Contains(plan.Failures, failure => failure.Contains("has changed since it was applied"));

            // And nothing is planned - a caller that ignores Failures must do no damage.
            Assert.Empty(plan.ToApply);
        }

        [Fact]
        public void A_gap_in_the_numbering_fails_rather_than_skipping_it()
        {
            // 0002 never got committed, or a merge dropped it. Applying 0003 onto a database
            // without it produces a schema no sequence can reproduce.
            var available = new[] { Script(1, "0001_initial.sql"), Script(3, "0003_sessions.sql") };

            MigrationPlanResult plan = MigrationPlan.Build(available, Array.Empty<AppliedMigration>());

            Assert.False(plan.CanProceed);
            Assert.Contains(plan.Failures, failure => failure.Contains("0002 is missing"));
        }

        [Fact]
        public void A_migration_numbered_below_the_highest_applied_one_fails()
        {
            // The merge case: two branches each added "the next" migration, one landed first.
            // Applying the loser now would run it out of order.
            var available = new[]
            {
                Script(1, "0001_initial.sql"),
                Script(2, "0002_from_branch_a.sql"),
                Script(3, "0003_from_branch_b.sql")
            };

            var applied = new[] { Applied(1, "0001_initial.sql"), Applied(3, "0003_from_branch_b.sql") };

            MigrationPlanResult plan = MigrationPlan.Build(available, applied);

            Assert.False(plan.CanProceed);
            Assert.Contains(plan.Failures, failure => failure.Contains("has never been applied, but migration 0003"));
        }

        [Fact]
        public void A_recorded_migration_with_no_file_fails()
        {
            // Deploying an older build over a newer database. Without this check the service
            // would start happily against a schema it does not know about.
            var available = new[] { Script(1, "0001_initial.sql") };
            var applied = new[] { Applied(1, "0001_initial.sql"), Applied(2, "0002_accounts.sql") };

            MigrationPlanResult plan = MigrationPlan.Build(available, applied);

            Assert.False(plan.CanProceed);
            Assert.Contains(plan.Failures, failure => failure.Contains("recorded as applied but its file is no longer present"));
        }

        [Fact]
        public void Two_files_sharing_a_number_fail_and_name_both()
        {
            var available = new[] { Script(2, "0002_accounts.sql"), Script(2, "0002_sessions.sql") };

            MigrationPlanResult plan = MigrationPlan.Build(
                new[] { Script(1, "0001_initial.sql") }.Concat(available).ToList(),
                Array.Empty<AppliedMigration>());

            Assert.False(plan.CanProceed);

            string failure = Assert.Single(plan.Failures, message => message.Contains("used by more than one file"));
            Assert.Contains("0002_accounts.sql", failure);
            Assert.Contains("0002_sessions.sql", failure);
        }

        [Fact]
        public void Every_problem_is_reported_in_one_pass()
        {
            // Same reasoning as ServerOptionsValidator: a maintainer fixing one problem,
            // restarting, and discovering the next is a slow loop over an SSH session.
            var available = new[]
            {
                Script(1, "0001_initial.sql", "CREATE TABLE a (id INT);"),
                Script(3, "0003_sessions.sql")
            };

            var applied = new[] { Applied(1, "0001_initial.sql", "CREATE TABLE b (id INT);") };

            MigrationPlanResult plan = MigrationPlan.Build(available, applied);

            Assert.False(plan.CanProceed);
            Assert.Contains(plan.Failures, failure => failure.Contains("has changed since it was applied"));
            Assert.Contains(plan.Failures, failure => failure.Contains("0002 is missing"));
        }

        // -----------------------------------------------------------------------------------
        // Checksums.
        // -----------------------------------------------------------------------------------

        [Fact]
        public void Line_endings_do_not_change_the_checksum()
        {
            // These files are edited on Windows and executed on Linux, and git's autocrlf can
            // rewrite them in transit. Without normalisation every migration would look modified
            // on the other platform - which would train the maintainer to ignore the one rule
            // most worth heeding.
            string windows = "CREATE TABLE t (\r\n  id INT\r\n);\r\n";
            string unix = "CREATE TABLE t (\n  id INT\n);\n";

            Assert.Equal(MigrationPlan.ComputeChecksum(windows), MigrationPlan.ComputeChecksum(unix));
        }

        [Fact]
        public void A_real_content_change_does_change_the_checksum()
        {
            // The other half of the above: normalisation must not be so aggressive that it hides
            // an actual edit.
            Assert.NotEqual(
                MigrationPlan.ComputeChecksum("CREATE TABLE t (id INT);"),
                MigrationPlan.ComputeChecksum("CREATE TABLE t (id BIGINT);"));
        }

        // -----------------------------------------------------------------------------------
        // Filename parsing.
        // -----------------------------------------------------------------------------------

        [Theory]
        [InlineData("0001_initial.sql", 1)]
        [InlineData("0002_accounts.sql", 2)]
        [InlineData("0042_something_with_underscores.sql", 42)]
        [InlineData("1_short.sql", 1)]
        public void A_well_formed_migration_filename_yields_its_number(string fileName, int expected)
        {
            Assert.True(MigrationPlan.TryParseNumber(fileName, out int number));
            Assert.Equal(expected, number);
        }

        [Theory]
        [InlineData("0001.sql")]              // no descriptive name - unidentifiable in a listing
        [InlineData("initial.sql")]           // no number at all
        [InlineData("_initial.sql")]          // empty number
        [InlineData("-1_backwards.sql")]      // would sort before everything
        [InlineData("0000_zero.sql")]         // migrations start at 1
        [InlineData("")]
        public void A_malformed_migration_filename_is_rejected(string fileName)
        {
            Assert.False(MigrationPlan.TryParseNumber(fileName, out _));
        }
    }
}
