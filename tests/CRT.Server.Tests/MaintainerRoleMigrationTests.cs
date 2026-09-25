using System.Reflection;
using System.Text.RegularExpressions;
using CRT.Server.Handlers.Submissions;
using Handlers.DataHandling;

namespace CRT.Server.Tests
{
    // ###########################################################################################
    // Holds the code's ROLE WORDS to the DATABASE's (owner decision, 2026-09-25: "reviewer" became
    // "maintainer" everywhere, migration 0010).
    //
    // The approval tables store the role as text and a CHECK constraint allows exactly two words.
    // MySqlSubmissionStore.RoleText writes them. Every flow test runs against the in-memory store, so
    // a word renamed in the code and not in the database - or the other way round - passed the whole
    // suite and failed every approval insert on the server. This reads the migrations, as
    // SubmissionCollectionStatesTests does for the submission states, and fails on either half moving
    // alone.
    // ###########################################################################################
    public sealed class MaintainerRoleMigrationTests
    {
        // Every migration in order, with "--" comments removed (their prose names old words freely).
        private static string MigrationsText() =>
            Regex.Replace(
                string.Join(
                    "\n",
                    Directory.GetFiles(Path.Combine(AppContext.BaseDirectory, "Migrations"), "*.sql")
                        .OrderBy(path => path, StringComparer.Ordinal)
                        .Select(File.ReadAllText)),
                "--[^\n]*",
                string.Empty);

        // The words the LATEST definition of this constraint allows - later migrations drop and re-add it.
        private static IReadOnlySet<string> RolesTheDatabaseAllows(string constraint)
        {
            MatchCollection definitions = Regex.Matches(
                MaintainerRoleMigrationTests.MigrationsText(),
                constraint + @"\s+CHECK\s*\(\s*role\s+IN\s*\((?<list>[^)]*)\)",
                RegexOptions.IgnoreCase);

            Assert.NotEmpty(definitions);

            return Regex.Matches(definitions[^1].Groups["list"].Value, "'(?<role>[a-z_]+)'")
                .Select(match => match.Groups["role"].Value)
                .ToHashSet(StringComparer.Ordinal);
        }

        // The words the store writes, from its private RoleText.
        private static string RoleText(ApproverRole role)
        {
            MethodInfo? method = typeof(MySqlSubmissionStore).GetMethod("RoleText", BindingFlags.NonPublic | BindingFlags.Static);
            Assert.NotNull(method);
            return (string)method!.Invoke(null, [role])!;
        }

        [Theory]
        [InlineData("ck_submission_approvals_role")]
        [InlineData("ck_production_approvals_role")]
        public void The_store_writes_exactly_the_role_words_the_database_allows(string constraint)
        {
            IReadOnlySet<string> allowed = MaintainerRoleMigrationTests.RolesTheDatabaseAllows(constraint);

            Assert.Equal(
                allowed.Order(StringComparer.Ordinal),
                Enum.GetValues<ApproverRole>().Select(MaintainerRoleMigrationTests.RoleText).Order(StringComparer.Ordinal));
        }

        [Fact]
        public void The_role_is_stored_as_maintainer_not_reviewer()
        {
            Assert.Equal("maintainer", MaintainerRoleMigrationTests.RoleText(ApproverRole.Maintainer));
            Assert.Equal("administrator", MaintainerRoleMigrationTests.RoleText(ApproverRole.Administrator));
        }

        // 0010 alone, comments stripped, as one statement per entry.
        private static IReadOnlyList<string> RenameMigrationStatements()
        {
            string path = Directory.GetFiles(Path.Combine(AppContext.BaseDirectory, "Migrations"), "0010_*.sql").Single();
            string sql = Regex.Replace(File.ReadAllText(path), "--[^\n]*", string.Empty);

            return sql.Split(';')
                .Select(statement => Regex.Replace(statement, @"\s+", " ").Trim())
                .Where(statement => statement.Length > 0)
                .ToList();
        }

        // ###########################################################################################
        // *** THE AUDIT TRAIL KEEPS ONE NAME PER EVENT (code review, 2026-09-25). *** The grant and
        // revoke actions became maintainer.granted / maintainer.revoked, and the BETA database already
        // holds reviewer.* rows from grants made before the rename. Left alone, a query for every
        // grant of a system misses all of them. 0010 must rewrite each old word to the one the code
        // now writes.
        // ###########################################################################################
        [Theory]
        [InlineData("reviewer.granted", MaintainerAssignmentFlows.GrantedAction)]
        [InlineData("reviewer.revoked", MaintainerAssignmentFlows.RevokedAction)]
        public void The_rename_rewrites_each_old_audit_action_to_the_one_the_code_writes(string old, string current)
        {
            Assert.Contains(
                MaintainerRoleMigrationTests.RenameMigrationStatements(),
                statement => Regex.IsMatch(
                    statement,
                    $@"^UPDATE audit SET action = '{Regex.Escape(current)}' WHERE action = '{Regex.Escape(old)}'$",
                    RegexOptions.IgnoreCase));
        }

        // ###########################################################################################
        // *** A FAILED RUN CAN BE RUN AGAIN (code review, 2026-09-25). *** MariaDB commits each DDL
        // statement on its own, so a failure part-way leaves what ran done and the migration
        // unrecorded, and the next start runs it again from the top. Everything before the rename
        // must therefore be repeatable - each constraint dropped IF EXISTS - and the one statement
        // that cannot be repeated, the RENAME, must come LAST: a rename done first and a failure after
        // it left "Table 'reviewers' doesn't exist" on every later start.
        // ###########################################################################################
        [Fact]
        public void The_rename_can_run_again_after_failing_part_way()
        {
            IReadOnlyList<string> statements = MaintainerRoleMigrationTests.RenameMigrationStatements();

            Assert.Matches(@"^RENAME TABLE reviewers TO maintainers$", statements[^1]);
            Assert.Single(statements, statement => statement.StartsWith("RENAME TABLE", StringComparison.OrdinalIgnoreCase));

            Assert.All(
                statements.Where(statement => statement.Contains("DROP CONSTRAINT", StringComparison.OrdinalIgnoreCase)),
                statement => Assert.Contains("DROP CONSTRAINT IF EXISTS", statement, StringComparison.OrdinalIgnoreCase));
        }

        // The pool table was created as `maintainers`, renamed to `reviewers` by 0006 and back by
        // 0010. Replaying every CREATE and RENAME must end on the name the stores' SQL uses.
        [Fact]
        public void After_every_migration_the_pool_table_is_called_maintainers()
        {
            string sql = MaintainerRoleMigrationTests.MigrationsText();

            Assert.Matches(@"CREATE TABLE maintainers\b", sql);

            string name = "maintainers";

            foreach (Match rename in Regex.Matches(sql, @"RENAME TABLE\s+(?<from>\w+)\s+TO\s+(?<to>\w+)", RegexOptions.IgnoreCase))
            {
                if (string.Equals(rename.Groups["from"].Value, name, StringComparison.OrdinalIgnoreCase))
                    name = rename.Groups["to"].Value;
            }

            Assert.Equal("maintainers", name);
        }
    }
}
