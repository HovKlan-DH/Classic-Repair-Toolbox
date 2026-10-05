using CRT.Server.Configuration;
using CRT.Server.Handlers.Database;

namespace CRT.Server.Tests
{
    // ###########################################################################################
    // Covers ServerExitCodes - the codes a start that cannot go on exits with.
    //
    // The code only ends the restart loop if the systemd unit names it in
    // RestartPreventExitStatus=. The unit ships beside the binaries as src/CRT.Server/crt-server.service
    // (2026-10-05 - it was typed in from the installation document before, which a test should not
    // read), so the two halves are checked against each other here: change the number on one side
    // only and this fails.
    // ###########################################################################################
    public sealed class ServerExitCodesTests
    {
        [Fact]
        public void The_shipped_unit_does_not_restart_a_service_whose_settings_were_refused()
        {
            Assert.Contains(ServerExitCodes.ConfigurationRefused, ServerExitCodesTests.CodesTheUnitDoesNotRestart());
        }

        [Fact]
        public void The_configuration_exit_code_reads_as_a_failure()
        {
            // Zero would be a clean stop: systemd would report the unit as merely inactive and
            // nothing would say a setting was wrong.
            Assert.NotEqual(0, ServerExitCodes.ConfigurationRefused);
        }

        // ###########################################################################################
        // A START THAT FAILED IN THE DATABASE (2026-09-27). A failed migration needs a person, so it
        // must NOT be retried - the unit has to name its code. An unreachable database is what a
        // restart fixes (MariaDB up a few seconds after this service at boot), so the unit must NOT
        // name its code. Both halves are read from the shipped unit, as above.
        // ###########################################################################################
        [Fact]
        public void A_failed_migration_stops_the_service_and_is_not_restarted()
        {
            var failure = new MigrationFailedException("Migration 0013 (0013_board_views.sql) failed: Table 'crt_board_views' already exists.");

            Assert.Equal(ServerExitCodes.MigrationFailed, ServerExitCodes.ForStartupFailure(failure));
            Assert.Contains(ServerExitCodes.MigrationFailed, ServerExitCodesTests.CodesTheUnitDoesNotRestart());
        }

        [Fact]
        public void An_unreachable_database_exits_without_a_crash_and_is_restarted()
        {
            Assert.Equal(ServerExitCodes.DatabaseUnreachable, ServerExitCodes.ForStartupFailure(new UnreachableDatabase()));
            Assert.NotEqual(0, ServerExitCodes.DatabaseUnreachable);
            Assert.DoesNotContain(ServerExitCodes.DatabaseUnreachable, ServerExitCodesTests.CodesTheUnitDoesNotRestart());
        }

        // Anything else is not a start failure this class knows, and Main lets it through as before.
        [Fact]
        public void Any_other_failure_is_left_alone()
        {
            Assert.Null(ServerExitCodes.ForStartupFailure(new InvalidOperationException("something else")));
            Assert.Null(ServerExitCodes.ForStartupFailure(new IOException("disk")));
        }

        // The numbers after RestartPreventExitStatus= in the shipped unit file - exactly one such line.
        private static IReadOnlyList<int> CodesTheUnitDoesNotRestart()
        {
            string line = File.ReadAllLines(ServerExitCodesTests.UnitPath())
                .Single(text => text.StartsWith("RestartPreventExitStatus=", StringComparison.Ordinal));

            return line["RestartPreventExitStatus=".Length..]
                .Split(' ', StringSplitOptions.RemoveEmptyEntries)
                .Select(int.Parse)
                .ToList();
        }

        private sealed class UnreachableDatabase : System.Data.Common.DbException
        {
            public UnreachableDatabase()
                : base("Unable to connect to any of the specified MySQL hosts.")
            {
            }
        }

        private static string UnitPath()
        {
            string? folder = AppContext.BaseDirectory;

            while (folder is not null && !File.Exists(Path.Combine(folder, "Classic-Repair-Toolbox.slnx")))
                folder = Path.GetDirectoryName(folder);

            Assert.NotNull(folder);

            return Path.Combine(folder!, "src", "CRT.Server", "crt-server.service");
        }
    }
}
