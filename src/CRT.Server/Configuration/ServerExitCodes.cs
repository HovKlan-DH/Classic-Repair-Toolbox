using System.Data.Common;
using CRT.Server.Handlers.Database;

namespace CRT.Server.Configuration
{
    // ###########################################################################################
    // The exit codes the service stops with when it cannot start (2026-09-25, migrations
    // 2026-09-27).
    //
    // A refused start used to THROW out of Main. On Linux an unhandled exception aborts the
    // process, so systemd-coredump wrote a core dump of every thread into the journal, and
    // Restart=always started the service again five seconds later - another dump every five
    // seconds, burying the one line that named the wrong setting. That is what the first deploy
    // after the security review looked like, and what the first deploy of migration 0013 looked
    // like two days later, when a table made by hand was in its way.
    //
    // So the service now logs the reason at crit and EXITS:
    //
    //   ConfigurationRefused (78, EX_CONFIG) - a setting was refused. A restart cannot fix it, and
    //     the unit's RestartPreventExitStatus= names the code so systemd stops there.
    //   MigrationFailed - the SAME code, deliberately: a migration that failed, or one MigrationPlan
    //     refused, needs a person to look at the database, and restarting only repeats it. Sharing
    //     78 means the shipped unit (crt-server.service) needed no change to stop retrying it.
    //   DatabaseUnreachable (69, EX_UNAVAILABLE) - the database did not answer. That IS what a
    //     restart fixes (MariaDB coming up a few seconds after this service at boot), so the unit
    //     does not name it and systemd tries again - now without a core dump each time.
    //
    // The shipped unit file, crt-server.service, names 78 and not 69; ServerExitCodesTests fails if they part.
    // ###########################################################################################
    public static class ServerExitCodes
    {
        public const int ConfigurationRefused = 78;

        public const int MigrationFailed = ServerExitCodes.ConfigurationRefused;

        public const int DatabaseUnreachable = 69;

        // The code a start that failed with `exception` exits with - null for anything else,
        // which Main lets through as before.
        public static int? ForStartupFailure(Exception exception) =>
            exception switch
            {
                MigrationFailedException => ServerExitCodes.MigrationFailed,
                DbException => ServerExitCodes.DatabaseUnreachable,
                _ => null
            };
    }
}
