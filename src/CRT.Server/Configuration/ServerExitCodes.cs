namespace CRT.Server.Configuration
{
    // ###########################################################################################
    // The exit code the service stops with when its settings are refused (2026-09-25).
    //
    // A refused start used to THROW out of Main. On Linux an unhandled exception aborts the
    // process, so systemd-coredump wrote a core dump of every thread into the journal, and
    // Restart=always started the service again five seconds later - another dump every five
    // seconds, burying the one line that named the wrong setting. That is what the first deploy
    // after the security review looked like.
    //
    // A wrong setting is not something a restart can fix, so the service now logs the reasons and
    // EXITS with this code, and the unit's RestartPreventExitStatus= names it so systemd stops
    // there. 78 is EX_CONFIG from sysexits.h, "configuration error".
    //
    // Everything else that stops a start - the database unreachable, a migration failing - still
    // throws and is still retried, on purpose: MariaDB coming up a few seconds after this service
    // at boot is exactly the case a restart does fix.
    //
    // DEPLOYMENT.md's unit file names the same number; ServerExitCodesTests fails if they part.
    // ###########################################################################################
    public static class ServerExitCodes
    {
        public const int ConfigurationRefused = 78;
    }
}
