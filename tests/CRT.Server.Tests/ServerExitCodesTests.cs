using CRT.Server.Configuration;

namespace CRT.Server.Tests
{
    // ###########################################################################################
    // Covers ServerExitCodes - the code a refused configuration exits with.
    //
    // The code only ends the restart loop if the systemd unit names it in
    // RestartPreventExitStatus=. The unit is written by hand from DEPLOYMENT.md, so the runbook is
    // the one place the two halves can be checked against each other: change the number on one
    // side only and this fails.
    // ###########################################################################################
    public sealed class ServerExitCodesTests
    {
        [Fact]
        public void The_runbook_unit_does_not_restart_a_service_whose_settings_were_refused()
        {
            string runbook = File.ReadAllText(ServerExitCodesTests.RunbookPath());

            Assert.Contains(
                "RestartPreventExitStatus=" + ServerExitCodes.ConfigurationRefused,
                runbook,
                StringComparison.Ordinal);
        }

        [Fact]
        public void The_configuration_exit_code_reads_as_a_failure()
        {
            // Zero would be a clean stop: systemd would report the unit as merely inactive and
            // nothing would say a setting was wrong.
            Assert.NotEqual(0, ServerExitCodes.ConfigurationRefused);
        }

        private static string RunbookPath()
        {
            string? folder = AppContext.BaseDirectory;

            while (folder is not null && !File.Exists(Path.Combine(folder, "Classic-Repair-Toolbox.slnx")))
                folder = Path.GetDirectoryName(folder);

            Assert.NotNull(folder);

            return Path.Combine(folder!, "src", "CRT.Server", "DEPLOYMENT.md");
        }
    }
}
