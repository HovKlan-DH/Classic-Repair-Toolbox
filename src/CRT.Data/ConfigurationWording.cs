namespace Handlers.DataHandling
{
    // ###########################################################################################
    // CRT's Configuration tab check boxes that OTHER texts tell people to tick (2026-10-03) - the
    // "try it in BETA" notice under CRT's tabs, and the server's "accepted into BETA" mail. Written
    // once here and read by all three: the Configuration tab's own markup (x:Static), CRT's notice
    // and CRT.Server's EmailTemplates. A label renamed in one place only would send a contributor
    // looking for a check box that is not there.
    // ###########################################################################################
    public static class ConfigurationWording
    {
        public const string BetaSourceCheckBox = "Download data from the BETA source instead of the stable source";

        public const string CheckDataOnLaunchCheckBox = "Check for new or updated data at application launch";
    }
}
