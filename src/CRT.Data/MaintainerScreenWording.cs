namespace Handlers.DataHandling
{
    // ###########################################################################################
    // THE MAINTAINER TAB'S SCREENS, AS ITS TAB STRIP NAMES THEM (owner request, 2026-10-04: rename
    // "Contributor Submissions" to "Queue: Contributor submissions", "BETA > Stable" to "Queue:
    // Awaiting push from BETA to stable" and "Admin" to "Administrator activities" - and later the
    // same day "Administrator activities" to "Account", made every maintainer's, with "My account"
    // as its first entry).
    //
    // Written once here and read by everything that names a screen: the tab strip itself (x:Static
    // in TabMaintainer.axaml), the Maintainer tab's own messages, CRT.Data's OneSubmissionInBeta and
    // CRT.Server's refusals and mails, which CRT shows as they come. A screen renamed in one place
    // only would send a maintainer looking for a tab that is not there - the reason
    // ConfigurationWording exists too.
    //
    // A sentence names a screen QUOTED (the *Quoted constants): bare, "It waits under Queue:
    // Contributor submissions, where ..." reads as two sentences run together. Constants, so a
    // `const` message can be built from them.
    // ###########################################################################################
    public static class MaintainerScreenWording
    {
        public const string Boards = "Boards";

        public const string ContributorQueue = "Queue: Contributor submissions";

        public const string BetaQueue = "Queue: Awaiting push from BETA to stable";

        // Every maintainer's: "My account" and "Server version", and below them the administrator's
        // own entries, each marked with a padlock.
        public const string Account = "Account";

        // The first entry under "Account": the name, the email address, the password and signing out.
        public const string MyAccount = "My account";

        public const string ContributorQueueQuoted = "\"" + ContributorQueue + "\"";

        public const string BetaQueueQuoted = "\"" + BetaQueue + "\"";

        public const string AccountQuoted = "\"" + Account + "\"";

        public const string MyAccountQuoted = "\"" + MyAccount + "\"";
    }
}
