namespace Handlers.DataHandling
{
    // ###########################################################################################
    // THE WORDS ABOUT AN ACCOUNT THAT MORE THAN ONE PLACE SAYS - CRT's account window, the account web pages the server
    // serves, and the mails (owner request, 2026-11-25: accounts for everybody, registered "both places", where "it should be
    // clear that it gives access to both"). A button a mail or a page tells somebody to press is named here, and the window's
    // button reads the very constant, so the two cannot drift apart.
    //
    // Where the account lives in CRT: the person icon in the top right corner of the window (Main.Account.cs), from any tab -
    // the Maintainer tab's own sign-in screen until 2026-11-25, when only maintainers had accounts.
    // ###########################################################################################
    public static class AccountWording
    {
        // What an account is, said first on the window and on every web page that makes one.
        public const string OneAccount =
            "One account for Classic Repair Toolbox, the same in the application and on its website. " +
            "If you maintain a board, it is also your sign-in for the \"Maintainer\" tab.";

        // Where CRT keeps it, as a sentence names it.
        public const string WhereInCrt = "the account icon in the top right corner of CRT's window";

        // The buttons, as the window labels them and the mails and pages quote them.
        public const string SignInButton = "Sign in";
        public const string CreateAccountButton = "Create an account";
        public const string ForgotPasswordButton = "I forgot my password";
        public const string HaveInvitationButton = "I have an invitation";
        public const string SendMailAgainButton = "Send the mail again";

        // ###########################################################################################
        // How long a mailed code and link last - the server's lifetimes (AccountFlows.PasswordResetTokenLifetime and
        // VerificationTokenLifetime are made from these), so the window, the pages and the mails cannot promise another
        // (code review, 2026-11-25: "2 hours" and "24 hours" were typed into three sentences by hand).
        // ###########################################################################################
        public const int PasswordResetCodeHours = 2;
        public const int VerificationLinkHours = 24;

        // Said where a code for a new password is asked for - CRT's account window and the web page.
        public static readonly string ResetCodeWorks =
            $"It works for {AccountWording.PasswordResetCodeHours} hours and can be used once.";

        public static readonly string ResetCodeHint = "It is mailed to the address, if it has an account. " + AccountWording.ResetCodeWorks;

        // The password rule as a person reads it beside the box - the server's own (AccountRules.MinimumPasswordLength).
        public const int MinimumPasswordLength = 12;

        public const string PasswordLine =
            "At least 12 characters. A few ordinary words together make a strong password that is easy to remember.";

        // The address not confirmed yet - said where the signed-in account is shown.
        public const string NotConfirmed =
            "Your address is not confirmed yet. Open the link in the mail you were sent when you made the account.";
    }
}
