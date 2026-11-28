namespace Handlers.DataHandling
{
    // ###########################################################################################
    // THE WORDS ABOUT AN ACCOUNT THAT MORE THAN ONE PLACE SAYS - CRT's account window, the account web pages the server
    // serves, and the mails. A button a mail or a page tells somebody to press is named here, and the window's button reads
    // the very constant, so the two cannot drift apart.
    //
    // An account is a MAINTAINER'S (owner decision, 2026-11-28: "Entirely remove the new "Account" possibility ... I do not want
    // the possibility for everyone to create an account. This should of course continue to work as-is for the "Maintainer"
    // tab") - made by an invitation, signed in to through the "Maintainer" tab's "Sign in", which opens the account window.
    // (From 2026-11-25 to 2026-11-28 anybody made one, through an account icon in the top right corner of CRT's window.)
    // ###########################################################################################
    public static class AccountWording
    {
        // Where CRT keeps it, as a sentence names it.
        public const string WhereInCrt = "the \"Maintainer\" tab's \"Sign in\"";

        // The buttons, as the window labels them and the mails and pages quote them.
        public const string SignInButton = "Sign in";
        public const string ForgotPasswordButton = "I forgot my password";
        public const string HaveInvitationButton = "I have an invitation";

        // ###########################################################################################
        // How long a mailed code lasts - the server's lifetime (AccountFlows.PasswordResetTokenLifetime is made from it), so the
        // window, the pages and the mails cannot promise another (code review, 2026-11-25: "2 hours" was typed into sentences
        // by hand).
        // ###########################################################################################
        public const int PasswordResetCodeHours = 2;

        // Said where a code for a new password is asked for - CRT's account window and the web page.
        public static readonly string ResetCodeWorks =
            $"It works for {AccountWording.PasswordResetCodeHours} hours and can be used once.";

        public static readonly string ResetCodeHint = "It is mailed to the address, if it has an account. " + AccountWording.ResetCodeWorks;

        // The password rule as a person reads it beside the box - the server's own (AccountRules.MinimumPasswordLength).
        public const int MinimumPasswordLength = 12;

        public const string PasswordLine =
            "At least 12 characters. A few ordinary words together make a strong password that is easy to remember.";
    }
}
