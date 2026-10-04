using System;
using System.Collections.Generic;

namespace Handlers.MaintainerHandling
{
    // ###########################################################################################
    // What "My account" lets the maintainer send, and how the account is named (owner request,
    // 2026-10-03: the maintainer "can edit his/her own email address and name"; a window of its own,
    // "Your account", until 2026-10-04, when it became the first entry under the Maintainer tab's
    // "Account" screen). Pure, so the view's buttons are tested without a display; the view only
    // asks.
    //
    // A button is on only when pressing it could change something: a name or address the same as
    // the one there now would only be refused by the server. Everything the server checks (a name
    // too short, a password breaking a rule) is left to the server - its sentence is shown - so the
    // rules are never written twice. No current password is asked for anywhere (owner decision,
    // 2026-10-03: "as I see it as you are already logged in").
    // ###########################################################################################
    public static class MyAccountRules
    {
        public const string PasswordsDiffer = "The two new passwords are not the same.";

        // ###########################################################################################
        // Who is logged in, as the Maintainer tab says it under its lists and "My account" at its top
        // (owner request, 2026-10-04: "Logged in as:<br /><b>Dennis</b> (dennis@...dk)"): the name
        // BOLD, then the address in brackets. A blank name gives the address alone, bold - it is
        // then what names the account. Runs, so the view can bold the name; the words "Logged in
        // as" are the caller's, since the tab puts them on a line of their own.
        // ###########################################################################################
        public static IReadOnlyList<ReviewNoteRun> LoggedInAs(string? displayName, string email)
        {
            string name = displayName?.Trim() ?? string.Empty;

            return name.Length == 0
                ? [new ReviewNoteRun(email, IsCount: true)]
                : [new ReviewNoteRun(name, IsCount: true), new ReviewNoteRun($" ({email})", IsCount: false)];
        }

        public static bool CanSaveName(string currentName, string? typed)
        {
            string name = typed?.Trim() ?? string.Empty;

            return name.Length > 0 && !string.Equals(name, currentName, StringComparison.Ordinal);
        }

        // Exactly the address there now is no change; the same address in other capitals IS one (the
        // server makes it at once, with no code).
        public static bool CanAskForCode(string currentEmail, string? typedEmail)
        {
            string address = typedEmail?.Trim() ?? string.Empty;

            return address.Length > 0 && !string.Equals(address, currentEmail, StringComparison.Ordinal);
        }

        public static bool CanConfirmCode(string? code) => !string.IsNullOrWhiteSpace(code);

        // Said only once both new passwords are typed - a half-typed second one is not a mistake yet.
        public static string? PasswordProblem(string? newPassword, string? repeated) =>
            !string.IsNullOrEmpty(newPassword) &&
            !string.IsNullOrEmpty(repeated) &&
            !string.Equals(newPassword, repeated, StringComparison.Ordinal)
                ? MyAccountRules.PasswordsDiffer
                : null;

        public static bool CanChangePassword(string? newPassword, string? repeated) =>
            !string.IsNullOrEmpty(newPassword) &&
            string.Equals(newPassword, repeated, StringComparison.Ordinal);
    }
}
