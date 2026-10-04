using System;
using Handlers.DataHandling;

namespace Handlers.MaintainerHandling
{
    // ###########################################################################################
    // WHICH EMAIL ADDRESS CRT USES WHEREVER IT ASKS FOR ONE (owner request, 2026-10-01: "When I am a
    // maintainer, and I have logged in, then I want to use that email address everywhere in the CRT
    // app - e.g. for the Feedback tab or in the Draft tab").
    //
    // Signed in on the Maintainer tab, it is the ACCOUNT's address, and the box showing it is read
    // only; otherwise the one the user typed last (UserSettings.ContactEmail), as before. The
    // Feedback tab and the Submit dialog both ask here, so the two cannot pick differently.
    //
    // *** A SUBMISSION SENT SIGNED IN GOES WITH THE ACCOUNT, NOT ONLY ITS ADDRESS. *** The Submit
    // dialog sends the session's token with it (SubmissionClient.CreateAsync), so the maintainer
    // reviewing it reads "Sent with an account - the address is verified" rather than "Sent without
    // an account - the address was typed in". That line, on the owner's own submission, is what
    // this was asked for.
    //
    // *** THE ACCOUNT'S ADDRESS IS NEVER SAVED AS THE TYPED ONE. *** Signing out brings back what
    // the user had typed, so the account's address does not stay behind on a machine somebody
    // else uses. A session no longer usable (ReviewSession.IsUsableAt) counts as signed out.
    // ###########################################################################################
    public sealed record ContactAddress(string Email, ReviewSession? Account)
    {
        public bool IsFromAccount => this.Account is not null;

        public static ContactAddress Choose(ReviewSession? signedIn, string? typedEmail, DateTimeOffset now)
        {
            if (signedIn is not null &&
                signedIn.IsUsableAt(now) &&
                EmailAddressRules.IsPlausible(signedIn.Email))
            {
                return new ContactAddress(signedIn.Email.Trim(), signedIn);
            }

            return new ContactAddress(typedEmail?.Trim() ?? string.Empty, null);
        }

        // Under the Feedback tab's address while it is the account's.
        public const string FeedbackNote =
            "Your maintainer account's address, as you are signed in on the Maintainer tab. Sign out there to use another one.";

        // Under the Submit dialog's address while it is the account's - in place of the line saying
        // there is no account to create, which would then be untrue.
        public const string SubmitNote =
            "Sent with your maintainer account, as you are signed in on the Maintainer tab, so the maintainer reviewing it " +
            "sees a verified address. The outcome is shown in \"My submissions\" on the Drafts tab. Sign out on the Maintainer " +
            "tab to send with another address.";
    }
}
