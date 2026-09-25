namespace CRT.Server.Handlers.Email
{
    // ###########################################################################################
    // Builds the mails this service sends. Pure: strings in, EmailMessage out, nothing sent.
    //
    // They fall into two groups with different readers. The ACCOUNT mails below (verification,
    // already-registered, password reset and changed) go to somebody who signed up. The three
    // REVIEW-OUTCOME mails further down go to a CONTRIBUTOR, who has no account at all - see
    // their own header, because that difference changes what the mail may assume and what it can
    // usefully link to.
    //
    // *** THE ANTI-ENUMERATION DESIGN LIVES HERE, NOT IN THE ENDPOINTS. *** Registration and
    // password reset both answer a neutral 202 whatever the address turns out to be, so the HTTP
    // response reveals nothing about who has an account. What differs is WHICH MAIL GETS SENT,
    // and that difference is only ever visible to whoever controls the mailbox - the one party
    // already entitled to know.
    //
    //   Register with an unknown address   -> Verification
    //   Register with a KNOWN address      -> AlreadyRegistered (carrying a reset link)
    //   Reset an unknown address           -> NO MAIL AT ALL
    //   Reset a known address              -> PasswordReset
    //
    // The AlreadyRegistered mail is the subtle one and it is doing real work. Someone re-
    // registering is usually a person who forgot they had an account, so the useful answer is a
    // reset link. It also means a stranger probing addresses learns nothing, while the genuine
    // owner is told that somebody tried - which is a security notification they would want.
    //
    // WHY NO MAIL for an unknown reset address: there is nothing useful to say, and mailing a
    // stranger "you have no account here" turns this endpoint into a way to send unsolicited mail
    // to any address at all.
    //
    // KNOW WHO IS READING. CRT's users are hobbyists fixing their own machines. No corporate
    // register, no "your account has been provisioned", no threats about policy. Short, plain,
    // and it says what to do.
    // ###########################################################################################
    public static class EmailTemplates
    {
        // The service's own name as it appears in subject lines.
        private const string ProductName = "Classic Repair Toolbox";

        // ###########################################################################################
        // Sent when a NEW address registers. Carries the link that proves the address is theirs.
        // ###########################################################################################
        public static EmailMessage Verification(
            string toAddress,
            string displayName,
            string verificationUrl,
            int validHours)
        {
            ArgumentException.ThrowIfNullOrWhiteSpace(toAddress);
            ArgumentException.ThrowIfNullOrWhiteSpace(verificationUrl);

            string greeting = EmailTemplates.Greeting(displayName);

            string body =
                $"""
                {greeting}

                Thanks for signing up to {EmailTemplates.ProductName}. Open this link to confirm
                your address:

                {verificationUrl}

                The link works for {validHours} hours. If it expires, you can ask for a new one
                from the app.

                If you did not sign up, you can ignore this message - the address will not be used
                for anything else.
                """;

            return new EmailMessage(
                toAddress,
                $"Confirm your {EmailTemplates.ProductName} address",
                EmailTemplates.Normalise(body));
        }

        // ###########################################################################################
        // Sent when registration is attempted with an address that ALREADY has an account.
        //
        // This is what lets registration answer the same 202 for both cases. It carries a reset
        // link because the overwhelmingly likely explanation is someone who forgot they had an
        // account - and it says plainly that somebody tried, which is what the real owner needs to
        // know if it was not them.
        //
        // It deliberately does NOT say "your password is X" or reveal when the account was made.
        // ###########################################################################################
        // Carries a CODE rather than a link, for the same reason PasswordReset below does - the
        // link it used to print pointed at a path the server maps nothing at.
        public static EmailMessage AlreadyRegistered(
            string toAddress,
            string displayName,
            string resetCode,
            int validHours)
        {
            ArgumentException.ThrowIfNullOrWhiteSpace(toAddress);
            ArgumentException.ThrowIfNullOrWhiteSpace(resetCode);

            string greeting = EmailTemplates.Greeting(displayName);

            string body =
                $"""
                {greeting}

                Someone just tried to create a {EmailTemplates.ProductName} account with this
                address, but you already have one.

                If that was you and you have forgotten your password, here is a reset code:

                {resetCode}

                Open CRT Maintainer, choose "I forgot my password", then paste this code
                and type the password you want.

                The code works for {validHours} hours.

                If it was not you, there is nothing to do. Your account has not changed, and
                whoever tried was not told whether this address is registered.
                """;

            return new EmailMessage(
                toAddress,
                $"You already have a {EmailTemplates.ProductName} account",
                EmailTemplates.Normalise(body));
        }

        // ###########################################################################################
        // Sent when a password reset is requested for an address that DOES have an account.
        //
        // *** IT CARRIES A CODE TO PASTE, NOT A LINK TO CLICK, AND THAT IS A FIX RATHER THAN A
        // PREFERENCE. *** This mail used to print a URL ending "/api/accounts/reset?token=...".
        // Nothing was ever mapped at that path - only POST /reset-password exists, which a browser
        // click cannot reach - so every reset mail ever sent led to a 404. Reported from live use
        // on 2026-09-22, having been unreachable since Phase 3.
        //
        // The fix could have been an HTML form served by the API. It is a pasted code instead
        // (owner's choice) because the audience is a handful of maintainers who are already
        // sitting in front of the maintainer app, and serving HTML would mean the API growing a page
        // to style, escape and keep accessible for one form.
        //
        // *** A PREFETCHING MAIL CLIENT CANNOT BURN THIS. *** The old link was a GET, so a scanner
        // or preview pane following it would redeem a single-use token before the person read the
        // mail. A code that must be pasted into an application cannot be consumed by anything that
        // merely fetches URLs - so this is also the safer shape, not only the simpler one.
        // ###########################################################################################
        public static EmailMessage PasswordReset(
            string toAddress,
            string displayName,
            string resetCode,
            int validHours)
        {
            ArgumentException.ThrowIfNullOrWhiteSpace(toAddress);
            ArgumentException.ThrowIfNullOrWhiteSpace(resetCode);

            string greeting = EmailTemplates.Greeting(displayName);

            string body =
                $"""
                {greeting}

                Here is your {EmailTemplates.ProductName} password reset code:

                {resetCode}

                Open CRT Maintainer, choose "I forgot my password", then paste this code
                and type the password you want.

                The code works for {validHours} hours and can only be used once.

                If you did not ask for this, you can ignore it - your password has not changed.
                """;

            return new EmailMessage(
                toAddress,
                $"Set a new {EmailTemplates.ProductName} password",
                EmailTemplates.Normalise(body));
        }

        // ###########################################################################################
        // Sent after a password is actually changed. Not a courtesy: it is the one notification
        // that lets someone whose account has been taken over find out, so it goes to the address
        // on file and names no link to click - a mail that asks for no action cannot be phished.
        // ###########################################################################################
        public static EmailMessage PasswordChanged(string toAddress, string displayName)
        {
            ArgumentException.ThrowIfNullOrWhiteSpace(toAddress);

            string greeting = EmailTemplates.Greeting(displayName);

            string body =
                $"""
                {greeting}

                Your {EmailTemplates.ProductName} password was just changed.

                If that was you, nothing further is needed.

                If it was not, your account may have been taken over. Use "forgot password" in the
                app to take it back, and any other sessions will be signed out.
                """;

            return new EmailMessage(
                toAddress,
                $"Your {EmailTemplates.ProductName} password was changed",
                EmailTemplates.Normalise(body));
        }

        // ###########################################################################################
        // THE THREE REVIEW-OUTCOME MAILS (owner request, 2026-09-23).
        //
        // *** THE AUDIENCE HERE IS NOT AN ACCOUNT HOLDER. *** Every mail above goes to somebody
        // who signed up. These go to a CONTRIBUTOR, who by design has no account, no password and
        // no inbox on this service - they typed an address into a dialog once and may have
        // forgotten they did. So each one names the board it is about, because a bare "your
        // submission was accepted" is meaningless to someone who sent one three weeks ago.
        //
        // *** NO LINKS, AND NOTHING TO CLICK. *** There is no account to sign in to, so a link
        // would have nowhere to go. The mail carries the outcome and the maintainer's own words;
        // "My submissions" in the app remains the place to see it in context, and is named rather
        // than linked.
        //
        // *** THE MAINTAINER'S COMMENT IS QUOTED VERBATIM AND IS THE POINT OF THE MAIL. *** For a
        // rejection or a change request it is the entire explanation, and the server already
        // refuses either without one (ReviewDecisionRules.IsUsableReason). An approval usually
        // carries none, which is why Published takes it as optional.
        //
        // KNOW WHO IS READING - the same rule the header states. A hobbyist fixing a C64 who sent
        // a correction. Not "your contribution has been processed"; they fixed something and want
        // to know if it is in.
        // ###########################################################################################
        public static EmailMessage SubmissionPublishedToBeta(
            string toAddress,
            string systemName,
            string? maintainerComment,
            bool amendedByMaintainer = false)
        {
            ArgumentException.ThrowIfNullOrWhiteSpace(toAddress);

            string comment = EmailTemplates.QuotedComment(maintainerComment);

            // A maintainer corrected some rows in the maintainer application before publishing
            // (2026-09-25). Said, so a contributor comparing the result with what they sent is not
            // left wondering where the difference came from.
            string amended = amendedByMaintainer
                ? "\n\nA maintainer changed some of the details before publishing them, so what is published is not\n" +
                  $"exactly what you sent - have a look in {EmailTemplates.ProductName} once your data has updated."
                : string.Empty;

            // *** "BETA source", NOT JUST "published" (owner wording, 2026-09-25). *** Since
            // the two-stage publish, an approval writes the BETA data; everyone else gets it once a
            // maintainer has checked it there and published it to the source. Saying only
            // "published" told somebody whose own copy of the data would not change for days that
            // it had. "BETA source" and "source" are the words CRT's Configuration tab uses.
            string body =
                $"""
                Hello,

                Your contribution to {EmailTemplates.DescribeSystem(systemName)} has been accepted
                and published to the BETA source. Once it has had a final check there, it is
                published to the source that everyone downloads from. You will get one more email
                when it is.

                Thank you - this data only exists because people send corrections in.{amended}{comment}

                You can see this in "My submissions" on the Drafts tab in {EmailTemplates.ProductName}.
                """;

            return new EmailMessage(
                toAddress,
                $"Your {EmailTemplates.ProductName} contribution was published to the BETA source",
                EmailTemplates.Normalise(body));
        }

        // ###########################################################################################
        // Sent when the contribution's board is published to PRODUCTION - the second of the two
        // stages, and the one that reaches everyone (2026-09-25).
        // ###########################################################################################
        public static EmailMessage SubmissionPublishedToSource(
            string toAddress,
            string systemName)
        {
            ArgumentException.ThrowIfNullOrWhiteSpace(toAddress);

            string body =
                $"""
                Hello,

                Your contribution to {EmailTemplates.DescribeSystem(systemName)} has been published
                to the source. It reaches everyone the next time their copy of the data updates, and
                that includes your own.

                Thank you again for sending it in.
                """;

            return new EmailMessage(
                toAddress,
                $"Your {EmailTemplates.ProductName} contribution was published to the source",
                EmailTemplates.Normalise(body));
        }

        // ###########################################################################################
        // Sent to the OTHER HALF of a two-person approval (2026-09-25): a change to a shared file
        // needs a maintainer of the board AND the administrator, one of them has approved, and it
        // waits for the other. `what` names the item - a submission, or publishing a board to
        // production.
        // ###########################################################################################
        public static EmailMessage ApprovalNeeded(
            string toAddress,
            string systemName,
            string what,
            string approvedBy)
        {
            ArgumentException.ThrowIfNullOrWhiteSpace(toAddress);

            string body =
                $"""
                Hello,

                {approvedBy} has approved {what} for {EmailTemplates.DescribeSystem(systemName)}.
                It changes shared files, so it needs your approval too before it is published.

                Open CRT Maintainer to look at it.
                """;

            return new EmailMessage(
                toAddress,
                $"{EmailTemplates.ProductName}: your approval is needed for {systemName}",
                EmailTemplates.Normalise(body));
        }

        // ###########################################################################################
        // Sent to the ADMINISTRATORS when a maintainer publishes a system to production (2026-09-25).
        //
        // The stand-in for Phase 6's administrator feed until that exists: publishing to production
        // is what every user downloads, and with no second factor on a maintainer's account, an
        // unexpected one must be noticed. Not sent when an administrator did it themselves.
        // ###########################################################################################
        public static EmailMessage PublishedToProduction(
            string toAddress,
            string systemName,
            string actor,
            string? revision,
            int filesCopied)
        {
            ArgumentException.ThrowIfNullOrWhiteSpace(toAddress);

            string revisionText = string.IsNullOrWhiteSpace(revision) ? "its current BETA state" : $"revision {revision}";

            string body =
                $"""
                Hello,

                {actor} has published {EmailTemplates.DescribeSystem(systemName)} to production, at
                {revisionText}. {filesCopied} file(s) were copied from BETA.

                If you did not expect this, look at the board in the production data and, if it is
                wrong, publish a correction. The audit trail on the server records the details.
                """;

            return new EmailMessage(
                toAddress,
                $"{EmailTemplates.ProductName}: {systemName} was published to production",
                EmailTemplates.Normalise(body));
        }

        // ###########################################################################################
        // Sent when a maintainer asks for something to be changed before it can go in.
        //
        // *** THIS IS THE ONE THAT MOST NEEDS TO ARRIVE. *** It is the only outcome where the
        // contributor has to DO something, and until this mail existed the request sat behind a
        // button in the app that they had no particular reason to press.
        //
        // It says the draft is still theirs, because the honest worry on reading "changes
        // requested" is that the work has been thrown away.
        // ###########################################################################################
        public static EmailMessage SubmissionChangesRequested(
            string toAddress,
            string systemName,
            string maintainerComment)
        {
            ArgumentException.ThrowIfNullOrWhiteSpace(toAddress);

            string body =
                $"""
                Hello,

                Someone has looked at your contribution to {EmailTemplates.DescribeSystem(systemName)}
                and has asked for a change before it can go in:

                {EmailTemplates.Indent(maintainerComment)}

                Your draft is still on your own computer, exactly as you left it. Open the
                Drafts tab in {EmailTemplates.ProductName}, make the change, and send it again.
                """;

            return new EmailMessage(
                toAddress,
                $"A change was requested on your {EmailTemplates.ProductName} contribution",
                EmailTemplates.Normalise(body));
        }

        // ###########################################################################################
        // Sent when a submission will not be going in.
        //
        // *** IT SAYS WHY, AND IT SAYS THE WORK IS NOT LOST. *** A rejection with no reason is
        // indistinguishable from being ignored, which is why the server will not record one
        // without a comment. The draft survives locally, and saying so is what stops this reading
        // as "your afternoon was wasted".
        //
        // No "please try again" flourish. If the maintainer wanted a change they would have asked
        // for one; inviting a resubmission of something just declined wastes everybody's time.
        // ###########################################################################################
        public static EmailMessage SubmissionRejected(
            string toAddress,
            string systemName,
            string maintainerComment)
        {
            ArgumentException.ThrowIfNullOrWhiteSpace(toAddress);

            string body =
                $"""
                Hello,

                Your contribution to {EmailTemplates.DescribeSystem(systemName)} will not be going
                in. The reason given was:

                {EmailTemplates.Indent(maintainerComment)}

                Your draft is still on your own computer and nothing has been deleted. You can keep
                using it locally, change it, or discard it from the Drafts tab in
                {EmailTemplates.ProductName}.
                """;

            return new EmailMessage(
                toAddress,
                $"Your {EmailTemplates.ProductName} contribution was not accepted",
                EmailTemplates.Normalise(body));
        }

        // ###########################################################################################
        // Sent to each MAINTAINER of a system when a submission to it is queued (Phase 6 task 11,
        // 2026-09-25) - or to the administrators, when the system has no maintainers or the
        // submission changes shared files.
        //
        // This one goes to an account holder, unlike the three above: somebody who agreed to look
        // after a board and would otherwise have to open the maintainer application on the off-chance.
        // It names the board and quotes the contributor's own summary, which is what tells a
        // maintainer whether it is a two-minute typo or an evening's work.
        //
        // No link, like every other mail here; the maintainer application is named.
        // ###########################################################################################
        public static EmailMessage SubmissionWaiting(
            string toAddress,
            string systemName,
            long submissionId,
            string? contributorSummary)
        {
            ArgumentException.ThrowIfNullOrWhiteSpace(toAddress);

            string summary = string.IsNullOrWhiteSpace(contributorSummary)
                ? "(no description given)"
                : contributorSummary.Trim();

            string body =
                $"""
                Hello,

                A contribution to {EmailTemplates.DescribeSystem(systemName)} is waiting for review
                (submission #{submissionId}). The contributor described it as:

                {EmailTemplates.Indent(summary)}

                Open CRT Maintainer to look at it. If somebody else reviews it first, it will simply
                be gone from the queue.
                """;

            return new EmailMessage(
                toAddress,
                $"A {EmailTemplates.ProductName} contribution is waiting for review",
                EmailTemplates.Normalise(body));
        }

        // ###########################################################################################
        // "the Commodore C64 250407 board", or a neutral phrase when the name is unusable.
        //
        // The system id arrives as "Commodore/C64/250407/Data C64 250407.xlsx" in some callers and
        // as a display name in others, so this only tidies what it is given rather than parsing -
        // the CALLER decides what the board is called, and a mail is not the place to start
        // guessing at path shapes.
        // ###########################################################################################
        private static string DescribeSystem(string? systemName)
        {
            string trimmed = systemName?.Trim() ?? string.Empty;

            return trimmed.Length == 0 ? "the board data" : trimmed;
        }

        // ###########################################################################################
        // The maintainer's words, indented so they read as a quotation rather than as the service
        // speaking. Every line is indented, not just the first, or a wrapped sentence loses the
        // visual distinction halfway through.
        // ###########################################################################################
        private static string Indent(string? comment)
        {
            string trimmed = comment?.Trim() ?? string.Empty;

            if (trimmed.Length == 0)
            {
                return "    (no reason was given)";
            }

            return string.Join(
                "\n",
                trimmed.Replace("\r\n", "\n").Split('\n').Select(line => "    " + line.TrimEnd()));
        }

        // An optional comment as its own paragraph, or nothing at all. Used by the published mail,
        // where a maintainer usually has nothing to add and an empty "Maintainer said:" heading would
        // be worse than silence.
        private static string QuotedComment(string? comment)
        {
            string trimmed = comment?.Trim() ?? string.Empty;

            if (trimmed.Length == 0)
            {
                return string.Empty;
            }

            return $"\n\nThe maintainer added:\n\n{EmailTemplates.Indent(trimmed)}";
        }

        // ###########################################################################################
        // "Hello Dennis," or a bare "Hello," when there is no usable name.
        //
        // A blank display name is possible for an account created before validation tightened, or
        // for the AlreadyRegistered path where the name supplied by whoever is registering is NOT
        // the name on the existing account - and must never be used, since that would let a
        // stranger put text into a mail sent to somebody else.
        // ###########################################################################################
        private static string Greeting(string? displayName)
        {
            string trimmed = displayName?.Trim() ?? string.Empty;

            return trimmed.Length == 0 ? "Hello," : $"Hello {trimmed},";
        }

        // ###########################################################################################
        // Raw string literals in this file are indented to match the surrounding code, and C#
        // strips that common indentation - but the result still carries the source file's line
        // breaks. Normalising to \r\n is what SMTP expects: RFC 5321 requires CRLF line endings,
        // and a bare \n can cause a message to be rejected or mangled by a strict MTA.
        // ###########################################################################################
        private static string Normalise(string body)
        {
            return body.Replace("\r\n", "\n").Replace("\n", "\r\n");
        }
    }
}
