namespace CRT.Server.Handlers.Email
{
    // ###########################################################################################
    // Builds the mails this service sends. Pure: strings in, EmailMessage out, nothing sent.
    //
    // *** EVERY MAIL IS HTML, WITH A PLAIN-TEXT COPY (owner request, 2026-10-03). *** Each is
    // written once, as MailBody blocks, and MailBody renders both - see its header. No template
    // writes markup, so nothing a person typed can turn into markup in somebody else's mail.
    //
    // They fall into three groups with different readers. The ACCOUNT mails (verification,
    // already-registered, password reset and changed, the three about a changed address, the
    // maintainer invitation) go to somebody who
    // has or is about to have an account. The CONTRIBUTION mails go to a CONTRIBUTOR, who usually
    // has no account at all - see their own header. The REVIEW mails (waiting, approval needed,
    // published to production) go to maintainers and administrators.
    //
    // *** THE CONTRIBUTION WORDING IS THE PROJECT OWNER'S (2026-10-03). *** He wrote the five main
    // mails - waiting for review (an existing system, and a new one), a change requested, accepted
    // into BETA and published to the stable source - and asked for them to be polished but kept
    // informal. The rejected and taken-out-of-BETA mails follow the same shape. Each greets by name
    // when there is one ("Hi Dennis,") and with "Hi there," otherwise, names the system in bold
    // inside brackets, quotes the other person's words in italics, and ends with who to write to.
    //
    // *** THE ANTI-ENUMERATION DESIGN LIVES HERE, NOT IN THE ENDPOINTS. *** Registration and
    // password reset both answer a neutral 202 whatever the address turns out to be, so the HTTP
    // response reveals nothing about who has an account. What differs is WHICH MAIL GETS SENT,
    // and that difference is only ever visible to whoever controls the mailbox - the one party
    // already entitled to know.
    //
    //   Register with an unknown address   -> Verification
    //   Register with a KNOWN address      -> AlreadyRegistered (carrying a reset code)
    //   Reset an unknown address           -> NO MAIL AT ALL
    //   Reset a known address              -> PasswordReset
    //
    // The AlreadyRegistered mail is the subtle one and it is doing real work. Someone re-
    // registering is usually a person who forgot they had an account, so the useful answer is a
    // reset code. It also means a stranger probing addresses learns nothing, while the genuine
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
        // The service's own name, in the account mails' subjects and text.
        private const string ProductName = "Classic Repair Toolbox";

        // ###########################################################################################
        // Who to write to with a question (owner request, 2026-10-03: "Please connect with Dennis,
        // if you have any questions for this"). The closing line of every mail but the one to the
        // administrators about a production publish, which goes to the project owner himself.
        // ###########################################################################################
        public const string ContactName = "Dennis";

        public const string ContactAddress = "dennis@classic-repair-toolbox.dk";

        // The words CRT's Configuration tab uses for the BETA check box - quoted in the BETA mail, so
        // the contributor can find it as written. The tab reads the same constant.
        public const string BetaCheckBoxLabel = global::Handlers.DataHandling.ConfigurationWording.BetaSourceCheckBox;

        // The Maintainer tab's words for the way to "My account" (2026-10-04) - named in the
        // address-change code mail, so the maintainer finds the screen and the entry as written.
        // The tab's strip and list read the same constants.
        private const string AccountScreenQuoted = global::Handlers.DataHandling.MaintainerScreenWording.AccountQuoted;

        private const string MyAccountQuoted = global::Handlers.DataHandling.MaintainerScreenWording.MyAccountQuoted;

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

            return EmailTemplates.Greeted(displayName)
                .Paragraph($"Thanks for signing up to {EmailTemplates.ProductName}. Open this link to confirm your address:")
                .Paragraph(MailText.Link(verificationUrl, verificationUrl))
                .Paragraph($"The link works for {validHours} hours. If it expires, you can ask for a new one from the app.")
                .Paragraph("If you did not sign up, you can ignore this message - the address will not be used for anything else.")
                .WithContactLine()
                .ToMessage(toAddress, $"Confirm your {EmailTemplates.ProductName} address");
        }

        // ###########################################################################################
        // Sent when registration is attempted with an address that ALREADY has an account.
        //
        // This is what lets registration answer the same 202 for both cases. It carries a reset
        // code because the overwhelmingly likely explanation is someone who forgot they had an
        // account - and it says plainly that somebody tried, which is what the real owner needs to
        // know if it was not them.
        //
        // It deliberately does NOT say "your password is X" or reveal when the account was made.
        // Carries a CODE rather than a link, for the same reason PasswordReset below does - the
        // link it used to print pointed at a path the server maps nothing at.
        // ###########################################################################################
        public static EmailMessage AlreadyRegistered(
            string toAddress,
            string displayName,
            string resetCode,
            int validHours)
        {
            ArgumentException.ThrowIfNullOrWhiteSpace(toAddress);
            ArgumentException.ThrowIfNullOrWhiteSpace(resetCode);

            return EmailTemplates.Greeted(displayName)
                .Paragraph($"Someone just tried to create a {EmailTemplates.ProductName} account with this address, but you already have one.")
                .Paragraph("If that was you and you have forgotten your password, here is a reset code:")
                .Code(resetCode)
                .Paragraph("Open CRT, go to the \"Maintainer\" tab, choose \"I forgot my password\", then paste this code and type the password you want.")
                .Paragraph($"The code works for {validHours} hours.")
                .Paragraph("If it was not you, there is nothing to do. Your account has not changed, and whoever tried was not told whether this address is registered.")
                .WithContactLine()
                .ToMessage(toAddress, $"You already have a {EmailTemplates.ProductName} account");
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
        // sitting in front of the Maintainer tab, and serving HTML would mean the API growing a page
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

            return EmailTemplates.Greeted(displayName)
                .Paragraph($"Here is your {EmailTemplates.ProductName} password reset code:")
                .Code(resetCode)
                .Paragraph("Open CRT, go to the \"Maintainer\" tab, choose \"I forgot my password\", then paste this code and type the password you want.")
                .Paragraph($"The code works for {validHours} hours and can only be used once.")
                .Paragraph("If you did not ask for this, you can ignore it - your password has not changed.")
                .WithContactLine()
                .ToMessage(toAddress, $"Set a new {EmailTemplates.ProductName} password");
        }

        // Where CRT is downloaded. A maintainer uses CRT's own Maintainer tab since 2026-09-29; the
        // separate CRT Maintainer application, and the repository its releases lived in, are gone.
        public const string CrtDownloadUrl = "https://github.com/HovKlan-DH/Classic-Repair-Toolbox/releases";

        // ###########################################################################################
        // AN INVITATION TO MAINTAIN A SYSTEM (owner request, 2026-09-27). The person has no account
        // and may not run CRT at all, so it says what a maintainer does, where CRT is, how to show
        // its Maintainer tab, and exactly which button to press - and, like the reset mail, carries a
        // CODE rather than a link: accepting needs a password, so it cannot be a click, and a
        // prefetching mail client cannot spend a code by following it.
        // ###########################################################################################
        public static EmailMessage MaintainerInvitation(
            string toAddress,
            string systemName,
            string invitedBy,
            string invitationCode,
            int validDays)
        {
            ArgumentException.ThrowIfNullOrWhiteSpace(toAddress);
            ArgumentException.ThrowIfNullOrWhiteSpace(invitationCode);

            string inviter = string.IsNullOrWhiteSpace(invitedBy) ? "The administrator" : invitedBy.Trim();

            return EmailTemplates.Greeted(null)
                .Paragraph(
                    $"{inviter} has invited you to be a maintainer of ",
                    MailText.Named(EmailTemplates.DescribeSystem(systemName)),
                    $" in {EmailTemplates.ProductName}. A maintainer looks through the changes people send in for that system, " +
                    "and publishes the good ones for everybody who uses the data.")
                .Paragraph("Your invitation code:")
                .Code(invitationCode)
                .Paragraph("To accept it:")
                .Steps(
                    ["Install Classic Repair Toolbox (CRT) from ", MailText.Link(EmailTemplates.CrtDownloadUrl, EmailTemplates.CrtDownloadUrl), ", or use the one you already have."],
                    ["In CRT's \"Configuration\" tab, tick \"Enable Maintainer tab\"."],
                    ["Open the \"Maintainer\" tab and choose \"I have an invitation\"."],
                    ["Paste the code, and pick the name others will see and a password."])
                .Paragraph($"The code works for {validDays} days and can only be used once.")
                .Paragraph("If you were not expecting this, you can ignore it - nothing happens unless the code is used.")
                .WithContactLine()
                // The subject in the owner's words (2026-10-04) - it read "You are invited to maintain
                // Amstrad / CPC 664 / MC0005A".
                .ToMessage(toAddress, $"For CRT you are invited to maintain the [{systemName}] system");
        }

        // ###########################################################################################
        // Sent after a password is actually changed. Not a courtesy: it is the one notification
        // that lets someone whose account has been taken over find out, so it goes to the address
        // on file and names no link to click - a mail that asks for no action cannot be phished.
        // ###########################################################################################
        public static EmailMessage PasswordChanged(string toAddress, string displayName)
        {
            ArgumentException.ThrowIfNullOrWhiteSpace(toAddress);

            return EmailTemplates.Greeted(displayName)
                .Paragraph($"Your {EmailTemplates.ProductName} password was just changed.")
                .Paragraph("If that was you, nothing further is needed.")
                .Paragraph("If it was not, your account may have been taken over. Use \"I forgot my password\" in the app to take it back, and any other sessions will be signed out.")
                .WithContactLine()
                .ToMessage(toAddress, $"Your {EmailTemplates.ProductName} password was changed");
        }

        // ###########################################################################################
        // A MAINTAINER CHANGING THEIR ADDRESS (owner request, 2026-10-03) - AccountSelfServiceFlows.
        // Three mails, one per mailbox involved:
        //
        //   EmailChangeCode          - to the NEW address: the code that proves it is theirs. A code
        //                              to type in, not a link to click, for the reason PasswordReset
        //                              gives. It names the way in - the "Account" tab, "My
        //                              account", then "I have received a code" - in the
        //                              Maintainer tab's own words (MaintainerScreenWording).
        //   EmailChangeAddressTaken  - to an address that ALREADY has an account, instead of a code.
        //                              The requester is told "a code is on its way" either way, so
        //                              only this mailbox's owner learns the address is registered -
        //                              registration's AlreadyRegistered rule.
        //   EmailChanged             - to the OLD address once the change is made: the one notice
        //                              that reaches the owner if it was not them. It names who to
        //                              write to, because that mailbox can no longer reset the
        //                              password.
        // ###########################################################################################
        public static EmailMessage EmailChangeCode(
            string toAddress,
            string displayName,
            string code,
            int validHours)
        {
            ArgumentException.ThrowIfNullOrWhiteSpace(toAddress);
            ArgumentException.ThrowIfNullOrWhiteSpace(code);

            return EmailTemplates.Greeted(displayName)
                .Paragraph($"You asked to use this address for your {EmailTemplates.ProductName} account. Here is the code that confirms it:")
                .Code(code)
                .Paragraph($"In CRT's \"Maintainer\" tab, choose {EmailTemplates.AccountScreenQuoted} and {EmailTemplates.MyAccountQuoted}, then \"I have received a code\", and type it in. Your address does not change until the code is used.")
                .Paragraph($"The code works for {validHours} hours and can only be used once.")
                .Paragraph("If you did not ask for this, you can ignore it - nothing changes unless the code is used.")
                .WithContactLine()
                .ToMessage(toAddress, $"Confirm your new {EmailTemplates.ProductName} address");
        }

        public static EmailMessage EmailChangeAddressTaken(string toAddress, string displayName)
        {
            ArgumentException.ThrowIfNullOrWhiteSpace(toAddress);

            return EmailTemplates.Greeted(displayName)
                .Paragraph($"Someone asked to move a {EmailTemplates.ProductName} account to this address, but this address already has an account - yours - so nothing was changed.")
                .Paragraph("If that was you, sign in with this address instead, or choose a different one for the other account.")
                .Paragraph("If it was not you, there is nothing to do. Your account has not changed, and whoever asked was not told whether this address is registered.")
                .WithContactLine()
                .ToMessage(toAddress, $"Someone tried to use your {EmailTemplates.ProductName} address");
        }

        public static EmailMessage EmailChanged(string toOldAddress, string displayName, string newAddress)
        {
            ArgumentException.ThrowIfNullOrWhiteSpace(toOldAddress);
            ArgumentException.ThrowIfNullOrWhiteSpace(newAddress);

            return EmailTemplates.Greeted(displayName)
                .Paragraph(
                    $"The email address of your {EmailTemplates.ProductName} account was just changed to ",
                    MailText.Bold(newAddress),
                    ". You sign in with that address from now on, and mail about the account goes there.")
                .Paragraph("If that was you, nothing further is needed.")
                .Paragraph(
                    "If it was not, your account may have been taken over, and this address can no longer be used to reset its password. " +
                    $"Write to {EmailTemplates.ContactName} straight away at ",
                    MailText.Link(EmailTemplates.ContactAddress, $"mailto:{EmailTemplates.ContactAddress}"),
                    ".")
                .ToMessage(toOldAddress, $"Your {EmailTemplates.ProductName} address was changed");
        }

        // ###########################################################################################
        // THE CONTRIBUTION MAILS (owner request, 2026-09-23; the owner's own wording since
        // 2026-10-03).
        //
        // *** THE AUDIENCE IS USUALLY NOT AN ACCOUNT HOLDER. *** A contributor typed an address into
        // a dialog once and may have forgotten they did - so each mail names the system it is about,
        // because a bare "your submission was accepted" means nothing three weeks later. A
        // contributor who sent with an account (a signed-in maintainer) is greeted by name; anybody
        // else as "Hi there,".
        //
        // *** NO LINKS TO THE SERVICE. *** There is no account to sign in to, so a link would have
        // nowhere to go. The mail carries the outcome and the maintainer's own words; CRT's Drafts
        // tab is named rather than linked.
        //
        // *** THE MAINTAINER'S COMMENT IS QUOTED VERBATIM AND IS THE POINT OF THE MAIL. *** For a
        // rejection or a change request it is the entire explanation, and the server already
        // refuses either without one (ReviewDecisionRules.IsUsableReason). An approval usually
        // carries none, which is why the BETA mail takes it as optional.
        // ###########################################################################################

        // ###########################################################################################
        // Accepted, and written to the BETA source - the first of the two stages (2026-09-25).
        //
        // *** IT SAYS HOW TO TRY IT (owner request, 2026-10-03). *** "the BETA source" and "the
        // stable source" are the words CRT's Configuration tab uses, and the check box is named as
        // CRT writes it. It also says that being turned down at this stage is normal: a maintainer
        // often only sees a problem once the data is live in BETA, and the contributor should not
        // read a later "taken back out" as a failure.
        // ###########################################################################################
        public static EmailMessage SubmissionPublishedToBeta(
            string toAddress,
            string systemName,
            string? maintainerComment,
            bool amendedByMaintainer = false,
            string? contributorName = null)
        {
            ArgumentException.ThrowIfNullOrWhiteSpace(toAddress);

            MailBody body = EmailTemplates.Greeted(contributorName)
                .Paragraph(
                    "A system maintainer has reviewed your contribution to ",
                    MailText.Named(EmailTemplates.DescribeSystem(systemName)),
                    ". It has been accepted and is now in the BETA online source, ready for you to test. To try it, go to the " +
                    "\"Configuration\" tab in CRT and tick \"",
                    MailText.Bold(EmailTemplates.BetaCheckBoxLabel),
                    "\".");

            // A maintainer corrected some rows in the Maintainer tab before accepting it (2026-09-25).
            // Said, so a contributor comparing the result with what they sent is not left wondering
            // where the difference came from.
            if (amendedByMaintainer)
            {
                body.Paragraph("The maintainer changed some of the details before accepting it, so what is in BETA is not exactly what you sent.");
            }

            if (!string.IsNullOrWhiteSpace(maintainerComment))
            {
                body.Paragraph("The maintainer added:")
                    .Quote(maintainerComment, string.Empty);
            }

            return body
                .Paragraph(
                    "The BETA source is for testing only, so you will hear from us again: either your contribution is published " +
                    "to the stable source, or it is taken back out of BETA, and then you are told why.")
                .Paragraph(
                    "Don't be put off if that happens - a maintainer often only spots a problem once the data is \"live\" in BETA. " +
                    "It is all part of the process :-)")
                .Paragraph("Your draft stays in the \"Drafts\" tab until your contribution has been published to the stable source.")
                .WithContactLine()
                .ToMessage(toAddress, "Your CRT contribution is accepted and ready to test in the BETA source");
        }

        // ###########################################################################################
        // Published to PRODUCTION - the second of the two stages, and the one that reaches everyone
        // (2026-09-25). It says how to get the data, to untick BETA if it was ticked for testing,
        // and that CRT removes the draft by itself once the data matches (PublishedDraftRetirer).
        // ###########################################################################################
        public static EmailMessage SubmissionPublishedToSource(
            string toAddress,
            string systemName,
            string? contributorName = null)
        {
            ArgumentException.ThrowIfNullOrWhiteSpace(toAddress);

            return EmailTemplates.Greeted(contributorName)
                .Paragraph(
                    "Your contribution to ",
                    MailText.Named(EmailTemplates.DescribeSystem(systemName)),
                    " has been fully accepted and is now published to the stable source, for everyone.")
                .Paragraph(
                    "Restart CRT to get the newest data from the stable source - if you ticked \"",
                    MailText.Bold(EmailTemplates.BetaCheckBoxLabel),
                    "\" to test it, untick it in the \"Configuration\" tab first. Once your data matches the published system, " +
                    "CRT removes your draft of it by itself.")
                .Paragraph(
                    "Everyone who uses ",
                    MailText.Italic(EmailTemplates.ProductName),
                    " is grateful for contributions like yours - they are what keeps this community alive.")
                .WithContactLine()
                .ToMessage(toAddress, "Your CRT contribution has been published to the stable source");
        }

        // ###########################################################################################
        // A maintainer asks for something to be changed before it can go in.
        //
        // *** THIS IS THE ONE THAT MOST NEEDS TO ARRIVE. *** It is the only outcome where the
        // contributor has to DO something. It says the draft is still theirs, because the honest
        // worry on reading "changes requested" is that the work has been thrown away.
        // ###########################################################################################
        public static EmailMessage SubmissionChangesRequested(
            string toAddress,
            string systemName,
            string maintainerComment,
            string? contributorName = null)
        {
            ArgumentException.ThrowIfNullOrWhiteSpace(toAddress);

            return EmailTemplates.Greeted(contributorName)
                .Paragraph(
                    "A system maintainer has reviewed your contribution to ",
                    MailText.Named(EmailTemplates.DescribeSystem(systemName)),
                    " and would like a change before it can go into the online sources:")
                .Quote(maintainerComment, "(no reason was given)")
                .Paragraph(
                    "Your draft is still on your own computer, exactly as you left it (unless you have discarded it). " +
                    "Open CRT, go to the \"Drafts\" tab, make the change, and submit it again.")
                .WithContactLine()
                .ToMessage(toAddress, "A change was requested on your CRT contribution");
        }

        // ###########################################################################################
        // A maintainer rolled a BETA board back and this submission went with it (owner decision,
        // 2026-09-27) - Beta > Prod's "Push back to queue".
        //
        // *** IT SAYS THE DATA WAS IN BETA AND NOW IS NOT. *** This contributor was already told
        // "accepted into BETA"; a mail that read like an ordinary change request would leave them
        // believing their work was still live. And it promises only what is true: the contribution
        // is back in the queue as it was - a maintainer may have amended it, so a newer submission
        // does not necessarily replace it (SubmissionReplacementRules).
        //
        // The comment is always present: BetaRollbackFlow refuses a rollback without one.
        // ###########################################################################################
        public static EmailMessage SubmissionReturnedToQueue(
            string toAddress,
            string systemName,
            string maintainerComment,
            string? contributorName = null)
        {
            ArgumentException.ThrowIfNullOrWhiteSpace(toAddress);

            return EmailTemplates.Greeted(contributorName)
                .Paragraph(
                    "Your contribution to ",
                    MailText.Named(EmailTemplates.DescribeSystem(systemName)),
                    " had been accepted into the BETA source, but a system maintainer has taken it back out and returned it " +
                    "to the review queue:")
                .Quote(maintainerComment, "(no reason was given)")
                .Paragraph(
                    "The BETA source no longer holds it. Your contribution is not lost: it is back in the queue as it was, " +
                    "and a maintainer will look at it again.")
                .Paragraph(
                    "If the note above asks you for something, make the change in the \"Drafts\" tab in CRT and submit it again. " +
                    "If you have discarded your draft, start a new one from the board.")
                .WithContactLine()
                .ToMessage(toAddress, "Your CRT contribution was taken back out of BETA");
        }

        // ###########################################################################################
        // A submission will not be going in - from the review queue, or Beta > Prod's "Reject".
        //
        // *** IT SAYS WHY, AND IT SAYS THE WORK IS NOT LOST. *** A rejection with no reason is
        // indistinguishable from being ignored, which is why the server will not record one
        // without a comment. No "please try again" flourish: if the maintainer wanted a change
        // they would have asked for one.
        // ###########################################################################################
        public static EmailMessage SubmissionRejected(
            string toAddress,
            string systemName,
            string maintainerComment,
            string? contributorName = null)
        {
            ArgumentException.ThrowIfNullOrWhiteSpace(toAddress);

            return EmailTemplates.Greeted(contributorName)
                .Paragraph(
                    "A system maintainer has reviewed your contribution to ",
                    MailText.Named(EmailTemplates.DescribeSystem(systemName)),
                    ", and it will not be going in. The reason given was:")
                .Quote(maintainerComment, "(no reason was given)")
                .Paragraph(
                    "Your draft is still on your own computer (unless you have discarded it) and nothing has been deleted. " +
                    "You can keep using it, change it, or discard it from the \"Drafts\" tab in CRT.")
                .WithContactLine()
                .ToMessage(toAddress, "Your CRT contribution was not accepted");
        }

        // ###########################################################################################
        // The administrator deleted a system while this contributor's submission to it was still in
        // play - waiting for review, approved once, or in BETA (owner decision, 2026-10-03: "Delete
        // them and mail the contributors"). See SystemDeletionFlow.
        //
        // *** IT SAYS THE SUBMISSION IS GONE, AND THE DRAFT IS NOT. *** CRT's "My submissions" keeps
        // showing the submission's last state (the server now answers "not found", which CRT treats
        // as "no news"), so this mail is the only place the contributor learns what happened to it.
        // ###########################################################################################
        public static EmailMessage SystemDeleted(
            string toAddress,
            string systemName,
            string reason,
            string? contributorName = null)
        {
            ArgumentException.ThrowIfNullOrWhiteSpace(toAddress);

            return EmailTemplates.Greeted(contributorName)
                .Paragraph(
                    "The administrator has removed ",
                    MailText.Named(EmailTemplates.DescribeSystem(systemName)),
                    " from CRT completely, and your contribution to it was removed along with it. The reason given was:")
                .Quote(reason, "(no reason was given)")
                .Paragraph(
                    "Your draft is still on your own computer (unless you have discarded it). You can keep it, or discard it " +
                    "from the \"Drafts\" tab in CRT.")
                .WithContactLine()
                .ToMessage(toAddress, "A system you contributed to was removed from CRT");
        }

        // ###########################################################################################
        // Sent to each MAINTAINER of a system when a submission to it is queued (Phase 6 task 11,
        // 2026-09-25) - or to the administrators, when the system has no maintainers or the
        // submission changes shared files.
        //
        // *** AN UPDATE AND A NEW SYSTEM ARE TOLD APART (owner request, 2026-10-03). *** A new
        // system is an evening's review, an update is often a typo - the first line says which, and
        // the contributor's own description is quoted, which is what tells a maintainer how big it
        // is. Greets the maintainer by the name on their account.
        // ###########################################################################################
        public static EmailMessage SubmissionWaiting(
            string toAddress,
            string systemName,
            long submissionId,
            string? contributorSummary,
            bool isNewSystem = false,
            string? recipientName = null)
        {
            ArgumentException.ThrowIfNullOrWhiteSpace(toAddress);

            string what = isNewSystem ? "a completely new system, " : "an update for an existing system, ";
            string described = isNewSystem ? "the new system" : "the changes";

            return EmailTemplates.Greeted(recipientName)
                .Paragraph(
                    $"A contributor has sent in {what}",
                    MailText.Named(EmailTemplates.DescribeSystem(systemName)),
                    $", and it is waiting for review (submission #{submissionId}). The contributor describes {described} like this:")
                .Quote(contributorSummary, "(no description given)")
                .Paragraph(
                    "Open CRT and go to the \"Maintainer\" tab to look at it. If somebody else reviews it first, " +
                    "it simply disappears from your queue.")
                .WithContactLine()
                .ToMessage(toAddress, "A CRT contribution is waiting for your review");
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
            string approvedBy,
            string? recipientName = null)
        {
            ArgumentException.ThrowIfNullOrWhiteSpace(toAddress);

            return EmailTemplates.Greeted(recipientName)
                .Paragraph(
                    $"{approvedBy} has approved {what} for ",
                    MailText.Named(EmailTemplates.DescribeSystem(systemName)),
                    ". It replaces a shared file that other boards may use, so it needs your approval too before it is published.")
                .Paragraph("Open CRT and go to the \"Maintainer\" tab to look at it.")
                .WithContactLine()
                .ToMessage(toAddress, $"CRT: your approval is needed for {systemName}");
        }

        // ###########################################################################################
        // Sent to the ADMINISTRATORS when a maintainer publishes a system to production (2026-09-25).
        //
        // The stand-in for Phase 6's administrator feed until that exists: publishing to production
        // is what every user downloads, and with no second factor on a maintainer's account, an
        // unexpected one must be noticed. Not sent when an administrator did it themselves. No
        // contact line: it goes to the project owner, who is the contact.
        // ###########################################################################################
        public static EmailMessage PublishedToProduction(
            string toAddress,
            string systemName,
            string actor,
            string? revision,
            int filesCopied,
            string? recipientName = null)
        {
            ArgumentException.ThrowIfNullOrWhiteSpace(toAddress);

            string revisionText = string.IsNullOrWhiteSpace(revision) ? "its current BETA state" : $"revision {revision}";

            return EmailTemplates.Greeted(recipientName)
                .Paragraph(
                    $"{actor} has published ",
                    MailText.Named(EmailTemplates.DescribeSystem(systemName)),
                    $" to the stable source, at {revisionText}. {filesCopied} file(s) were copied from BETA.")
                .Paragraph(
                    "If you did not expect this, look at the board in the stable data and, if it is wrong, publish a correction. " +
                    "The audit trail on the server records the details.")
                .ToMessage(toAddress, $"CRT: {systemName} was published to the stable source");
        }

        // ###########################################################################################
        // FEEDBACK from CRT's Feedback tab, to the project owner (owner request, 2026-10-03: "can we
        // retire the old "Feedback" backend PHP ... these mails should then change into HTML
        // format, except that e.g. logfile should still show as monospace as this is kind of quoted
        // text"). The PHP page's mail, section for section: the sender's words as a quotation, CRT's
        // own text files as written in a fixed-width font, and the attached files listed under the
        // folder they were saved in - its "Internal reference", the name to look for on the share.
        //
        // Reply goes to the sender when an address was given (the From is the service's own - see
        // EmailMessage). No greeting and no contact line: it goes to the contact. Wide, for the logs.
        // ###########################################################################################
        public static EmailMessage Feedback(
            string toAddress,
            string version,
            string? senderAddress,
            string feedback,
            string? reference,
            IReadOnlyList<FeedbackShownFile> shownFiles,
            IReadOnlyList<string> savedFiles,
            int savedFilesNotListed,
            IReadOnlyList<string> notes)
        {
            ArgumentException.ThrowIfNullOrWhiteSpace(toAddress);
            ArgumentNullException.ThrowIfNull(shownFiles);
            ArgumentNullException.ThrowIfNull(savedFiles);
            ArgumentNullException.ThrowIfNull(notes);

            var body = new MailBody(MailBody.WideWidth);

            if (string.IsNullOrWhiteSpace(senderAddress))
                body.Paragraph("From: ", MailText.Italic("no email address given"));
            else
                body.Paragraph("From: ", MailText.Link(senderAddress, $"mailto:{senderAddress}"), " - replying to this mail answers there.");

            if (!string.IsNullOrWhiteSpace(reference))
                body.Paragraph("Internal reference: ", MailText.Bold(reference));

            body.Paragraph(MailText.Bold("Feedback from user:"))
                .Quote(feedback, "(no text given)");

            foreach (FeedbackShownFile shown in shownFiles)
            {
                body.Paragraph(MailText.Bold($"{shown.Title}:"))
                    .Preformatted(shown.Text.Length == 0 ? "(empty)" : shown.Text);
            }

            if (savedFiles.Count > 0 || notes.Count > 0)
            {
                body.Paragraph(MailText.Bold("Files associated:"));

                if (savedFiles.Count > 0)
                {
                    string list = string.Join("\n", savedFiles);

                    if (savedFilesNotListed > 0)
                        list += $"\n... and {savedFilesNotListed} more file(s) not listed";

                    body.Preformatted(list);
                }

                foreach (string note in notes)
                    body.Paragraph(note);
            }

            string subjectVersion = string.IsNullOrWhiteSpace(version) ? "Unknown" : version.Trim();

            return body.ToMessage(toAddress, $"Feedback from CRT {subjectVersion}", string.IsNullOrWhiteSpace(senderAddress) ? null : senderAddress);
        }

        // ###########################################################################################
        // "Commodore/C128/250477", or a neutral phrase when the name is unusable.
        //
        // The caller decides what the system is called - the id itself in every submission mail
        // (owner wording, 2026-10-03) - so this only tidies what it is given.
        // ###########################################################################################
        private static string DescribeSystem(string? systemName)
        {
            string trimmed = systemName?.Trim() ?? string.Empty;

            return trimmed.Length == 0 ? "the board data" : trimmed;
        }

        // ###########################################################################################
        // "Hi Dennis," or "Hi there," when there is no usable name.
        //
        // A blank name is normal for a contributor who sent without an account, possible for an
        // account created before validation tightened, and FORCED for the AlreadyRegistered path's
        // caller to supply the EXISTING account's name - the name typed by whoever is registering
        // must never be used, since that would let a stranger put text into a mail sent to
        // somebody else.
        // ###########################################################################################
        private static MailBody Greeted(string? name)
        {
            string trimmed = name?.Trim() ?? string.Empty;

            return new MailBody().Paragraph(trimmed.Length == 0 ? "Hi there," : $"Hi {trimmed},");
        }

        // The closing line: who to write to with a question.
        private static MailBody WithContactLine(this MailBody body) =>
            body.Paragraph(
                $"If you have any questions about this, just get in touch with {EmailTemplates.ContactName} at ",
                MailText.Link(EmailTemplates.ContactAddress, $"mailto:{EmailTemplates.ContactAddress}"),
                ".");
    }

    // One of CRT's own text files, shown in the feedback mail under its heading ("Logfile").
    public sealed record FeedbackShownFile(string Title, string Text);
}
