using CRT.Server.Handlers.Accounts;
using CRT.Server.Handlers.Submissions;
using CRT.Server.Handlers.Email;

namespace CRT.Server.Tests
{
    // ###########################################################################################
    // Covers the four mails this service sends.
    //
    // THE ANTI-ENUMERATION DESIGN IS TESTED HERE, because that is where it lives: registration
    // answers the same neutral 202 whatever the address turns out to be, and the only thing that
    // differs is WHICH MAIL is sent - visible solely to whoever controls the mailbox.
    //
    // Mail content is worth pinning precisely because it is the one part of this service a user
    // reads directly, and because a template that lost its link, or carried the wrong one, would
    // strand every new account without anything failing.
    // ###########################################################################################
    public class EmailTemplatesTests
    {
        private const string VerifyUrl = "https://classic-repair-toolbox.dk/api/accounts/verify?token=abc123";
        // *** A CODE, NOT A URL - and this constant used to hold the URL that shipped broken. ***
        // The reset mail printed "https://.../api/accounts/reset?token=...", a path the server maps
        // nothing at (only POST /reset-password exists), so every reset link 404'd from Phase 3
        // until the project owner clicked one on 2026-09-22. These tests asserted the dead URL was
        // present, which is why they never caught it.
        private const string ResetCode = "iAXr2z0PPffPwpmzHR-bOFo5ZPCPfHEK1hZbFAKAoYQ";

        // -----------------------------------------------------------------------------------
        // The maintainer's "something is waiting" mail (Phase 6 task 11).
        // -----------------------------------------------------------------------------------

        [Fact]
        public void The_waiting_mail_names_the_board_the_submission_and_what_the_contributor_said()
        {
            EmailMessage message = EmailTemplates.SubmissionWaiting(
                "anna@example.com", "Commodore/C64/250407", 42, "Corrected R12.");

            Assert.Equal("anna@example.com", message.ToAddress);
            Assert.Contains("Commodore/C64/250407", message.Body, StringComparison.Ordinal);
            Assert.Contains("#42", message.Body, StringComparison.Ordinal);
            Assert.Contains("Corrected R12.", message.Body, StringComparison.Ordinal);

            // Where to look: CRT's Maintainer tab since 2026-09-29, when the separate CRT Maintainer
            // application was folded into CRT - and never the application that no longer exists.
            Assert.Contains("the \"Maintainer\" tab", message.Body, StringComparison.Ordinal);
            Assert.DoesNotContain("CRT Maintainer", message.Body, StringComparison.Ordinal);
        }

        [Fact]
        public void The_waiting_mail_says_so_when_the_contributor_gave_no_description()
        {
            // A blank quote reads as a rendering fault; the ordinary "(no description given)"
            // the Maintainer tab's queue also shows is used instead.
            EmailMessage message = EmailTemplates.SubmissionWaiting("anna@example.com", "X/Y/Z", 1, "  ");

            Assert.Contains("(no description given)", message.Body, StringComparison.Ordinal);
        }

        [Fact]
        public void The_waiting_mail_carries_no_link()
        {
            // Like every other mail here - there is nothing to click, the application is named.
            EmailMessage message = EmailTemplates.SubmissionWaiting("anna@example.com", "X/Y/Z", 1, "x");

            Assert.DoesNotContain("http", message.Body, StringComparison.OrdinalIgnoreCase);
        }

        // -----------------------------------------------------------------------------------
        // The two-stage publish's contributor mails (2026-09-25).
        // -----------------------------------------------------------------------------------

        [Fact]
        public void The_first_mail_says_it_is_in_the_BETA_SOURCE_how_to_try_it_and_that_more_is_coming()
        {
            // The project owner's words (2026-10-03): accepted and ready to test in the BETA source,
            // the Configuration tab's check box named exactly as CRT writes it, and the promise of
            // a second mail either way.
            EmailMessage message = EmailTemplates.SubmissionPublishedToBeta("c@example.com", "Commodore/C64/250407", null);

            Assert.Equal("Your CRT contribution is accepted and ready to test in the BETA source", message.Subject);
            Assert.Contains("[Commodore/C64/250407]", message.Body, StringComparison.Ordinal);
            Assert.Contains("BETA online source", message.Body, StringComparison.Ordinal);
            Assert.Contains("\"Configuration\" tab", message.Body, StringComparison.Ordinal);
            Assert.Contains("Download data from the BETA source instead of the stable source", message.Body, StringComparison.Ordinal);
            Assert.Contains("you will hear from us again", message.Body, StringComparison.Ordinal);
            Assert.DoesNotContain("http", message.Body, StringComparison.OrdinalIgnoreCase);
        }

        // A maintainer changed rows in the Maintainer tab before publishing (2026-09-25): the
        // contributor is told, and only then - and through the notifier's mapping too.
        [Fact]
        public void The_first_mail_says_when_a_maintainer_changed_the_submission()
        {
            EmailMessage changed = EmailTemplates.SubmissionPublishedToBeta("c@example.com", "X/Y/Z", null, amendedByMaintainer: true);
            EmailMessage unchanged = EmailTemplates.SubmissionPublishedToBeta("c@example.com", "X/Y/Z", null);

            Assert.Contains("The maintainer changed some of the details", changed.Body, StringComparison.Ordinal);
            Assert.DoesNotContain("maintainer changed", unchanged.Body, StringComparison.Ordinal);

            EmailMessage? mapped = SubmissionNotifier.BuildMessage("c@example.com", "X/Y/Z", "merged", null, amendedByMaintainer: true);
            Assert.Contains("The maintainer changed some of the details", mapped!.Body, StringComparison.Ordinal);
        }

        [Fact]
        public void The_second_mail_says_published_to_the_SOURCE_and_is_a_different_mail()
        {
            EmailMessage beta = EmailTemplates.SubmissionPublishedToBeta("c@example.com", "X/Y/Z", null);
            EmailMessage source = EmailTemplates.SubmissionPublishedToSource("c@example.com", "X/Y/Z");

            Assert.Contains("published to the stable source", source.Subject, StringComparison.Ordinal);
            Assert.DoesNotContain("BETA", source.Subject, StringComparison.Ordinal);
            Assert.NotEqual(beta.Subject, source.Subject);
            Assert.NotEqual(beta.Body, source.Body);
        }

        // -----------------------------------------------------------------------------------
        // Verification.
        // -----------------------------------------------------------------------------------

        [Fact]
        public void The_verification_mail_carries_the_link_and_says_how_long_it_lasts()
        {
            EmailMessage message = EmailTemplates.Verification(
                "dennis@example.com", "Dennis", EmailTemplatesTests.VerifyUrl, 24);

            Assert.Equal("dennis@example.com", message.ToAddress);
            Assert.Contains(EmailTemplatesTests.VerifyUrl, message.Body, StringComparison.Ordinal);
            Assert.Contains("24 hours", message.Body, StringComparison.Ordinal);
            Assert.Contains("Hi Dennis,", message.Body, StringComparison.Ordinal);
        }

        [Fact]
        public void The_verification_mail_tells_an_uninvolved_recipient_they_can_ignore_it()
        {
            // Someone may receive this because another person typed their address by mistake.
            // They need to know that doing nothing is safe and sufficient.
            EmailMessage message = EmailTemplates.Verification(
                "dennis@example.com", "Dennis", EmailTemplatesTests.VerifyUrl, 24);

            Assert.Contains("did not sign up", message.Body, StringComparison.Ordinal);
        }

        // -----------------------------------------------------------------------------------
        // The already-registered mail. This is what makes registration non-enumerable.
        // -----------------------------------------------------------------------------------

        [Fact]
        public void The_already_registered_mail_carries_a_reset_link_rather_than_a_verification_one()
        {
            // Someone re-registering is usually a person who forgot they had an account, so the
            // useful answer is a way back in.
            EmailMessage message = EmailTemplates.AlreadyRegistered(
                "dennis@example.com", "Dennis", EmailTemplatesTests.ResetCode, 2);

            Assert.Contains(EmailTemplatesTests.ResetCode, message.Body, StringComparison.Ordinal);
            Assert.DoesNotContain(EmailTemplatesTests.VerifyUrl, message.Body, StringComparison.Ordinal);
        }

        [Fact]
        public void The_already_registered_mail_tells_the_owner_that_somebody_tried()
        {
            // If it was not them, this is a security notification they would want.
            EmailMessage message = EmailTemplates.AlreadyRegistered(
                "dennis@example.com", "Dennis", EmailTemplatesTests.ResetCode, 2);

            Assert.Contains("Someone just tried", message.Body, StringComparison.Ordinal);
            Assert.Contains("not you", message.Body, StringComparison.Ordinal);
        }

        [Fact]
        public void The_already_registered_mail_states_that_the_prober_learned_nothing()
        {
            // The reassurance is the point: the recipient should understand that their account is
            // not exposed by this message existing.
            EmailMessage message = EmailTemplates.AlreadyRegistered(
                "dennis@example.com", "Dennis", EmailTemplatesTests.ResetCode, 2);

            Assert.Contains("was not told whether this address is registered", message.Body, StringComparison.Ordinal);
        }

        [Fact]
        public void A_display_name_supplied_by_a_STRANGER_is_never_used_in_the_already_registered_mail()
        {
            // THE injection this mail is exposed to. Whoever is attempting the registration
            // supplies a display name, but the mail goes to the EXISTING account's owner - so
            // using the attacker's text would let a stranger put arbitrary words into a message
            // sent to somebody else's inbox. The caller must pass the account's own name, or
            // nothing; passing nothing must produce a clean greeting rather than an empty one.
            EmailMessage message = EmailTemplates.AlreadyRegistered(
                "dennis@example.com", displayName: null!, EmailTemplatesTests.ResetCode, 2);

            Assert.Contains("Hi there,", message.Body, StringComparison.Ordinal);
            Assert.DoesNotContain("Hi ,", message.Body, StringComparison.Ordinal);
        }

        // -----------------------------------------------------------------------------------
        // Password reset and the change notification.
        // -----------------------------------------------------------------------------------

        [Fact]
        public void The_reset_mail_says_the_code_is_single_use()
        {
            EmailMessage message = EmailTemplates.PasswordReset(
                "dennis@example.com", "Dennis", EmailTemplatesTests.ResetCode, 2);

            Assert.Contains(EmailTemplatesTests.ResetCode, message.Body, StringComparison.Ordinal);
            Assert.Contains("once", message.Body, StringComparison.Ordinal);
        }

        [Fact]
        public void The_reset_mail_contains_NO_CLICKABLE_LINK_at_all()
        {
            // *** THE REGRESSION TEST FOR THE 404 THAT REACHED A REAL USER. *** This mail used to
            // print "https://<host>/api/accounts/reset?token=...". Nothing is mapped at that path -
            // completing a reset needs a new PASSWORD, so it is POST /reset-password, which a
            // browser click cannot reach. Every reset mail sent between Phase 3 and 2026-09-22 led
            // to "This page can't be found".
            //
            // Asserting on the ABSENCE of a URL rather than on the code's presence is deliberate:
            // a future edit that helpfully adds "or click here" would reintroduce exactly the
            // original bug, and only this shape of assertion catches that.
            EmailMessage message = EmailTemplates.PasswordReset(
                "dennis@example.com", "Dennis", EmailTemplatesTests.ResetCode, 2);

            Assert.DoesNotContain("http://", message.Body, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("https://", message.Body, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("token=", message.Body, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void The_ALREADY_REGISTERED_mail_carries_no_link_either()
        {
            // It offers a reset for the same reason and carried the same dead URL.
            EmailMessage message = EmailTemplates.AlreadyRegistered(
                "dennis@example.com", "Dennis", EmailTemplatesTests.ResetCode, 2);

            Assert.DoesNotContain("https://", message.Body, StringComparison.OrdinalIgnoreCase);
            Assert.Contains(EmailTemplatesTests.ResetCode, message.Body, StringComparison.Ordinal);
        }

        [Fact]
        public void The_reset_mail_SAYS_WHERE_TO_PASTE_THE_CODE()
        {
            // A bare code with no instructions is worse than the broken link was: at least the
            // link said what to do with it. The mail must name the application and the button.
            EmailMessage message = EmailTemplates.PasswordReset(
                "dennis@example.com", "Dennis", EmailTemplatesTests.ResetCode, 2);

            Assert.Contains("the \"Maintainer\" tab", message.Body, StringComparison.Ordinal);
            Assert.Contains("I forgot my password", message.Body, StringComparison.Ordinal);
            Assert.DoesNotContain("CRT Maintainer", message.Body, StringComparison.Ordinal);
        }

        [Fact]
        public void The_reset_mail_reassures_someone_who_did_not_ask_for_it()
        {
            EmailMessage message = EmailTemplates.PasswordReset(
                "dennis@example.com", "Dennis", EmailTemplatesTests.ResetCode, 2);

            Assert.Contains("password has not changed", message.Body, StringComparison.Ordinal);
        }

        [Fact]
        public void The_password_changed_mail_contains_NO_link()
        {
            // Deliberate: a notification that asks for no action cannot be phished. It is the one
            // message that tells someone whose account was taken over that it happened, so it must
            // not train them to click links in mail about their password.
            EmailMessage message = EmailTemplates.PasswordChanged("dennis@example.com", "Dennis");

            Assert.DoesNotContain("http://", message.Body, StringComparison.Ordinal);
            Assert.DoesNotContain("https://", message.Body, StringComparison.Ordinal);
        }

        [Fact]
        public void The_password_changed_mail_says_what_to_do_if_it_was_not_you()
        {
            EmailMessage message = EmailTemplates.PasswordChanged("dennis@example.com", "Dennis");

            Assert.Contains("I forgot my password", message.Body, StringComparison.Ordinal);
        }

        // -----------------------------------------------------------------------------------
        // A maintainer changing their address (2026-10-03).
        // -----------------------------------------------------------------------------------

        // The code mail: the code on a line of its own, where to paste it, no link to click - the
        // reset mail's rules, for the reset mail's reasons.
        [Fact]
        public void The_address_code_mail_carries_the_code_says_where_it_goes_and_has_no_link()
        {
            EmailMessage message = EmailTemplates.EmailChangeCode("bench@example.com", "Dennis", EmailTemplatesTests.ResetCode, 24);

            Assert.Equal("bench@example.com", message.ToAddress);
            Assert.Contains(message.Body.Split("\r\n"), line => line.Trim() == EmailTemplatesTests.ResetCode);
            // The way in, in the Maintainer tab's own words (owner wording, 2026-10-03): since
            // 2026-10-04 the "Account" tab and its "My account" entry, where the "Your account"
            // window's sections now are. Written out, so a change to the shared words fails here.
            Assert.Contains("choose \"Account\" and \"My account\", then \"I have received a code\"", message.Body, StringComparison.Ordinal);
            Assert.Contains("24 hours", message.Body, StringComparison.Ordinal);
            Assert.Contains("nothing changes unless the code is used", message.Body, StringComparison.Ordinal);
            Assert.DoesNotContain("https://", message.Body, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("token=", message.Body, StringComparison.OrdinalIgnoreCase);
        }

        // ###########################################################################################
        // To an address another account already holds: no code, nothing that could move anything,
        // and the reassurance that whoever asked learned nothing - the AlreadyRegistered rule.
        // ###########################################################################################
        [Fact]
        public void The_address_taken_mail_carries_no_code_and_says_the_asker_learned_nothing()
        {
            EmailMessage message = EmailTemplates.EmailChangeAddressTaken("anna@example.com", "Anna");

            Assert.StartsWith("Hi Anna,", message.Body, StringComparison.Ordinal);
            Assert.Contains("nothing was changed", message.Body, StringComparison.Ordinal);
            Assert.Contains("was not told whether this address is registered", message.Body, StringComparison.Ordinal);
            Assert.DoesNotContain(EmailTemplatesTests.ResetCode, message.Body, StringComparison.Ordinal);
        }

        // ###########################################################################################
        // To the OLD address once it changed: the one notice that reaches the owner if it was not
        // them. It names the new address, and - since this mailbox can no longer reset the
        // password - who to write to, straight away.
        // ###########################################################################################
        [Fact]
        public void The_address_changed_mail_names_the_new_address_and_who_to_write_to()
        {
            EmailMessage message = EmailTemplates.EmailChanged("dennis@example.com", "Dennis", "bench@example.com");

            Assert.Equal("dennis@example.com", message.ToAddress);
            Assert.Contains("bench@example.com", message.Body, StringComparison.Ordinal);
            Assert.Contains("<b>bench@example.com</b>", message.HtmlBody, StringComparison.Ordinal);
            Assert.Contains("can no longer be used to reset its password", message.Body, StringComparison.Ordinal);
            Assert.EndsWith($"Write to Dennis straight away at {EmailTemplates.ContactAddress}.", message.Body, StringComparison.Ordinal);
        }

        // -----------------------------------------------------------------------------------
        // Properties shared by every mail.
        // -----------------------------------------------------------------------------------

        [Fact]
        public void Every_mail_uses_CRLF_line_endings()
        {
            // RFC 5321 requires CRLF, and a bare \n can be rejected or mangled by a strict MTA.
            // The raw string literals in the source carry whatever the file uses, so this is the
            // test that catches the normalisation being removed.
            var messages = new[]
            {
                EmailTemplates.Verification("a@b.com", "A", EmailTemplatesTests.VerifyUrl, 24),
                EmailTemplates.AlreadyRegistered("a@b.com", "A", EmailTemplatesTests.ResetCode, 2),
                EmailTemplates.PasswordReset("a@b.com", "A", EmailTemplatesTests.ResetCode, 2),
                EmailTemplates.PasswordChanged("a@b.com", "A")
            };

            foreach (EmailMessage message in messages)
            {
                // No bare \n anywhere: every \n must be preceded by \r.
                for (int index = 0; index < message.Body.Length; index++)
                {
                    if (message.Body[index] == '\n')
                        Assert.True(index > 0 && message.Body[index - 1] == '\r', "Found a bare \\n in a mail body.");
                }
            }
        }

        [Fact]
        public void Every_mail_names_the_product_in_its_subject()
        {
            // These land in an inbox beside everything else; the subject has to say what it is
            // about without being opened.
            var messages = new[]
            {
                EmailTemplates.Verification("a@b.com", "A", EmailTemplatesTests.VerifyUrl, 24),
                EmailTemplates.AlreadyRegistered("a@b.com", "A", EmailTemplatesTests.ResetCode, 2),
                EmailTemplates.PasswordReset("a@b.com", "A", EmailTemplatesTests.ResetCode, 2),
                EmailTemplates.PasswordChanged("a@b.com", "A")
            };

            foreach (EmailMessage message in messages)
                Assert.Contains("Classic Repair Toolbox", message.Subject, StringComparison.Ordinal);
        }

        [Fact]
        public void No_mail_assumes_the_reader_is_doing_this_for_a_living()
        {
            // CRT's audience is hobbyists repairing their own machines. No corporate register:
            // nothing about provisioning, policy, compliance or an organisation.
            string[] forbidden = ["provision", "policy", "compliance", "your organisation", "your organization", "employee"];

            var messages = new[]
            {
                EmailTemplates.Verification("a@b.com", "A", EmailTemplatesTests.VerifyUrl, 24),
                EmailTemplates.AlreadyRegistered("a@b.com", "A", EmailTemplatesTests.ResetCode, 2),
                EmailTemplates.PasswordReset("a@b.com", "A", EmailTemplatesTests.ResetCode, 2),
                EmailTemplates.PasswordChanged("a@b.com", "A")
            };

            foreach (EmailMessage message in messages)
            {
                foreach (string word in forbidden)
                {
                    Assert.DoesNotContain(word, message.Body, StringComparison.OrdinalIgnoreCase);
                    Assert.DoesNotContain(word, message.Subject, StringComparison.OrdinalIgnoreCase);
                }
            }
        }

        // -----------------------------------------------------------------------------------
        // HTML, greetings and the contact line (owner request, 2026-10-03).
        // -----------------------------------------------------------------------------------

        private const string C64 = "Commodore/C64/250407";

        // One of every mail, each with whatever a test needs to find in it.
        private static IReadOnlyList<EmailMessage> EveryMail() =>
        [
            EmailTemplates.Verification("a@b.com", "A", EmailTemplatesTests.VerifyUrl, 24),
            EmailTemplates.AlreadyRegistered("a@b.com", "A", EmailTemplatesTests.ResetCode, 2),
            EmailTemplates.PasswordReset("a@b.com", "A", EmailTemplatesTests.ResetCode, 2),
            EmailTemplates.PasswordChanged("a@b.com", "A"),
            EmailTemplates.EmailChangeCode("a@b.com", "A", EmailTemplatesTests.ResetCode, 24),
            EmailTemplates.EmailChangeAddressTaken("a@b.com", "A"),
            EmailTemplates.EmailChanged("a@b.com", "A", "c@d.com"),
            EmailTemplates.MaintainerInvitation("a@b.com", EmailTemplatesTests.C64, "Dennis", "CODE-123", 7),
            EmailTemplates.SubmissionWaiting("a@b.com", EmailTemplatesTests.C64, 18, "Fixed R12."),
            EmailTemplates.SubmissionChangesRequested("a@b.com", EmailTemplatesTests.C64, "Add the PAL region."),
            EmailTemplates.SubmissionPublishedToBeta("a@b.com", EmailTemplatesTests.C64, null),
            EmailTemplates.SubmissionPublishedToSource("a@b.com", EmailTemplatesTests.C64),
            EmailTemplates.SubmissionRejected("a@b.com", EmailTemplatesTests.C64, "Wrong board."),
            EmailTemplates.SubmissionReturnedToQueue("a@b.com", EmailTemplatesTests.C64, "Needs a check."),
            EmailTemplates.ApprovalNeeded("a@b.com", EmailTemplatesTests.C64, "submission #18", "Anna"),
            EmailTemplates.PublishedToProduction("a@b.com", EmailTemplatesTests.C64, "Anna (anna@example.com)", null, 3),
            EmailTemplates.BoardDeleted("a@b.com", EmailTemplatesTests.C64, "It was a test board."),
        ];

        // ###########################################################################################
        // Every mail is HTML (owner request, 2026-10-03: "All mails should be sent in HTML format")
        // with the same words as plain text beside it - the part spam filters expect, which is never
        // shown by a client that can show HTML.
        // ###########################################################################################
        [Fact]
        public void Every_mail_is_HTML_with_the_same_words_in_plain_text_beside_it()
        {
            foreach (EmailMessage message in EmailTemplatesTests.EveryMail())
            {
                Assert.NotNull(message.HtmlBody);
                Assert.StartsWith("<!DOCTYPE html>", message.HtmlBody, StringComparison.Ordinal);
                Assert.Contains("<meta charset=\"utf-8\">", message.HtmlBody, StringComparison.Ordinal);

                // The greeting opens both.
                string greeting = message.Body.Split("\r\n")[0];
                Assert.Contains($"<p style=\"margin:0 0 14px 0;\">{greeting}</p>", message.HtmlBody, StringComparison.Ordinal);
            }
        }

        // *** WHAT SOMEBODY TYPED NEVER BECOMES MARKUP. *** A contributor's description and a name
        // reach a mail sent to somebody else; in the HTML they are text, whatever they contain.
        [Fact]
        public void Text_somebody_typed_is_escaped_in_the_HTML()
        {
            EmailMessage message = EmailTemplates.SubmissionWaiting(
                "anna@example.com", EmailTemplatesTests.C64, 18, "<script>alert(1)</script> & <b>bold</b>", recipientName: "<i>Eve</i>");

            Assert.DoesNotContain("<script>", message.HtmlBody, StringComparison.Ordinal);
            Assert.DoesNotContain("<i>Eve</i>", message.HtmlBody, StringComparison.Ordinal);
            Assert.Contains("&lt;script&gt;alert(1)&lt;/script&gt; &amp; &lt;b&gt;bold&lt;/b&gt;", message.HtmlBody, StringComparison.Ordinal);
            Assert.Contains("Hi &lt;i&gt;Eve&lt;/i&gt;,", message.HtmlBody, StringComparison.Ordinal);

            // The plain text is plain: the words as they were typed.
            Assert.Contains("<script>alert(1)</script> & <b>bold</b>", message.Body, StringComparison.Ordinal);
        }

        // The board in bold inside brackets, as the project owner wrote it; the other person's
        // words in italics.
        [Fact]
        public void The_board_is_bold_inside_brackets_and_the_maintainers_words_are_in_italics()
        {
            EmailMessage message = EmailTemplates.SubmissionChangesRequested("c@example.com", EmailTemplatesTests.C64, "Please view the duplicate rows.");

            Assert.Contains("[<b>Commodore/C64/250407</b>]", message.HtmlBody, StringComparison.Ordinal);
            Assert.Contains("<i>Please view the duplicate rows.</i>", message.HtmlBody, StringComparison.Ordinal);
            Assert.Contains("[Commodore/C64/250407]", message.Body, StringComparison.Ordinal);
            Assert.Contains("    Please view the duplicate rows.", message.Body, StringComparison.Ordinal);
        }

        // "Hi {name}," when a name is known, "Hi there," otherwise - in every contribution mail.
        [Fact]
        public void A_contributor_is_greeted_by_name_when_known_and_as_there_otherwise()
        {
            Assert.StartsWith("Hi Anna,", EmailTemplates.SubmissionRejected("c@example.com", EmailTemplatesTests.C64, "No.", "Anna").Body, StringComparison.Ordinal);
            Assert.StartsWith("Hi Anna,", EmailTemplates.SubmissionPublishedToSource("c@example.com", EmailTemplatesTests.C64, "Anna").Body, StringComparison.Ordinal);
            Assert.StartsWith("Hi there,", EmailTemplates.SubmissionRejected("c@example.com", EmailTemplatesTests.C64, "No.").Body, StringComparison.Ordinal);
            Assert.StartsWith("Hi there,", EmailTemplates.SubmissionChangesRequested("c@example.com", EmailTemplatesTests.C64, "Fix.", "  ").Body, StringComparison.Ordinal);
        }

        // ###########################################################################################
        // A deleted board's open submission (owner decision, 2026-10-03: "Delete them and mail the
        // contributors"). CRT's "My submissions" keeps the submission's last state, so this mail is
        // where the contributor learns it is gone - it says so, quotes the reason, and says the draft
        // on their own computer is untouched. Greeted by name like every contributor mail.
        // ###########################################################################################
        [Fact]
        public void The_board_deleted_mail_says_the_submission_is_gone_quotes_the_reason_and_keeps_the_draft()
        {
            EmailMessage message = EmailTemplates.BoardDeleted("c@example.com", "Commodore/C64/999999", "It was a test board.", "Anna");

            Assert.Equal("c@example.com", message.ToAddress);
            Assert.Equal("A board you contributed to was removed from CRT", message.Subject);
            Assert.StartsWith("Hi Anna,", message.Body, StringComparison.Ordinal);
            Assert.Contains("Commodore/C64/999999", message.Body, StringComparison.Ordinal);
            Assert.Contains("your contribution to it was removed along with it", message.Body, StringComparison.Ordinal);
            Assert.Contains("It was a test board.", message.Body, StringComparison.Ordinal);
            Assert.Contains("Your draft is still on your own computer", message.Body, StringComparison.Ordinal);

            // A reason somebody typed is never markup in the mail.
            EmailMessage hostile = EmailTemplates.BoardDeleted("c@example.com", "Commodore/C64/999999", "<b>gone</b>");
            Assert.DoesNotContain("<b>gone</b>", hostile.HtmlBody, StringComparison.Ordinal);
        }

        // An update and a completely new board are told apart (owner request, 2026-10-03).
        [Fact]
        public void The_waiting_mail_tells_a_new_board_from_an_update()
        {
            EmailMessage update = EmailTemplates.SubmissionWaiting("anna@example.com", EmailTemplatesTests.C64, 18, "Fixed R12.");
            EmailMessage newBoard = EmailTemplates.SubmissionWaiting("anna@example.com", EmailTemplatesTests.C64, 18, "A new board.", isNewBoard: true);

            Assert.Contains("an update for an existing board, [Commodore/C64/250407]", update.Body, StringComparison.Ordinal);
            Assert.Contains("The contributor describes the changes like this:", update.Body, StringComparison.Ordinal);
            Assert.Contains("a completely new board, [Commodore/C64/250407]", newBoard.Body, StringComparison.Ordinal);
            Assert.Contains("The contributor describes the new board like this:", newBoard.Body, StringComparison.Ordinal);
            Assert.Equal("A CRT contribution is waiting for your review", update.Subject);
        }

        // The stable-source mail says how to get the data (and to untick BETA), and that the draft
        // goes by itself.
        [Fact]
        public void The_stable_source_mail_says_how_to_get_the_data_and_that_the_draft_goes_by_itself()
        {
            EmailMessage message = EmailTemplates.SubmissionPublishedToSource("c@example.com", EmailTemplatesTests.C64);

            Assert.Equal("Your CRT contribution has been published to the stable source", message.Subject);
            Assert.Contains("Restart CRT", message.Body, StringComparison.Ordinal);
            Assert.Contains("untick it", message.Body, StringComparison.Ordinal);
            Assert.Contains("CRT removes your draft of it by itself", message.Body, StringComparison.Ordinal);
            Assert.Contains("<i>Classic Repair Toolbox</i>", message.HtmlBody, StringComparison.Ordinal);
        }

        // ###########################################################################################
        // Who to write to (owner request, 2026-10-03: "Please connect with Dennis, if you have any
        // questions for this") - every mail ends with it, as a mail link, except the notice to the
        // administrators about a production publish, which goes to the project owner himself.
        //
        // The "your address was changed" notice ends with its own, more urgent line naming the same
        // address: it may be the only thing the owner of a taken-over account receives, and "if you
        // have any questions" undersells it.
        // ###########################################################################################
        [Fact]
        public void Every_mail_but_the_production_notice_says_who_to_write_to()
        {
            foreach (EmailMessage message in EmailTemplatesTests.EveryMail())
            {
                bool isProductionNotice = message.Subject.EndsWith("was published to the stable source", StringComparison.Ordinal)
                    && message.Subject.StartsWith("CRT:", StringComparison.Ordinal);

                if (isProductionNotice)
                {
                    Assert.DoesNotContain(EmailTemplates.ContactAddress, message.Body, StringComparison.Ordinal);
                    continue;
                }

                string closing = message.Subject.EndsWith("address was changed", StringComparison.Ordinal)
                    ? $"Write to Dennis straight away at {EmailTemplates.ContactAddress}."
                    : $"just get in touch with Dennis at {EmailTemplates.ContactAddress}.";

                Assert.EndsWith(closing, message.Body, StringComparison.Ordinal);
                Assert.Contains($"href=\"mailto:{EmailTemplates.ContactAddress}\"", message.HtmlBody, StringComparison.Ordinal);
            }
        }

        // -----------------------------------------------------------------------------------
        // Tokens. Here because the links above are built from them.
        // -----------------------------------------------------------------------------------

        [Fact]
        public void A_token_is_url_safe()
        {
            // These travel in verification and reset links. A '+' silently becomes a space when a
            // form decoder gets hold of it, producing a token that looks right and never matches.
            for (int attempt = 0; attempt < 50; attempt++)
            {
                string token = SecureToken.Create();

                Assert.DoesNotContain('+', token);
                Assert.DoesNotContain('/', token);
                Assert.DoesNotContain('=', token);
            }
        }

        [Fact]
        public void Two_tokens_are_never_the_same()
        {
            var tokens = Enumerable.Range(0, 200).Select(_ => SecureToken.Create()).ToHashSet();

            Assert.Equal(200, tokens.Count);
        }

        [Fact]
        public void A_token_hash_is_stable_and_the_right_length_for_the_column()
        {
            // The database columns are CHAR(64) because SHA-256 hex always is.
            string token = SecureToken.Create();

            Assert.Equal(SecureToken.HashLength, SecureToken.Hash(token).Length);
            Assert.Equal(SecureToken.Hash(token), SecureToken.Hash(token));
        }

        [Fact]
        public void A_token_hash_does_not_reveal_the_token()
        {
            // Stating the obvious as a test, because the whole storage design rests on it: what
            // is stored must not contain what was sent.
            string token = SecureToken.Create();

            Assert.DoesNotContain(token, SecureToken.Hash(token), StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void Hash_comparison_handles_nulls_and_length_mismatches()
        {
            string hash = SecureToken.Hash("a token");

            Assert.True(SecureToken.HashesEqual(hash, hash));
            Assert.False(SecureToken.HashesEqual(hash, null));
            Assert.False(SecureToken.HashesEqual(null, hash));
            Assert.False(SecureToken.HashesEqual(hash, "short"));
        }
    }
}
