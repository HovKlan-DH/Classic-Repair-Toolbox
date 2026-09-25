using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// Characterisation tests for ContactLinkFormatter - turning a contributor's free-text contact
// string from the credits data into something the About tab can open.
//
// Extracted from TabAbout. Contact strings are contributed data, so they are whatever somebody
// typed. The rules that matter: email wins over the URL check, a bare "www." host is promoted to
// https, and anything unrecognised is returned untouched rather than guessed at.
public class ContactLinkFormatterTests
{
    // -------------------------------------------------------------- IsContactWebUrl

    [Theory]
    [InlineData("http://example.com", true)]
    [InlineData("https://example.com", true)]
    [InlineData("www.example.com", true)]
    [InlineData("HTTP://EXAMPLE.COM", true)]
    [InlineData("WWW.EXAMPLE.COM", true)]
    [InlineData("example.com", false)]          // no scheme and no www - not treated as a URL
    [InlineData("ftp://example.com", false)]    // deliberately unsupported
    [InlineData("", false)]
    public void A_contact_is_a_web_url_only_with_a_known_scheme_or_a_www_prefix(string contact, bool expected)
    {
        Assert.Equal(expected, ContactLinkFormatter.IsContactWebUrl(contact));
    }

    // The check is a prefix match, so a URL mentioned mid-string does not count.
    [Fact]
    public void A_url_appearing_later_in_the_string_is_not_a_web_url()
    {
        Assert.False(ContactLinkFormatter.IsContactWebUrl("see https://example.com"));
    }

    // -------------------------------------------------------------- IsContactEmail

    [Theory]
    [InlineData("someone@example.com", true)]
    [InlineData("a@b", true)]
    [InlineData("someone at example.com", false)]
    [InlineData("example.com", false)]
    [InlineData("", false)]
    public void A_contact_is_an_email_when_it_has_an_at_sign_and_no_space(string contact, bool expected)
    {
        Assert.Equal(expected, ContactLinkFormatter.IsContactEmail(contact));
    }

    // The no-space rule is what stops a sentence that happens to mention an address being turned
    // into a mailto link.
    [Fact]
    public void A_sentence_mentioning_an_address_is_not_treated_as_an_email()
    {
        Assert.False(ContactLinkFormatter.IsContactEmail("mail me at someone@example.com"));
    }

    // -------------------------------------------------------------- BuildContactHref

    [Fact]
    public void An_email_becomes_a_mailto_link()
    {
        Assert.Equal("mailto:someone@example.com",
            ContactLinkFormatter.BuildContactHref("someone@example.com"));
    }

    [Fact]
    public void A_bare_www_host_is_promoted_to_https()
    {
        Assert.Equal("https://www.example.com", ContactLinkFormatter.BuildContactHref("www.example.com"));
    }

    [Theory]
    [InlineData("http://example.com")]
    [InlineData("https://example.com")]
    public void A_url_that_already_has_a_scheme_is_returned_unchanged(string contact)
    {
        Assert.Equal(contact, ContactLinkFormatter.BuildContactHref(contact));
    }

    // Text that is neither an email nor a URL is passed through rather than being guessed at -
    // a contributor writing "ask on the forum" must not become a broken link.
    [Fact]
    public void Unrecognised_contact_text_is_returned_untouched()
    {
        Assert.Equal("ask on the forum", ContactLinkFormatter.BuildContactHref("ask on the forum"));
    }

    // The email check runs FIRST, so a string that looks like both is treated as an address.
    // "www.a@b.com" becomes a mailto, not an https link.
    [Fact]
    public void The_email_rule_wins_when_a_contact_could_be_read_as_both()
    {
        Assert.Equal("mailto:www.a@b.com", ContactLinkFormatter.BuildContactHref("www.a@b.com"));
    }

    [Fact]
    public void An_empty_contact_produces_an_empty_href()
    {
        Assert.Equal(string.Empty, ContactLinkFormatter.BuildContactHref(""));
    }

    // The https promotion is case-insensitive on the prefix but preserves the original casing.
    [Fact]
    public void Promoting_a_www_host_preserves_its_original_casing()
    {
        Assert.Equal("https://WWW.Example.COM", ContactLinkFormatter.BuildContactHref("WWW.Example.COM"));
    }
}
