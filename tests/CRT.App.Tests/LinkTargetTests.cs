using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// Where a stored link goes when opened - the Contribute tab's "Open" on a board link and the PDF
// export's hyperlinks by one rule (WorkbookZipExportTests pins the export through it).
// ###########################################################################################
public sealed class LinkTargetTests
{
    [Theory]
    [InlineData("https://www.lemon64.com/forum", "https://www.lemon64.com/forum")]
    [InlineData("  http://example.com  ", "http://example.com")]
    [InlineData("HTTPS://EXAMPLE.COM", "HTTPS://EXAMPLE.COM")]
    public void An_address_with_a_scheme_is_opened_as_it_stands(string stored, string expected)
    {
        Assert.Equal(expected, LinkTarget.For(stored));
    }

    // Typed without a scheme, as a link row often is - the browser needs one.
    [Theory]
    [InlineData("www.zimmers.net", "https://www.zimmers.net")]
    [InlineData("example.com/thread/12", "https://example.com/thread/12")]
    public void An_address_with_no_scheme_gets_https(string stored, string expected)
    {
        Assert.Equal(expected, LinkTarget.For(stored));
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("   ")]
    [InlineData("see the forum")]
    public void Nothing_or_words_is_no_link(string? stored)
    {
        Assert.Null(LinkTarget.For(stored));
    }

    // ###########################################################################################
    // A MAIL ADDRESS STAYS ONE (code review, 2026-10-30): "mailto:..." got https:// in front, which a
    // browser reads as the user name "mailto:anna" on the host example.org - "Open" went to a web site.
    // mailto: is the launcher's third scheme, so it stands as typed.
    // ###########################################################################################
    [Theory]
    [InlineData("mailto:anna@example.org", "mailto:anna@example.org")]
    [InlineData("  MAILTO:anna@example.org ", "MAILTO:anna@example.org")]
    public void A_mail_address_is_opened_as_it_stands(string stored, string expected)
    {
        Assert.Equal(expected, LinkTarget.For(stored));
    }

    // Any other scheme is no link CRT opens - https:// in front would send it somewhere else. A host
    // with a port is no scheme, and a bare "mailto:" names nobody.
    [Theory]
    [InlineData("ftp://ftp.zimmers.net/pub")]
    [InlineData("file:///home/me/manuals/c64.pdf")]
    [InlineData("mailto:")]
    public void An_address_of_another_scheme_is_no_link(string stored)
    {
        Assert.Null(LinkTarget.For(stored));
    }

    [Fact]
    public void A_host_with_a_port_still_gets_https()
    {
        Assert.Equal("https://www.example.com:8080/forum", LinkTarget.For("www.example.com:8080/forum"));
    }
}
