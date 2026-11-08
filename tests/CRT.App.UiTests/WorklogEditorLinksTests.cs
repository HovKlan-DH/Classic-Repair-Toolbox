using CRT;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// A worklog's "Links of interest" opened from the worklog editor - WorklogEntryEditorWindow.SubLists.cs. The opening
// stands in through its seam, as a test starts no browser (CLAUDE.md, test rule 6).
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class WorklogEditorLinksTests
{
    // ###########################################################################################
    // *** A LINK TYPED WITHOUT ITS SCHEME OPENS AS THE WEB ADDRESS IT IS (code review, 2026-11-08). *** A link is stored as
    // typed: handed over as it stood, "www.c64-wiki.com" was tried as a file and refused, and the error card said it could
    // not be opened - while the exported PDF opened it (LinkTarget, the one rule). A mail address stays one, and what is no
    // address CRT opens is not handed over.
    // ###########################################################################################
    [Fact]
    public void A_link_opens_by_the_one_rule_a_scheme_given_where_none_was_typed()
    {
        UiTest.Run(() =>
        {
            var window = new WorklogEntryEditorWindow();
            var opened = new List<string>();
            window.OpenLinkOverrideForTests = opened.Add;

            window.OpenLink("www.c64-wiki.com");
            window.OpenLink("https://zimmers.net");
            window.OpenLink("mailto:anna@example.org");
            window.OpenLink("ftp://example.org/rom.bin");

            Assert.Equal(["https://www.c64-wiki.com", "https://zimmers.net", "mailto:anna@example.org"], opened);
        });
    }
}
