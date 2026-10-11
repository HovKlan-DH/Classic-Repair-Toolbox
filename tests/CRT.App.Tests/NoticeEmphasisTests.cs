using Handlers.DataHandling;
using Handlers.Theming;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// NoticeEmphasis - what a notice card over the tabs puts in bold (owner request, 2026-10-11:
// "Checking data from the <b>stable</b> source", "[<b>3</b>] files updated in background", "and do
// that for others also"). A bracketed value and a data source's name are bold; a quoted label is
// left as it reads on its control. Each case is written as the bold pieces between * marks.
// ###########################################################################################
public sealed class NoticeEmphasisTests
{
    [Theory]
    [InlineData("[3] files updated in background - please refresh board", "[*3*] files updated in background - please refresh board")]
    [InlineData("[1] file updated - please refresh board", "[*1*] file updated - please refresh board")]
    [InlineData("Version [3.0.0-beta.5] is available (simulated)", "Version [*3.0.0-beta.5*] is available (simulated)")]
    [InlineData("Checking data from the stable source - please wait...", "Checking data from the *stable* source - please wait...")]
    [InlineData("Checking data from the BETA source - please wait...", "Checking data from the *BETA* source - please wait...")]
    [InlineData("All data files are up to date with the BETA source", "All data files are up to date with the *BETA* source")]
    [InlineData("Could not check the stable source for new data - see the log", "Could not check the *stable* source for new data - see the log")]
    [InlineData(
        "Checking data from the BETA source - please wait...\nDownloading file [2] of [9] from the BETA source:\nCommodore/C64/250407/pcb.png",
        "Checking data from the *BETA* source - please wait...\nDownloading file [*2*] of [*9*] from the *BETA* source:\nCommodore/C64/250407/pcb.png")]
    public void A_bracketed_value_and_a_source_name_are_bold(string text, string expected)
    {
        Assert.Equal(expected, Marked(text));
    }

    // Every outcome the data-check card can say, through the rule as the card applies it.
    [Fact]
    public void Every_data_check_outcome_bolds_its_count_or_its_source()
    {
        Assert.Equal("[*3*] files updated in background - please refresh board", Marked(DataCheckNotice.After(3, atLaunch: true, "BETA source").Text));
        Assert.Equal("[*1*] file updated - please refresh board", Marked(DataCheckNotice.After(1, atLaunch: false, "stable source").Text));
        Assert.Equal("All data files are up to date with the *stable* source", Marked(DataCheckNotice.After(0, atLaunch: true, "stable source").Text));
        Assert.Equal("Could not check the *BETA* source for new data - see the log", Marked(DataCheckNotice.After(-1, atLaunch: false, "BETA source").Text));
        Assert.Equal("Checking data from the *stable* source - please wait...", Marked(DataSyncStatusTip.Checking("stable source", "", "")));
    }

    // A quoted label reads as it does on its control - the BETA-try notice names the Configuration
    // tab's check box, which itself names both sources.
    [Fact]
    public void A_quoted_label_is_left_as_it_reads_on_its_control()
    {
        string text = $"Commodore / C64 / 250407 is now in the BETA source. To try it before everyone else, tick \"{ConfigurationWording.BetaSourceCheckBox}\" on the Configuration tab.";

        Assert.Equal(
            $"Commodore / C64 / 250407 is now in the *BETA* source. To try it before everyone else, tick \"{ConfigurationWording.BetaSourceCheckBox}\" on the Configuration tab.",
            Marked(text));
        Assert.Contains("source", ConfigurationWording.BetaSourceCheckBox, StringComparison.Ordinal);
    }

    // An opening quote with no closing one quotes nothing.
    [Fact]
    public void An_unclosed_quote_quotes_nothing()
    {
        Assert.Equal("Say \"[*2*] of the *stable* source", Marked("Say \"[2] of the stable source"));
    }

    // Only a source's NAME before "source" - not the word anywhere else, nor "source" itself.
    [Theory]
    [InlineData("A stable connection", "A stable connection")]
    [InlineData("BETA versions are offered", "BETA versions are offered")]
    [InlineData("the online source", "the online source")]
    [InlineData("the stablesource", "the stablesource")]
    [InlineData("the stable sources", "the stable sources")]
    [InlineData("the Stable source", "the *Stable* source")]
    public void Only_a_source_name_before_source_is_bold(string text, string expected)
    {
        Assert.Equal(expected, Marked(text));
    }

    // Brackets that hold nothing, span a line or never close are not values.
    [Theory]
    [InlineData("Empty [] stays", "Empty [] stays")]
    [InlineData("Open [3 never closes", "Open [3 never closes")]
    [InlineData("Across [3\n4] lines", "Across [3\n4] lines")]
    [InlineData("Nested [[3]]", "Nested [[*3*]]")]
    public void Brackets_without_a_value_are_left_alone(string text, string expected)
    {
        Assert.Equal(expected, Marked(text));
    }

    // Joined again, the pieces are the words exactly - nothing lost, nothing added - and words
    // with nothing to emphasise are one plain piece (a plain Text, which lays out cheaper).
    [Theory]
    [InlineData("[3] files updated in background - please refresh board")]
    [InlineData("Downloading update...")]
    [InlineData("Line one\r\n\r\nLine two from the BETA source [7]")]
    [InlineData("tick \"Download data from the BETA source\" now")]
    public void The_pieces_are_the_words_exactly(string text)
    {
        Assert.Equal(text, string.Concat(NoticeEmphasis.Runs(text).Select(run => run.Text)));
    }

    [Fact]
    public void Words_with_nothing_to_emphasise_are_one_plain_piece_and_no_words_none()
    {
        Assert.Equal([("Downloading update: 40%", false)], NoticeEmphasis.Runs("Downloading update: 40%"));
        Assert.Equal([("Say \"[3]\" here", false)], NoticeEmphasis.Runs("Say \"[3]\" here"));
        Assert.Empty(NoticeEmphasis.Runs(""));
        Assert.Empty(NoticeEmphasis.Runs(null));
    }

    // The bold pieces between * marks, so a case reads as the card does.
    private static string Marked(string text) =>
        string.Concat(NoticeEmphasis.Runs(text).Select(run => run.IsBold ? $"*{run.Text}*" : run.Text));
}
