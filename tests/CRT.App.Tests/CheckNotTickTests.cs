using System.Text.RegularExpressions;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// *** "SELECT" OR "CHECK", NEVER "TICK", IN ANYTHING A PERSON READS (owner request, 2026-11-20: "Never use this term "tick" and
// instead use "Select" or "Check" depending on the situation. Do that always, and check everywhere if there are other places this
// is mentioned"). *** "Check" for a checkbox ("check "Enable Worklog"", "unchecked"), "Select" for what is picked from a list
// ("Select the pins to capture"). Every string literal in src/ - the app, the shared library, the server's mails - and every text
// attribute of the markup is held to it, as AmericanSpellingTests holds "colour". Code names and comments are not text a person
// reads; nor is a member of a time ("x.Ticks", "ElapsedTicks"). The Wiki is checked by hand (tests read no Markdown).
// ###########################################################################################
public sealed class CheckNotTickTests
{
    // "tick", "ticks", "ticked", "ticking", "untick..." as a word - not part of a longer name, not a member after a dot.
    private static readonly Regex Tick = new(@"(?<![\w.{])(un)?tick(s|ed|ing)?\b", RegexOptions.IgnoreCase | RegexOptions.Compiled);

    [Fact]
    public void No_string_a_person_reads_says_tick()
    {
        List<string> offenders = AmericanSpellingTests.OffendersOf(Tick);

        Assert.True(offenders.Count == 0, "\"select\" or \"check\", never \"tick\", in what a person reads (owner request, 2026-11-20):\n" + string.Join("\n", offenders));
    }

    // The detector finds the word however it is bent, and leaves comments, code names and a time's ticks alone.
    [Theory]
    [InlineData("""    public const string Hint = "Tick the pins to capture";""", true)]
    [InlineData("""        Text = $"so this stays ticked until {when}",""", true)]
    [InlineData("""    "untick \"Show worklogs\" to hide them" +""", true)]
    [InlineData("""    "Check the pins to capture" +""", false)]
    [InlineData("""    // Tick the pins - a comment.""", false)]
    [InlineData("""    string stamp = $"{info.Length}|{info.LastWriteTimeUtc.Ticks}";""", false)]
    [InlineData("""    "the stopwatch's ElapsedTicks are logged",""", false)]
    [InlineData("""    "a sticker on the chip",""", false)]
    public void The_detector_reads_literals_only(string line, bool found)
    {
        Assert.Equal(found, AmericanSpellingTests.LiteralHolds(line, Tick));
    }

    [Theory]
    [InlineData("ToolTip.Tip", "Pick none - then tick the pins you want", true)]
    [InlineData("Content", "Untick all", true)]
    [InlineData("ToolTip.Tip", "Pick none - then select the pins you want", false)]
    [InlineData("x:Name", "TickBox", false)]
    public void Markup_text_is_held_and_names_are_not(string name, string value, bool found)
    {
        Assert.Equal(found, AmericanSpellingTests.MarkupHolds(name, value, Tick));
    }
}
