using Handlers.MaintainerHandling;

namespace ClassicRepairToolbox.Tests.Maintainer;

// ###########################################################################################
// PersonRuns - every maintainer and contributor the Maintainer tab names is in bold (owner request,
// 2026-10-09: "do 'bold' all maintainers and contributors in the UI, so it gets more visibility who
// we are talking about"), by name where there is one, otherwise by address.
// ###########################################################################################
public sealed class PersonRunsTests
{
    private static List<(string, bool)> Runs(IEnumerable<ReviewNoteRun> runs) => runs.Select(run => (run.Text, run.IsCount)).ToList();

    [Fact]
    public void A_name_is_bold_and_its_address_beside_it_plain()
    {
        Assert.Equal([("Anna", true), (" (anna@example.com)", false)], Runs(PersonRuns.NameAndAddress(" Anna ", " anna@example.com ", "nobody")));
    }

    [Fact]
    public void With_no_name_the_address_is_the_person_and_with_neither_the_fallback_is_plain()
    {
        Assert.Equal([("anna@example.com", true)], Runs(PersonRuns.NameAndAddress(null, "anna@example.com", "nobody")));
        Assert.Equal([("Anna", true)], Runs(PersonRuns.NameAndAddress("Anna", "  ", "nobody")));
        Assert.Equal([("(no address)", false)], Runs(PersonRuns.NameAndAddress("", null, "(no address)")));
    }

    [Fact]
    public void A_sentences_person_is_bold_and_nobody_is_said_plainly()
    {
        Assert.Equal(("Dennis", true), (PersonRuns.PersonOr(" Dennis ", "somebody").Text, PersonRuns.PersonOr(" Dennis ", "somebody").IsCount));
        Assert.Equal(("somebody", false), (PersonRuns.PersonOr(null, "somebody").Text, PersonRuns.PersonOr(null, "somebody").IsCount));
        Assert.Equal("by Dennis", PersonRuns.Text([PersonRuns.Plain("by "), PersonRuns.Person("Dennis")]));
    }
}
