using CRT;
using Handlers.DataHandling;
using Handlers.Theming;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// What the component pane says with several components picked (SeveralComponentsWording; owner request, 2026-11-08: "It
// must be very clear for the user, that whatever the user changes will affect all selected components").
// ###########################################################################################
public sealed class SeveralComponentsWordingTests
{
    [Fact]
    public void A_few_are_named_and_many_are_counted_past_the_first()
    {
        Assert.Equal("R21, R22", SeveralComponentsWording.Labels(["R21", "R22"]));

        string[] many = [.. Enumerable.Range(1, 9).Select(i => $"R{i}")];
        Assert.Equal("R1, R2, R3, R4, R5, R6 and 3 more", SeveralComponentsWording.Labels(many));
        Assert.Equal("Components: R21, R22", SeveralComponentsWording.Title(["R21", "R22"]));
    }

    // ###########################################################################################
    // How many and who, the number and each name in bold - not the words between them (owner requests, 2026-11-09: "show the
    // amount of components and the component names with bold", then "it should be "Editing <b>2</b> components:
    // <b>RP7</b>,<b>RP8</b> - every...""). Past the first six, the rest counted in plain words.
    // ###########################################################################################
    [Fact]
    public void Editing_says_how_many_and_who_in_bold_and_that_every_change_reaches_all_of_them()
    {
        string editing = SeveralComponentsWording.Editing(["RP7", "RP8"]);

        Assert.Equal("Editing 2 components: RP7, RP8 - every change you make here is made to all of them.", BoldRuns.Plain(editing));
        Assert.Equal(["2", "RP7", "RP8"], BoldRuns.Marked(editing).Where(run => run.IsBold).Select(run => run.Text).ToArray());

        string many = SeveralComponentsWording.Editing([.. Enumerable.Range(1, 9).Select(i => $"R{i}")]);
        Assert.StartsWith("Editing 9 components: R1, R2, R3, R4, R5, R6 and 3 more - ", BoldRuns.Plain(many), StringComparison.Ordinal);
        Assert.Equal(["9", "R1", "R2", "R3", "R4", "R5", "R6"], BoldRuns.Marked(many).Where(run => run.IsBold).Select(run => run.Text).ToArray());

        // A label holding a mark of its own bolds nothing wrong (code review, 2026-11-09).
        string marked = SeveralComponentsWording.Editing(["R<b>1", "R2</b>x"]);
        Assert.Equal(["2", "R1", "R2x"], BoldRuns.Marked(marked).Where(run => run.IsBold).Select(run => run.Text).ToArray());
        Assert.Equal("Editing 2 components: R1, R2x - every change you make here is made to all of them.", BoldRuns.Plain(marked));
    }

    // Alike, there is nothing to say (owner request, 2026-11-09: the line "they are all the same" made way for the footnote).
    [Fact]
    public void A_list_page_says_only_what_only_some_have()
    {
        Assert.Equal(string.Empty, SeveralComponentsWording.RowsLine(ComponentPage.Replacements, []));

        string some = SeveralComponentsWording.RowsLine(ComponentPage.Links, [("R21", "Datasheet (https://example.org)")]);
        Assert.Contains("Only some have R21: Datasheet (https://example.org)", some, StringComparison.Ordinal);
        Assert.Contains("a change reaches every one", some, StringComparison.Ordinal);
    }

    // A field they differ in says so in its own box ("1 field here differs ..." over "Basics" went - owner request, 2026-11-09).
    [Fact]
    public void A_field_they_differ_in_says_how_many_values_it_has()
    {
        Assert.Equal("Differs - 2 values", SeveralComponentsWording.Differs(2));
    }

    // ###########################################################################################
    // "MAKE ALL LIKE" SAYS IT CHANGES THE PAGE SHOWN ALONE (owner request, 2026-11-09: "Make all have same data as [RP7]
    // *1" and "*1 Pressing button will only change data on this page." - "word this better, if you have any ideas").
    // ###########################################################################################
    [Fact]
    public void Make_all_like_names_the_one_and_its_note_names_the_page_it_changes_alone()
    {
        Assert.Equal("Make all have the same data as [RP7]", SeveralComponentsWording.MakeAllLike("RP7"));
        Assert.Equal("*1", SeveralComponentsWording.FootnoteMark);
        Assert.Equal(
            "The button changes only the data on this page, \"Basics\" - the components' other pages stay as they are.",
            SeveralComponentsWording.MakeAllLikeNote(ComponentPage.Basics));
    }

    [Fact]
    public void Packages_that_differ_are_named_each_once_and_none_as_none()
    {
        // The package box stays on "Pinout" - "Basics" has none (code review, 2026-11-09: the words sent them there).
        Assert.Equal(
            "Their packages differ (DIP-14, none) - give them one package in \"Package\" below to edit their pins together, or pick one of them on its own to change its pins.",
            SeveralComponentsWording.PackagesDiffer(["DIP-14", "", "DIP-14"]));
    }
}
