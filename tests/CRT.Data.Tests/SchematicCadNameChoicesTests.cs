using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// SchematicCadNameChoices - which KiCad view names a schematic image's "CAD name" drop-down offers
// in the "Schematic images" window (owner request, 2026-10-10: "so I can select a KiCad name. Once
// selected, it cannot be selected again").
public sealed class SchematicCadNameChoicesTests
{
    private static readonly string[] Views = ["Open128 - PCB Top", "Open128 - PCB Bottom", "Open128", "Reset"];

    private static IReadOnlyList<SchematicCadNameChoice> For(int imageIndex, params (string SchematicName, string CadName)[] images) =>
        SchematicCadNameChoices.For(images, imageIndex, Views);

    [Fact]
    public void No_CAD_name_comes_first_then_every_view_in_the_projects_order()
    {
        IReadOnlyList<SchematicCadNameChoice> choices = For(0, ("Top", ""));

        Assert.Equal(["", .. Views], choices.Select(choice => choice.CadName));
        Assert.Equal(["No CAD name", .. Views], choices.Select(choice => choice.Label));
        Assert.Equal(SchematicCadNameChoices.NoCadNameLabel, choices[0].Label);
    }

    // ###########################################################################################
    // *** ONCE ONE IMAGE HAS A NAME, IT IS GONE FROM EVERY OTHER IMAGE'S LIST (owner request,
    // 2026-10-10: "remove the CAD name entirely from the drop-down when it has been selected once,
    // so it will be easier and easier to select the correct one"). *** It REPLACES a test that
    // pinned the first version - listed, greyed out, "(used by ...)" - changed on purpose.
    // ###########################################################################################
    [Fact]
    public void A_name_another_image_has_is_not_offered_at_all()
    {
        IReadOnlyList<SchematicCadNameChoice> choices = For(
            2,
            ("C128 Neo.3.F", "open128 - pcb top"),
            ("C128 Neo.3.B", "Reset"),
            ("Bottom scan", ""));

        Assert.Equal(["", "Open128 - PCB Bottom", "Open128"], choices.Select(choice => choice.CadName));
        Assert.DoesNotContain(choices, choice => choice.Label.Contains("used by", StringComparison.Ordinal));
    }

    // The list gets shorter as names are picked - and longer again when one is given back.
    [Fact]
    public void Each_pick_shortens_the_other_lists_and_giving_it_back_restores_it()
    {
        Assert.Equal(Views.Length, For(1, ("Top", "Reset"), ("Bottom", "")).Count);
        Assert.Equal(Views.Length - 1, For(1, ("Top", "Reset"), ("Bottom", ""), ("Clock", "Open128")).Count);
        Assert.Equal(Views.Length + 1, For(1, ("Top", ""), ("Bottom", "")).Count);
    }

    [Fact]
    public void An_images_own_name_is_always_offered_to_it_in_the_projects_order()
    {
        IReadOnlyList<SchematicCadNameChoice> choices = For(0, ("C128 Neo.3.F", "Open128 - PCB Top"), ("Bottom scan", "Reset"));

        Assert.Equal(["", "Open128 - PCB Top", "Open128 - PCB Bottom", "Open128"], choices.Select(choice => choice.CadName));
        Assert.Equal("Open128 - PCB Top", choices[1].Label);
    }

    // Two images given the same name in Excel: nothing here takes it from either of them.
    [Fact]
    public void Two_images_sharing_a_name_each_keep_it()
    {
        (string, string)[] images = [("Scan A", "Reset"), ("Scan B", "Reset")];

        Assert.Contains(For(0, images), choice => choice.CadName == "Reset");
        Assert.Contains(For(1, images), choice => choice.CadName == "Reset");
    }

    // "Top (replica)" and page numbers resolve on the Schematics tab without being a view's name,
    // so a name that is no view is listed as it is - never dropped, never called wrong.
    [Fact]
    public void An_images_name_that_is_no_view_is_listed_last_as_it_is()
    {
        IReadOnlyList<SchematicCadNameChoice> choices = For(0, ("Top", "Top (replica)"));

        SchematicCadNameChoice last = choices[^1];
        Assert.Equal("Top (replica)", last.CadName);
        Assert.Equal("Top (replica)", last.Label);
        Assert.Equal(Views.Length + 2, choices.Count);
    }

    [Fact]
    public void Blank_and_repeated_view_names_are_listed_once()
    {
        IReadOnlyList<SchematicCadNameChoice> choices = SchematicCadNameChoices.For(
            [("Top", "")],
            0,
            ["Reset", " ", "reset", "", " Clock "]);

        Assert.Equal(["", "Reset", "Clock"], choices.Select(choice => choice.CadName));
    }

    [Fact]
    public void With_no_KiCad_data_only_No_CAD_name_and_the_images_own_name_are_offered()
    {
        Assert.Equal(
            ["", "Reset"],
            SchematicCadNameChoices.For([("Top", "Reset")], 0, []).Select(choice => choice.CadName));
        Assert.Equal(
            [""],
            SchematicCadNameChoices.For([("Top", "")], 0, []).Select(choice => choice.CadName));
    }

    [Fact]
    public void The_selected_choice_is_the_images_name_matched_without_case_or_No_CAD_name()
    {
        IReadOnlyList<SchematicCadNameChoice> choices = For(0, ("Top", " reset "));

        Assert.Equal("Reset", SchematicCadNameChoices.Selected(choices, " reset ").CadName);
        Assert.Equal(string.Empty, SchematicCadNameChoices.Selected(choices, "").CadName);
        Assert.Equal(string.Empty, SchematicCadNameChoices.Selected(choices, null).CadName);
    }

    [Fact]
    public void An_index_outside_the_images_is_refused()
    {
        Assert.Throws<ArgumentOutOfRangeException>(() => SchematicCadNameChoices.For([("Top", "")], 1, Views));
        Assert.Throws<ArgumentOutOfRangeException>(() => SchematicCadNameChoices.For([("Top", "")], -1, Views));
    }
}
