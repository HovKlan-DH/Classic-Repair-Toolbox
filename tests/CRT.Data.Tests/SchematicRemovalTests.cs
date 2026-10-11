using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// SchematicRemoval - what removing a schematic image from a draft takes with it, and what the
// question before it says (code review, 2026-10-10): the highlights drawn on the image and its
// KiCad calibration go with the row, and adding the image again brings none of them back, so
// "Remove" asks first when there is any.
// ###########################################################################################
public sealed class SchematicRemovalTests : IDisposable
{
    private readonly TempWorkspace thisWorkspace = new();

    public void Dispose() => this.thisWorkspace.Dispose();

    private static ComponentHighlightEntry Highlight(string schematicName, string label) =>
        new() { SchematicName = schematicName, BoardLabel = label, X = "1", Y = "2", Width = "3", Height = "4" };

    [Fact]
    public void The_components_marked_on_the_image_are_counted_once_each()
    {
        SchematicRemovalLoss loss = SchematicRemoval.LossOf(
            "Top",
            [Highlight("Top", "U1"), Highlight("top ", "U1"), Highlight("Top", "u1"), Highlight("Top", "U2"), Highlight("Top", " "), Highlight("Bottom", "U3")],
            hasCalibration: false);

        Assert.Equal(new SchematicRemovalLoss(2, false), loss);
        Assert.False(loss.IsNothing);
    }

    [Fact]
    public void An_image_with_nothing_drawn_on_it_loses_nothing()
    {
        Assert.True(SchematicRemoval.LossOf("Top", [Highlight("Bottom", "U1")], hasCalibration: false).IsNothing);
        Assert.Equal(string.Empty, SchematicRemoval.WhatGoes(SchematicRemovalLoss.Nothing));
    }

    [Fact]
    public void The_question_names_the_image_and_says_what_goes_and_that_it_does_not_come_back()
    {
        Assert.Equal("Remove [PCB; Top] from your draft?", SchematicRemoval.Question("PCB; Top"));

        Assert.Equal(
            "Its highlights for 37 components and its KiCad calibration go with it. Adding the image again later does not bring them back - the image file itself stays in your draft folder.",
            SchematicRemoval.WhatGoes(new SchematicRemovalLoss(37, true)));

        Assert.StartsWith("Its highlights for 1 component go with it.", SchematicRemoval.WhatGoes(new SchematicRemovalLoss(1, false)));
        Assert.StartsWith("Its KiCad calibration goes with it.", SchematicRemoval.WhatGoes(new SchematicRemovalLoss(0, true)));
        Assert.Contains("does not bring it back", SchematicRemoval.WhatGoes(new SchematicRemovalLoss(0, true)));
    }

    // Read from the sidecar beside the workbook - where the highlights and the calibration live.
    [Fact]
    public void What_goes_is_read_from_the_drafts_sidecar()
    {
        string workbook = Path.Combine(this.thisWorkspace.Root, "Data C64 250407.xlsx");

        Assert.Equal(SchematicRemovalLoss.Nothing, SchematicRemoval.Read(workbook, "Top"));

        BoardSidecarWriter.Write(
            workbook,
            [Highlight("Top", "U1"), Highlight("Top", "U2"), Highlight("Bottom", "U3")],
            [new KiCadCalibrationEntry { SchematicName = "Top", CadName = "Open128 - PCB Top" }],
            []);

        Assert.Equal(new SchematicRemovalLoss(2, true), SchematicRemoval.Read(workbook, "Top"));
        Assert.Equal(new SchematicRemovalLoss(1, false), SchematicRemoval.Read(workbook, "Bottom"));
    }

    // ###########################################################################################
    // *** A SAVE OF THE TABLE THAT REMOVES IMAGES ASKS THE SAME QUESTION (code review, 2026-10-11). ***
    // A row deleted - or renamed with its image file changed in the same save, which reads as one
    // image removed and another added - took what was drawn on it without a word. The question
    // names every image, says what goes, and how a rename keeps them.
    // ###########################################################################################
    [Fact]
    public void A_save_that_removes_one_image_says_what_goes_and_how_a_rename_keeps_it()
    {
        SchematicDrop[] drops = [new("PCB; Top", new SchematicRemovalLoss(3, true))];

        Assert.Equal("Save, and remove [PCB; Top] from your draft?", SchematicRemoval.SaveQuestion(drops));

        string whatGoes = SchematicRemoval.SaveWhatGoes(drops);
        Assert.StartsWith(SchematicRemoval.WhatGoes(new SchematicRemovalLoss(3, true)), whatGoes);
        Assert.Contains("Renamed it, not removed it?", whatGoes);
        Assert.Contains("change the name and the file in two separate saves", whatGoes);
        Assert.EndsWith("Cancel saves nothing, and Ctrl+Z in the table brings back a row you deleted.", whatGoes);
    }

    [Fact]
    public void A_save_that_removes_several_images_names_each_and_what_goes_with_it()
    {
        SchematicDrop[] drops =
        [
            new("Top", new SchematicRemovalLoss(2, false)),
            new("Bottom", new SchematicRemovalLoss(0, true)),
            new("Clock", new SchematicRemovalLoss(1, true)),
        ];

        Assert.Equal("Save, and remove [Top], [Bottom] and [Clock] from your draft?", SchematicRemoval.SaveQuestion(drops));
        Assert.Equal("Save, and remove [Top] and [Bottom] from your draft?", SchematicRemoval.SaveQuestion(drops[..2]));

        string whatGoes = SchematicRemoval.SaveWhatGoes(drops);
        Assert.StartsWith(
            "[Top]: Its highlights for 2 components go with it. [Bottom]: Its KiCad calibration goes with it. " +
            "[Clock]: Its highlights for 1 component and its KiCad calibration go with it. " +
            "Adding the images again later does not bring any of it back",
            whatGoes);
        Assert.Contains("Renamed them, not removed them?", whatGoes);
    }
}
