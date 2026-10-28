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
    public void The_question_names_the_image_and_says_what_goes_and_that_Undo_brings_it_back()
    {
        Assert.Equal("Remove [PCB; Top] from your draft?", SchematicRemoval.Question("PCB; Top"));

        Assert.Equal(
            "Its highlights for 37 components and its KiCad calibration go with it. Undo at the top of the Contribute tab brings them back.",
            SchematicRemoval.WhatGoes(new SchematicRemovalLoss(37, true)));

        Assert.StartsWith("Its highlights for 1 component go with it.", SchematicRemoval.WhatGoes(new SchematicRemovalLoss(1, false)));
        Assert.StartsWith("Its KiCad calibration goes with it.", SchematicRemoval.WhatGoes(new SchematicRemovalLoss(0, true)));
        Assert.Contains("brings it back", SchematicRemoval.WhatGoes(new SchematicRemovalLoss(0, true)));
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
}
