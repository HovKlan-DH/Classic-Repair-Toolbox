using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// What each editor's change is called in a draft's history (DraftStepWording, CRT 4.0 step 1,
// 2026-10-25) - "Undone: Changes to [U8]." on the table's status line, the draft's Undo later. The
// component, schematic or count it names is what tells two steps apart.
// ###########################################################################################
public sealed class DraftStepWordingTests
{
    [Fact]
    public void Each_editors_step_names_what_it_changed()
    {
        Assert.Equal("Changes to [U8]", DraftStepWording.ComponentSaved("U8"));
        Assert.Equal("[U8] deleted", DraftStepWording.ComponentDeleted("U8"));
        Assert.Equal("[U8] back to its published values", DraftStepWording.ComponentReverted("U8"));
        Assert.Equal("Components marked on [Top]", DraftStepWording.ComponentsMarked("Top"));
        Assert.Equal("KiCad calibration of [Top]", DraftStepWording.Calibration("Top"));
        Assert.Equal("Schematic image [Top] removed", DraftStepWording.SchematicRemoved("Top"));
        Assert.Equal("Schematic image [Top] renamed [Front]", DraftStepWording.SchematicRenamed("Top", "Front"));
        Assert.Equal("CAD name of [Top]", DraftStepWording.CadName("Top"));

        // The component pane's own sections (CRT 4.0 step 2).
        Assert.Equal("Pictures of [U8]", DraftStepWording.ComponentPictures("U8"));
        Assert.Equal("Scope baselines of [U8]", DraftStepWording.ComponentBaselines("U8"));
        Assert.Equal("Files of [U8]", DraftStepWording.ComponentFiles("U8"));
        Assert.Equal("[U9] added", DraftStepWording.ComponentAdded("U9"));
        Assert.Equal("[U7] copied from Commodore 64 - 250425", DraftStepWording.ComponentCopied("U7", "Commodore 64 - 250425"));
        Assert.Equal("Pins of [U8] from KiCad", DraftStepWording.ComponentPinsFromKiCad("U8"));
        Assert.Equal("[Datasheet] copied from Commodore 64 - 250407", DraftStepWording.FilesCopied(["Datasheet"], "Commodore 64 - 250407"));
        Assert.Equal("Files copied from Commodore 64 - 250407", DraftStepWording.FilesCopied(["Datasheet", "Pinout"], "Commodore 64 - 250407"));

        // The board's own lists (slice 4a).
        Assert.Equal("The board's files", DraftStepWording.BoardFiles);
        Assert.Equal("The board's links", DraftStepWording.BoardLinks);
        Assert.Equal("Credits", DraftStepWording.Credits);
        Assert.Equal("Important signals", DraftStepWording.ImportantSignals);
        Assert.Equal("Colours of the pictures", DraftStepWording.PictureColours);
    }

    [Fact]
    public void Images_added_are_counted_and_read_in_the_singular_for_one()
    {
        Assert.Equal("Schematic image added", DraftStepWording.SchematicsAdded(1));
        Assert.Equal("3 schematic images added", DraftStepWording.SchematicsAdded(3));
    }

    // A step's name is read by a hobbyist on the status line: plain ASCII, no program named.
    [Fact]
    public void Every_step_name_is_plain_ASCII_and_names_no_program()
    {
        string[] names =
        [
            DraftStepWording.ComponentSaved("U8"),
            DraftStepWording.ComponentDeleted("U8"),
            DraftStepWording.ComponentAdded("U8"),
            DraftStepWording.ComponentReverted("U8"),
            DraftStepWording.ComponentsMarked("Top"),
            DraftStepWording.Calibration("Top"),
            DraftStepWording.SchematicsAdded(2),
            DraftStepWording.SchematicRemoved("Top"),
            DraftStepWording.SchematicRenamed("Top", "Front"),
            DraftStepWording.SchematicOrder,
            DraftStepWording.CadName("Top"),
            DraftStepWording.KiCadFileRemoved,
            DraftStepWording.BroughtUpToDate,
            DraftStepWording.BoardFiles,
            DraftStepWording.BoardLinks,
            DraftStepWording.Credits,
            DraftStepWording.ImportantSignals,
            DraftStepWording.PictureColours,
        ];

        Assert.All(names, name =>
        {
            Assert.All(name, character => Assert.True(character < 128, $"[{name}] is not plain ASCII"));
            Assert.DoesNotContain("Excel", name, StringComparison.Ordinal);
        });
    }
}
