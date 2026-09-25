using Handlers.DataHandling;

namespace CRT.Data.Tests;

// ###########################################################################################
// Covers PublishMerge - turning a submission into the board that will be written
// (NewContributeStrategy.md Phase 5, task 6: "the caller that assembles the merged BoardData").
//
// *** "MERGE" IS THE WRONG WORD FOR WHAT THIS DOES, and getting that backwards would corrupt
// boards. *** The manifest is BY CONTRACT the complete intended state - "a file present in the
// base revision but ABSENT here is a deletion". So the submitted rows REPLACE the published ones
// wholesale; they are not overlaid onto them. Anything that tried to combine the two would
// resurrect every row a contributor deliberately deleted, and nothing would ever report it.
//
// What this class genuinely has to decide is the handful of things the manifest does NOT carry:
// the revision date, and the KiCad calibrations that live outside BoardData entirely.
public sealed class PublishMergeTests
{
    private static SubmissionManifest Manifest(string? revisionDate = null)
    {
        var manifest = new SubmissionManifest
        {
            SystemId = "Commodore/C64/250407",
            Manufacturer = "Commodore",
            Hardware = "C64",
            Board = "250407"
        };

        manifest.Rows.Components.Add(new ComponentEntry
        {
            BoardLabel = "U8",
            FriendlyName = "PLA",
            TechnicalNameOrValue = "906114-01"
        });

        if (revisionDate is not null)
            manifest.Rows.RevisionDate = revisionDate;

        return manifest;
    }

    private static BoardData Published(string revisionDate = "2026-August-21") => new()
    {
        RevisionDate = revisionDate,
        Components =
        [
            new ComponentEntry { BoardLabel = "U1", FriendlyName = "OLD", TechnicalNameOrValue = "x" },
            new ComponentEntry { BoardLabel = "U2", FriendlyName = "ALSO OLD", TechnicalNameOrValue = "y" }
        ]
    };

    // -----------------------------------------------------------------------------------------
    // Replacement, not overlay.
    // -----------------------------------------------------------------------------------------

    [Fact]
    public void The_submitted_rows_REPLACE_the_published_ones()
    {
        // *** THE CASE THAT DEFINES THIS CLASS. *** The published board has U1 and U2; the
        // submission carries only U8. The result is U8 alone - the other two were DELETED by the
        // contributor, and an overlay would silently put them back.
        BoardData merged = PublishMerge.Build(
            PublishMergeTests.Manifest(), PublishMergeTests.Published());

        ComponentEntry component = Assert.Single(merged.Components);

        Assert.Equal("U8", component.BoardLabel);
    }

    [Fact]
    public void A_NEW_SYSTEM_needs_no_published_board_at_all()
    {
        // The highest-risk submission there is, and the one where a null published board is the
        // normal case rather than an error.
        BoardData merged = PublishMerge.Build(PublishMergeTests.Manifest(), published: null);

        Assert.Equal("U8", Assert.Single(merged.Components).BoardLabel);
    }

    [Fact]
    public void EVERY_board_section_travels()
    {
        // A section silently dropped here is data loss at publish time, exactly like the sidecar
        // gap was. Asserted section by section rather than trusted to a mapping nobody rereads.
        SubmissionManifest manifest = PublishMergeTests.Manifest();

        manifest.Rows.Schematics.Add(new BoardSchematicEntry { SchematicName = "Sheet 1" });
        manifest.Rows.ComponentImages.Add(new ComponentImageEntry { BoardLabel = "U8", Name = "cap" });
        manifest.Rows.ComponentHighlights.Add(new ComponentHighlightEntry { SchematicName = "Sheet 1", BoardLabel = "U8" });
        manifest.Rows.ComponentLocalFiles.Add(new ComponentLocalFileEntry { BoardLabel = "U8", Name = "ds.pdf" });
        manifest.Rows.ComponentLinks.Add(new ComponentLinkEntry { BoardLabel = "U8", Name = "link" });
        manifest.Rows.BoardLocalFiles.Add(new BoardLocalFileEntry { Category = "Docs", Name = "manual.pdf" });
        manifest.Rows.BoardLinks.Add(new BoardLinkEntry { Category = "Docs", Name = "site" });
        manifest.Rows.Credits.Add(new CreditEntry { Category = "Data", NameOrHandle = "Someone" });
        manifest.Rows.KiCadImportantSignals.Add(new KiCadImportantSignalEntry { DisplayName = "CLK" });

        BoardData merged = PublishMerge.Build(manifest, published: null);

        Assert.Single(merged.Schematics);
        Assert.Single(merged.Components);
        Assert.Single(merged.ComponentImages);
        Assert.Single(merged.ComponentHighlights);
        Assert.Single(merged.ComponentLocalFiles);
        Assert.Single(merged.ComponentLinks);
        Assert.Single(merged.BoardLocalFiles);
        Assert.Single(merged.BoardLinks);
        Assert.Single(merged.Credits);
        Assert.Single(merged.KiCadImportantSignals);
    }

    // -----------------------------------------------------------------------------------------
    // The revision date - the one field the rows do not reliably carry.
    // -----------------------------------------------------------------------------------------

    [Fact]
    public void The_SUBMITTED_revision_date_wins_when_there_is_one()
    {
        BoardData merged = PublishMerge.Build(
            PublishMergeTests.Manifest("2026-September-22"), PublishMergeTests.Published());

        Assert.Equal("2026-September-22", merged.RevisionDate);
    }

    [Fact]
    public void The_PUBLISHED_date_is_KEPT_when_the_submission_carries_none()
    {
        // *** NOT BLANKED. *** The revision date is what a contributor's draft records as its
        // BaseRevision, and it is what the drift warning compares against. Publishing an empty one
        // would make every existing draft look as though it were based on nothing, and the drift
        // check would have no answer to give.
        BoardData merged = PublishMerge.Build(
            PublishMergeTests.Manifest(), PublishMergeTests.Published("2026-August-21"));

        Assert.Equal("2026-August-21", merged.RevisionDate);
    }

    [Fact]
    public void A_NEW_SYSTEM_with_no_date_anywhere_gets_an_EMPTY_one_rather_than_a_fabricated_date()
    {
        // Inventing today's date would state, in a published file, that the board data was revised
        // on a day nobody revised it. Empty is honest and the reader already tolerates it.
        BoardData merged = PublishMerge.Build(PublishMergeTests.Manifest(), published: null);

        Assert.Equal(string.Empty, merged.RevisionDate);
    }

    // -----------------------------------------------------------------------------------------
    // The PUBLISHED REVISION - what `systems.current_revision` becomes.
    // -----------------------------------------------------------------------------------------

    [Fact]
    public void The_published_revision_IS_the_boards_own_revision_date()
    {
        // *** THIS IS A CONTRACT WITH THE CLIENT, not a free choice. *** DraftBaseRevision stamps
        // a draft's BaseRevision from the official board's REVISION DATE, and the submission sends
        // that value back. So `systems.current_revision` must be the same thing, or every
        // contributor would be diffing against a value their drafts never carry and the drift
        // warning would fire on every board forever.
        BoardData merged = PublishMerge.Build(
            PublishMergeTests.Manifest("2026-September-22"), PublishMergeTests.Published());

        Assert.Equal("2026-September-22", PublishMerge.RevisionOf(merged));
    }

    [Fact]
    public void A_board_with_no_revision_date_has_no_revision()
    {
        // Rather than a placeholder that would later be mistaken for a real one.
        Assert.Equal(string.Empty, PublishMerge.RevisionOf(new BoardData()));
    }

    // -----------------------------------------------------------------------------------------
    // Calibrations - outside BoardData entirely.
    // -----------------------------------------------------------------------------------------

    [Fact]
    public void CALIBRATIONS_come_from_the_manifest_and_are_NOT_part_of_the_board()
    {
        // BoardData has no calibration section - they live only in the JSON sidecar - so they are
        // read straight off the manifest and handed to the sidecar writer separately. A caller
        // that looked for them on BoardData would find nothing and publish a board with a
        // contributor's calibration work silently removed.
        SubmissionManifest manifest = PublishMergeTests.Manifest();

        manifest.Rows.KiCadCalibrations.Add(new KiCadCalibrationEntry
        {
            SchematicName = "Sheet 1",
            CadName = "board.kicad_pcb",
            ScaleX = 0.75
        });

        IReadOnlyList<KiCadCalibrationEntry> calibrations = PublishMerge.CalibrationsOf(manifest);

        Assert.Equal("Sheet 1", Assert.Single(calibrations).SchematicName);
    }

    [Fact]
    public void A_submission_with_NO_calibrations_yields_an_EMPTY_list_rather_than_null()
    {
        // PublishExecutor refuses null outright, so an empty list is what "this board has none"
        // has to look like.
        Assert.Empty(PublishMerge.CalibrationsOf(PublishMergeTests.Manifest()));
    }

    [Fact]
    public void A_NULL_manifest_is_refused()
    {
        Assert.Throws<ArgumentNullException>(() => PublishMerge.Build(null!, null));
        Assert.Throws<ArgumentNullException>(() => PublishMerge.CalibrationsOf(null!));
    }

    [Fact]
    public void The_merged_board_does_not_ALIAS_the_manifests_own_lists()
    {
        // The manifest is reloaded from the store per request and the merged board is handed to a
        // writer; sharing list instances means a later mutation of one silently changes the other.
        SubmissionManifest manifest = PublishMergeTests.Manifest();

        BoardData merged = PublishMerge.Build(manifest, published: null);

        manifest.Rows.Components.Add(new ComponentEntry { BoardLabel = "U99" });

        Assert.Single(merged.Components);
    }
}
