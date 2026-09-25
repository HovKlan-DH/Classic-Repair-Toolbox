using Handlers.DataHandling;

namespace CRT.Data.Tests;

// ###########################################################################################
// Covers the KiCad CALIBRATION section of the review summary (Phase 5, tasks 3 and 6).
//
// *** THIS EXISTS SO CALIBRATIONS ARE NOT PUBLISHED UNREVIEWED. *** Calibrations began travelling
// with submissions on 2026-09-22. `ReviewSummary` compares two BoardData, and BoardData has no
// calibration section at all - so without this they would have reached the published tree with no
// reviewer ever having seen them. That is precisely the property this whole phase exists to
// prevent, and wiring the submission side without this would have created it.
//
// A calibration is the offset/scale/mirror box mapping a schematic image onto the KiCad board's
// coordinate space. Getting one wrong does not break anything visibly - the trace overlay simply
// lands in the wrong place - so it is exactly the kind of change a reviewer has to be TOLD about
// rather than expected to notice.
public sealed class ReviewCalibrationSectionTests
{
    private static KiCadCalibrationEntry Calibration(
        string schematic, double scaleX = 0.75, bool mirrorX = false) => new()
        {
            SchematicName = schematic,
            CadName = "board.kicad_pcb",
            OffsetX = 1.5,
            OffsetY = 2.5,
            ScaleX = scaleX,
            ScaleY = 0.8,
            MirrorX = mirrorX,
            MirrorY = false
        };

    private static ReviewSectionChange SectionOf(ReviewChangeSummary summary) =>
        summary.Sections.First(section => section.Section == ReviewSummary.SectionKiCadCalibrations);

    [Fact]
    public void An_ADDED_calibration_is_reported()
    {
        ReviewChangeSummary summary = ReviewSummary.Compare(
            new BoardData(),
            new BoardData(),
            renames: null,
            publishedCalibrations: [],
            submittedCalibrations: [ReviewCalibrationSectionTests.Calibration("Sheet 1")]);

        Assert.Equal("Sheet 1", Assert.Single(ReviewCalibrationSectionTests.SectionOf(summary).Added));
    }

    [Fact]
    public void A_REMOVED_calibration_is_reported()
    {
        // The least recoverable direction: a board that HAD a working trace overlay loses it, and
        // nothing on the schematic looks different.
        ReviewChangeSummary summary = ReviewSummary.Compare(
            new BoardData(),
            new BoardData(),
            renames: null,
            publishedCalibrations: [ReviewCalibrationSectionTests.Calibration("Sheet 1")],
            submittedCalibrations: []);

        Assert.Equal("Sheet 1", Assert.Single(ReviewCalibrationSectionTests.SectionOf(summary).Removed));
    }

    [Fact]
    public void A_CHANGED_calibration_is_reported_with_the_field_that_moved()
    {
        // *** THE CASE A REVIEWER CANNOT OTHERWISE SEE. *** The schematic looks identical; only
        // the numbers behind the overlay moved. "Sheet 1 changed" is not reviewable, but
        // "ScaleX: 0.75 -> 0.9" is.
        ReviewChangeSummary summary = ReviewSummary.Compare(
            new BoardData(),
            new BoardData(),
            renames: null,
            publishedCalibrations: [ReviewCalibrationSectionTests.Calibration("Sheet 1", scaleX: 0.75)],
            submittedCalibrations: [ReviewCalibrationSectionTests.Calibration("Sheet 1", scaleX: 0.9)]);

        ReviewSectionChange section = ReviewCalibrationSectionTests.SectionOf(summary);

        Assert.Equal("Sheet 1", Assert.Single(section.Changed));

        Assert.Contains(
            section.FieldChanges["Sheet 1"],
            field => field.Field == "ScaleX" && field.Before == "0.75" && field.After == "0.9");
    }

    [Fact]
    public void A_flipped_MIRROR_is_reported()
    {
        // Mirroring is the change most likely to be a mistake and least likely to be noticed in a
        // number - it puts the whole overlay on the wrong side of the board.
        ReviewChangeSummary summary = ReviewSummary.Compare(
            new BoardData(),
            new BoardData(),
            renames: null,
            publishedCalibrations: [ReviewCalibrationSectionTests.Calibration("Sheet 1", mirrorX: false)],
            submittedCalibrations: [ReviewCalibrationSectionTests.Calibration("Sheet 1", mirrorX: true)]);

        Assert.Contains(
            ReviewCalibrationSectionTests.SectionOf(summary).FieldChanges["Sheet 1"],
            field => field.Field == "MirrorX");
    }

    [Fact]
    public void An_UNCHANGED_calibration_is_not_reported()
    {
        // The anti-vacuity half. Submissions carry the complete state, so an untouched board
        // re-sends its calibration every time; reporting that as a change would put a line on
        // every review that means nothing.
        ReviewChangeSummary summary = ReviewSummary.Compare(
            new BoardData(),
            new BoardData(),
            renames: null,
            publishedCalibrations: [ReviewCalibrationSectionTests.Calibration("Sheet 1")],
            submittedCalibrations: [ReviewCalibrationSectionTests.Calibration("Sheet 1")]);

        Assert.False(ReviewCalibrationSectionTests.SectionOf(summary).HasChanges);
    }

    [Fact]
    public void Numbers_are_compared_INVARIANT_so_the_server_agrees_with_itself()
    {
        // The values are doubles here rather than the strings BoardData uses, so they are
        // FORMATTED for display - and a culture-sensitive format would render 0.75 as "0,75" on a
        // Danish server, reporting a change on every field of every calibration forever.
        System.Globalization.CultureInfo original = System.Globalization.CultureInfo.CurrentCulture;

        try
        {
            System.Globalization.CultureInfo.CurrentCulture = new System.Globalization.CultureInfo("da-DK");

            ReviewChangeSummary summary = ReviewSummary.Compare(
                new BoardData(),
                new BoardData(),
                renames: null,
                publishedCalibrations: [ReviewCalibrationSectionTests.Calibration("Sheet 1", scaleX: 0.75)],
                submittedCalibrations: [ReviewCalibrationSectionTests.Calibration("Sheet 1", scaleX: 0.9)]);

            Assert.Contains(
                ReviewCalibrationSectionTests.SectionOf(summary).FieldChanges["Sheet 1"],
                field => field.Before == "0.75");
        }
        finally
        {
            System.Globalization.CultureInfo.CurrentCulture = original;
        }
    }

    [Fact]
    public void The_section_appears_even_when_NOTHING_is_calibrated()
    {
        // The summary always carries all of its sections; the client drops the empty ones. A
        // section that vanished entirely would make the client's filter the only thing standing
        // between a reviewer and a missing category.
        ReviewChangeSummary summary = ReviewSummary.Compare(new BoardData(), new BoardData());

        Assert.Contains(summary.Sections, section =>
            section.Section == ReviewSummary.SectionKiCadCalibrations);
    }

    [Fact]
    public void Calibrations_default_to_EMPTY_so_existing_callers_are_unaffected()
    {
        // Both parameters are optional and trailing. A caller that passes neither reports no
        // calibration changes, which is exactly right for one that does not know about them.
        ReviewChangeSummary summary = ReviewSummary.Compare(new BoardData(), new BoardData());

        Assert.False(ReviewCalibrationSectionTests.SectionOf(summary).HasChanges);
    }

    [Fact]
    public void A_NEW_SYSTEM_counts_its_calibrations_as_additions()
    {
        ReviewChangeSummary summary = ReviewSummary.Compare(
            published: null,
            new BoardData(),
            renames: null,
            publishedCalibrations: null,
            submittedCalibrations: [ReviewCalibrationSectionTests.Calibration("Sheet 1")]);

        Assert.True(summary.IsNewSystem);
        Assert.Single(ReviewCalibrationSectionTests.SectionOf(summary).Added);
    }
}
