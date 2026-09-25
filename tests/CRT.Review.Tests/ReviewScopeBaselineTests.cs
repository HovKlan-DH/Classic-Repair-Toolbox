using CRT.Review.Handlers;

namespace CRT.Review.Tests;

// Covers ReviewScopeBaseline - the oscilloscope half of task 4's visual review.
//
// *** THE PICTURE ALONE IS NOT THE REVIEW, AND THAT IS THE WHOLE POINT OF THIS FILE. *** A scope
// baseline is already an image, so the side-by-side comparison shows the waveform. What it cannot
// show is that the SETTINGS changed - and a trace captured at 2 V/div next to one at 5 V/div looks
// like a different signal when it is the same signal drawn at a different scale.
//
// So a reviewer looking at two baselines needs to be told when the axes moved underneath them.
// Without it, the commonest way a baseline goes wrong - recapturing at different settings and not
// saying so - reads as a genuine change in the circuit.
public sealed class ReviewScopeBaselineTests
{
    private static ReviewSectionView Section(
        string rowKey,
        params ReviewFieldChangeView[] fields) =>
        new(
            ReviewScopeBaseline.SectionName,
            Added: [],
            Removed: [],
            Changed: [rowKey],
            Renamed: [],
            FieldChanges: new Dictionary<string, IReadOnlyList<ReviewFieldChangeView>>
            {
                [rowKey] = fields
            });

    [Fact]
    public void A_changed_VOLTS_PER_DIV_is_reported_as_a_settings_change()
    {
        // *** THE CASE THAT MATTERS MOST. *** The same waveform at a different vertical scale is a
        // different-looking picture. A reviewer comparing the two images without knowing this is
        // comparing apples with a scaled copy of apples.
        Assert.True(ReviewScopeBaseline.TryReadChange(
            ReviewScopeBaselineTests.Section("U8", new ReviewFieldChangeView("V/DIV", "2", "5")),
            "U8",
            out ReviewScopeSettingsChange change));

        Assert.Contains(change.Settings, setting =>
            setting.Field == "V/DIV" && setting.Before == "2" && setting.After == "5");
    }

    [Fact]
    public void A_changed_TIME_PER_DIV_is_reported_too()
    {
        // The horizontal twin: the same signal at a different sweep speed shows a different number
        // of cycles, which reads as a frequency change that did not happen.
        Assert.True(ReviewScopeBaseline.TryReadChange(
            ReviewScopeBaselineTests.Section("U8", new ReviewFieldChangeView("T/DIV", "1ms", "10us")),
            "U8",
            out ReviewScopeSettingsChange change));

        Assert.Single(change.Settings);
    }

    [Fact]
    public void A_changed_TRIGGER_LEVEL_is_reported()
    {
        Assert.True(ReviewScopeBaseline.TryReadChange(
            ReviewScopeBaselineTests.Section("U8", new ReviewFieldChangeView("T.LVL", "1.5", "2.5")),
            "U8",
            out ReviewScopeSettingsChange change));

        Assert.Single(change.Settings);
    }

    [Fact]
    public void ALL_THREE_changing_at_once_are_all_reported()
    {
        // A recapture on a differently-configured scope moves all three, and each is worth naming
        // - a reviewer told only "settings changed" still has to open the board to find out which.
        Assert.True(ReviewScopeBaseline.TryReadChange(
            ReviewScopeBaselineTests.Section(
                "U8",
                new ReviewFieldChangeView("T/DIV", "1ms", "10us"),
                new ReviewFieldChangeView("V/DIV", "2", "5"),
                new ReviewFieldChangeView("T.LVL", "1.5", "2.5")),
            "U8",
            out ReviewScopeSettingsChange change));

        Assert.Equal(3, change.Settings.Count);
    }

    [Fact]
    public void A_row_whose_SETTINGS_did_not_change_is_NOT_a_settings_change()
    {
        // *** THE ANTI-VACUITY CASE. *** A component image row changes for all sorts of reasons -
        // a corrected note, a renamed capture. Reporting those as a settings change would put a
        // warning on a submission that did not touch the scope at all, and a warning that fires on
        // everything is one a reviewer learns to ignore.
        Assert.False(ReviewScopeBaseline.TryReadChange(
            ReviewScopeBaselineTests.Section("U8", new ReviewFieldChangeView("Note", "a", "b")),
            "U8",
            out _));
    }

    [Fact]
    public void A_CLEARED_setting_is_reported_rather_than_passing_as_unchanged()
    {
        // Deleting the settings does not delete the picture, so the baseline stays on the board
        // with nothing recording how it was taken - which makes it unreproducible. That is a
        // deletion of information and is exactly what a reviewer should catch.
        Assert.True(ReviewScopeBaseline.TryReadChange(
            ReviewScopeBaselineTests.Section("U8", new ReviewFieldChangeView("V/DIV", "2", "")),
            "U8",
            out ReviewScopeSettingsChange change));

        Assert.Equal(string.Empty, Assert.Single(change.Settings).After);
    }

    [Fact]
    public void A_key_the_section_does_not_carry_is_refused()
    {
        Assert.False(ReviewScopeBaseline.TryReadChange(
            ReviewScopeBaselineTests.Section("U8", new ReviewFieldChangeView("V/DIV", "2", "5")),
            "U9",
            out _));
    }

    [Fact]
    public void A_null_section_is_refused_rather_than_throwing()
    {
        // Called while drawing a panel, from a submission whose payload may not have loaded.
        Assert.False(ReviewScopeBaseline.TryReadChange(null, "U8", out _));
    }

    [Fact]
    public void The_section_name_matches_what_the_SERVER_calls_it()
    {
        // A contract with BoardWorkbookSchema across a process boundary. If these drift the app
        // finds no component-image section and silently reports no settings changes at all - the
        // screen looks fine and omits the warning.
        Assert.Equal(
            global::Handlers.DataHandling.BoardWorkbookSchema.SheetComponentImages,
            ReviewScopeBaseline.SectionName);
    }

    [Fact]
    public void The_FIELD_NAMES_match_the_workbook_columns_the_server_reports()
    {
        // The three settings are matched by the workbook's own column headers, which is what
        // ReviewSummary puts in a field diff. A typo here means that setting is never reported and
        // nothing indicates it was missed.
        Assert.Equal(
            global::Handlers.DataHandling.BoardWorkbookSchema.ColTimeDiv,
            ReviewScopeBaseline.FieldTimeDiv);

        Assert.Equal(
            global::Handlers.DataHandling.BoardWorkbookSchema.ColVoltsDiv,
            ReviewScopeBaseline.FieldVoltsDiv);

        Assert.Equal(
            global::Handlers.DataHandling.BoardWorkbookSchema.ColTriggerLevelVolts,
            ReviewScopeBaseline.FieldTriggerLevel);
    }

    // -----------------------------------------------------------------------------------------
    // How the warning READS.
    // -----------------------------------------------------------------------------------------

    [Fact]
    public void The_warning_says_the_two_traces_are_NOT_directly_comparable()
    {
        // The reviewer's takeaway has to be actionable. "Settings changed" is a fact; "these two
        // are not drawn at the same scale" is the consequence they need in order to judge the
        // pictures above it correctly.
        string described = ReviewScopeBaseline.Describe(new ReviewScopeSettingsChange(
            [new ReviewScopeSetting("V/DIV", "2", "5")]));

        Assert.Contains("V/DIV", described);
        Assert.Contains("2", described);
        Assert.Contains("5", described);
        Assert.Contains("not", described, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void A_CLEARED_setting_reads_as_blank_rather_than_trailing_off()
    {
        // The same rule the field-level diff already follows: an empty right-hand side reads as
        // the app having failed to draw something.
        string described = ReviewScopeBaseline.Describe(new ReviewScopeSettingsChange(
            [new ReviewScopeSetting("V/DIV", "2", "")]));

        Assert.Contains("(blank)", described);
    }

    [Fact]
    public void A_change_with_NO_settings_is_refused_rather_than_described_emptily()
    {
        Assert.Throws<ArgumentException>(() =>
            ReviewScopeBaseline.Describe(new ReviewScopeSettingsChange([])));
    }

    // -----------------------------------------------------------------------------------------
    // Making a natural key readable.
    // -----------------------------------------------------------------------------------------

    [Fact]
    public void A_row_key_is_spelled_out_with_a_VISIBLE_separator()
    {
        // *** THE RAW KEY IS UNREADABLE ON SCREEN. *** BoardDraftNaturalKeys joins with U+241F,
        // which renders as a box or as nothing depending on the font - so a reviewer sees
        // "U8ASSY 250407 12Clock" or worse. The parts are genuinely useful (which component, which
        // region, which pin), so they are separated visibly rather than hidden.
        string key = global::Handlers.DataHandling.BoardDraftNaturalKeys
            .ForComponentImage("U8", "ASSY 250407", "12", "Clock");

        string readable = ReviewScopeBaseline.DescribeRowKey(key);

        Assert.Equal("U8 / ASSY 250407 / 12 / Clock", readable);
        Assert.DoesNotContain(
            global::Handlers.DataHandling.BoardDraftNaturalKeys.Separator,
            readable,
            StringComparison.Ordinal);
    }

    [Fact]
    public void An_EMPTY_part_is_dropped_rather_than_left_as_a_gap()
    {
        // A component image with no region carries an empty segment, and "U8 /  / 12" reads as
        // missing data rather than as a field that does not apply.
        string key = global::Handlers.DataHandling.BoardDraftNaturalKeys
            .ForComponentImage("U8", "", "12", "Clock");

        Assert.Equal("U8 / 12 / Clock", ReviewScopeBaseline.DescribeRowKey(key));
    }

    [Fact]
    public void A_blank_row_key_describes_as_nothing_rather_than_throwing()
    {
        Assert.Equal(string.Empty, ReviewScopeBaseline.DescribeRowKey(null));
        Assert.Equal(string.Empty, ReviewScopeBaseline.DescribeRowKey(""));
    }
}
