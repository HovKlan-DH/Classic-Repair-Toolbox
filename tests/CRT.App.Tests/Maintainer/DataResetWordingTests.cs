using Handlers.DataHandling;
using Handlers.MaintainerHandling;

namespace ClassicRepairToolbox.Tests.Maintainer;

// ###########################################################################################
// "Account" > Reset contribution data (owner request, 2026-10-04: "When I go-live
// with this, it should not have old data visible ... the real sources of BETA and stable must not be
// touched, but all contributor and maintainer data should go away").
//
// What matters most is CanReset: a reset cannot be undone, so the button works only while the server
// allows it, on counts the server can hold it to, and with the word typed exactly.
// ###########################################################################################
public sealed class DataResetWordingTests
{
    private static DataResetPlanAnswer Plan(bool enabled = true, string fingerprint = "f") =>
        new(enabled, fingerprint, 14, 6, 1, 4, 2, 7, 310, 1200, 80, enabled ? null : "Switched off.");

    private static string Text(IReadOnlyList<ReviewNoteRun> line) => string.Concat(line.Select(run => run.Text));

    [Fact]
    public void The_button_works_only_with_the_word_typed_exactly_on_counts_a_server_that_allows_it_sent()
    {
        Assert.True(DataResetWording.CanReset(DataResetWordingTests.Plan(), "RESET"));
        Assert.True(DataResetWording.CanReset(DataResetWordingTests.Plan(), "  RESET "));

        // Not quite the word.
        Assert.False(DataResetWording.CanReset(DataResetWordingTests.Plan(), "reset"));
        Assert.False(DataResetWording.CanReset(DataResetWordingTests.Plan(), "RESE"));
        Assert.False(DataResetWording.CanReset(DataResetWordingTests.Plan(), string.Empty));
        Assert.False(DataResetWording.CanReset(DataResetWordingTests.Plan(), null));

        // The word, but nothing the server would accept it for.
        Assert.False(DataResetWording.CanReset(DataResetWordingTests.Plan(enabled: false), "RESET"));
        Assert.False(DataResetWording.CanReset(DataResetWordingTests.Plan(fingerprint: " "), "RESET"));
        Assert.False(DataResetWording.CanReset(null, "RESET"));
    }

    [Fact]
    public void The_prompt_names_the_word_to_type()
    {
        Assert.Contains(DataResetWording.ConfirmWord, DataResetWording.ConfirmPrompt, StringComparison.Ordinal);
    }

    // ###########################################################################################
    // One line per kind of thing, each count bold and nothing else - and the administrators named as
    // KEPT, since "every account" alone would read as signing everybody out for good.
    // ###########################################################################################
    [Fact]
    public void Each_count_is_its_own_line_with_only_the_number_bold()
    {
        IReadOnlyList<IReadOnlyList<ReviewNoteRun>> lines = DataResetWording.Lines(DataResetWordingTests.Plan());

        Assert.Equal(
            [
                "[14] submissions - every one, whatever its state, with its uploaded files",
                "[6] accounts - every one but the administrators (1 administrator kept)",
                "[4] maintainers of a system",
                "[2] invitations to maintain",
                "[7] system records - made again as each system is next used; the systems themselves stay",
                "[310] history entries - the reset itself becomes the first",
                "[1,200] board views",
                "[80] API usage rows"
            ],
            lines.Select(DataResetWordingTests.Text));

        Assert.All(lines, line => Assert.Equal(line.Single(run => run.IsCount).Text, line[1].Text));
        Assert.Equal(["14", "6", "4", "2", "7", "310", "1,200", "80"], lines.Select(line => line.Single(run => run.IsCount).Text));
    }

    [Fact]
    public void One_of_a_kind_is_said_in_the_singular()
    {
        IReadOnlyList<IReadOnlyList<ReviewNoteRun>> lines = DataResetWording.Lines(new DataResetPlanAnswer(true, "f", 1, 1, 2, 1, 1, 1, 1, 1, 1));

        Assert.Equal("[1] submission - every one, whatever its state, with its uploaded files", DataResetWordingTests.Text(lines[0]));
        Assert.Equal("[1] account - every one but the administrators (2 administrators kept)", DataResetWordingTests.Text(lines[1]));
        Assert.Equal("[1] history entry - the reset itself becomes the first", DataResetWordingTests.Text(lines[5]));
    }

    [Fact]
    public void What_the_reset_deleted_is_one_sentence()
    {
        IReadOnlyList<ReviewNoteRun> done = DataResetWording.Done(new DataResetAnswer(14, 6, 4, 2, 7, 310, 1200, 80, 23));

        Assert.Equal(
            "The contribution data was reset: [14] submissions, [6] accounts, [4] maintainers, [2] invitations, " +
            "[7] system records, [310] history entries, [1,200] board views and [80] API usage rows deleted, " +
            "and [23] stored files removed from the disk.",
            DataResetWordingTests.Text(done));
    }

    // The reset never claims to have touched the data trees - the words say they stay.
    [Fact]
    public void The_words_say_the_BETA_and_stable_data_are_not_touched()
    {
        Assert.Contains("BETA and stable data are not touched", DataResetWording.Explanation, StringComparison.Ordinal);
        Assert.StartsWith("Not touched: the BETA and stable data", DataResetWording.Kept, StringComparison.Ordinal);
    }
}
