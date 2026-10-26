using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// SubmitDraftsWording - what the Submit dialog says when it sends several drafts (CRT 4.0 step 2,
// slice 5): why a draft is not offered, each tick box, and what came of each board.
// ###########################################################################################
public sealed class SubmitDraftsWordingTests
{
    // In the order the row's own Submit reasons go: nothing in it, already sent, then errors.
    [Fact]
    public void A_draft_is_offered_unless_it_holds_nothing_was_just_sent_or_has_errors()
    {
        Assert.Null(SubmitDraftsWording.NotOffered(hasSomething: true, alreadySent: false, errors: 0));
        Assert.Equal("nothing to send yet", SubmitDraftsWording.NotOffered(hasSomething: false, alreadySent: true, errors: 3));
        Assert.Equal("already sent as it is", SubmitDraftsWording.NotOffered(hasSomething: true, alreadySent: true, errors: 3));
        Assert.Equal("1 error to fix first", SubmitDraftsWording.NotOffered(hasSomething: true, alreadySent: false, errors: 1));
        Assert.Equal("3 errors to fix first", SubmitDraftsWording.NotOffered(hasSomething: true, alreadySent: false, errors: 3));
    }

    [Fact]
    public void A_tick_box_names_the_board_and_what_its_draft_holds_or_why_it_cannot_go()
    {
        Assert.Equal("Commodore 64 - 250425 (12 rows changed)", SubmitDraftsWording.Choice("Commodore 64 - 250425", "12 rows changed", null));
        Assert.Equal("Commodore 64 - 250425 - already sent as it is", SubmitDraftsWording.Choice("Commodore 64 - 250425", "12 rows changed", SubmitDraftsWording.AlreadySent));
    }

    // A refusal's sentence as the tail of "Board: not sent - ...": no full stop, a small first letter -
    // but an abbreviation keeps its capitals, and nothing at all still says something.
    [Fact]
    public void A_sentence_becomes_a_reason()
    {
        Assert.Equal("some files are missing", SubmitDraftsWording.AsReason(" Some files are missing. "));
        Assert.Equal("KiCad data is missing", SubmitDraftsWording.AsReason("KiCad data is missing."));
        Assert.Equal(SubmitDraftsWording.Unexpected, SubmitDraftsWording.AsReason("  "));
        Assert.Equal(SubmitDraftsWording.Unexpected, SubmitDraftsWording.AsReason(null));
    }

    [Fact]
    public void The_outcome_says_how_many_went_and_what_came_of_each_board()
    {
        Assert.Equal("Contributions sent", SubmitDraftsWording.Heading(2, 2));
        Assert.Equal("1 of 3 contributions sent", SubmitDraftsWording.Heading(1, 3));
        Assert.Equal("Contributions not sent", SubmitDraftsWording.Heading(0, 2));

        Assert.Equal("Nothing was sent. Your drafts are untouched.", SubmitDraftsWording.Intro(0));
        Assert.Contains("queued for review on its own", SubmitDraftsWording.Intro(2), StringComparison.Ordinal);

        Assert.Equal("A: queued for review", SubmitDraftsWording.Queued("A"));
        Assert.Equal("A: not accepted - the server found problems to fix first", SubmitDraftsWording.NotAccepted("A"));
        Assert.Equal("A: not sent - its draft could not be read", SubmitDraftsWording.NotSent("A", SubmitDraftsWording.NotRead));
        Assert.Equal("A: Uploading", SubmitDraftsWording.Sending("A", "Uploading"));
    }

    // Plain ASCII, no Excel, and no "their" for a contributor (CLAUDE.md).
    [Fact]
    public void Every_sentence_is_plain_ASCII_and_names_no_contributor_by_pronoun()
    {
        Assert.All(SubmitDraftsWording.All(), sentence =>
        {
            Assert.All(sentence, character => Assert.True(character < 128, $"[{sentence}] is not plain ASCII"));
            Assert.DoesNotContain("Excel", sentence, StringComparison.Ordinal);
            Assert.DoesNotMatch(@"\b(their|them|they)\b", sentence);
        });
    }
}
