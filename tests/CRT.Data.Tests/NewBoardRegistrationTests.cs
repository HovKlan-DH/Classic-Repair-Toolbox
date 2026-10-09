using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// Tests for NewBoardRegistration.NotesForSubmission - a new board's notes from "Create board"
// as a submission carries them to the server, where the maintainer's placement starts with them and
// the main Excel data file's notes column receives them (owner request, 2026-10-05).
// ###########################################################################################
public sealed class NewBoardRegistrationTests
{
    [Theory]
    [InlineData("  Open-source replica.  ", "Open-source replica.")]
    [InlineData("Line one.\nLine two.", "Line one.\nLine two.")]
    [InlineData("   ", "")]
    [InlineData("", "")]
    public void The_notes_are_sent_trimmed(string typed, string sent)
    {
        Assert.Equal(sent, new NewBoardRegistration { HardwareNotes = typed }.NotesForSubmission());
    }

    // ###########################################################################################
    // "Create board" stops typing at the notes column's limit, but a draft made before it did, or a
    // marker edited by hand, can hold more - and CRT has nowhere to shorten them afterwards. Cut to
    // the column rather than refused, which would leave the draft unsendable for good.
    // ###########################################################################################
    [Fact]
    public void Notes_longer_than_the_notes_column_holds_are_cut_to_it()
    {
        string typed = new string('x', MasterListing.MaximumNotesLength) + " and more";

        Assert.Equal(
            new string('x', MasterListing.MaximumNotesLength),
            new NewBoardRegistration { HardwareNotes = typed }.NotesForSubmission());
    }

    // A marker read from disk can carry null where the property promised a string.
    [Fact]
    public void Missing_notes_are_none()
    {
        Assert.Equal(string.Empty, new NewBoardRegistration { HardwareNotes = null! }.NotesForSubmission());
    }
}
