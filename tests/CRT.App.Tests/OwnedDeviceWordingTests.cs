using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// OwnedDeviceWording - what CRT says about the user's devices (CRT 4.0 step 4): "Device" where a
// person reads it, plain ASCII, hobbyist words.
// ###########################################################################################
public sealed class OwnedDeviceWordingTests
{
    [Fact]
    public void A_board_key_reads_as_hardware_and_board()
    {
        Assert.Equal("Commodore 64 - 250407", OwnedDeviceWording.BoardText("Commodore 64|250407"));
        Assert.Equal("no board", OwnedDeviceWording.BoardText("  "));
    }

    // Region and serial number only when the device has them.
    [Fact]
    public void The_detail_line_names_the_board_then_what_the_device_has()
    {
        var full = new OwnedDeviceRecord { Region = " PAL ", SerialNumber = "CA123456" };
        var bare = new OwnedDeviceRecord();

        Assert.Equal("Commodore 64 - 250407 - PAL - serial CA123456", OwnedDeviceWording.DetailLine(full, "Commodore 64 - 250407"));
        Assert.Equal("Commodore 64 - 250407", OwnedDeviceWording.DetailLine(bare, "Commodore 64 - 250407"));
    }

    // What stays is said where it stays - the workbook list (code review, 2026-10-28: it named a
    // "repairs without a device" list CRT does not have).
    [Fact]
    public void Deleting_says_what_goes_and_where_the_repairs_stay()
    {
        Assert.Equal(
            "Delete [Breadbin #2]? Its photos, notes and changes are deleted with it. Its 3 repairs stay in the workbook list, on no device.",
            OwnedDeviceWording.DeleteQuestion(" Breadbin #2 ", 3));
        Assert.EndsWith("It has no repairs.", OwnedDeviceWording.DeleteQuestion("Plus/4", 0), StringComparison.Ordinal);
        Assert.Contains("Its 1 repair stays in the workbook list", OwnedDeviceWording.DeleteQuestion("Plus/4", 1), StringComparison.Ordinal);
    }

    // "Add photos" says how many went in, and how many could not (code review, 2026-10-28).
    [Fact]
    public void Photos_added_are_said_with_those_that_could_not_be()
    {
        Assert.Equal("Added 2 photos.", OwnedDeviceWording.PhotosAdded(2, 0));
        Assert.StartsWith("Added 1 photo. 1 could not be added - it is", OwnedDeviceWording.PhotosAdded(1, 1), StringComparison.Ordinal);
        Assert.StartsWith("Added 0 photos. 3 could not be added - they are", OwnedDeviceWording.PhotosAdded(0, 3), StringComparison.Ordinal);
    }

    [Fact]
    public void Headings_count_what_they_hold()
    {
        Assert.Equal("Changes made to it", OwnedDeviceWording.ChangesHeading(0));
        Assert.Equal("Changes made to it (2)", OwnedDeviceWording.ChangesHeading(2));
        Assert.Equal("Notes and photos (1 photo)", OwnedDeviceWording.PhotosHeading(1));
    }

    // The plan's own words: "The bar reads 'Breadbin #2 - No video'".
    [Fact]
    public void The_bar_names_the_device_then_the_repair()
    {
        Assert.Equal("Breadbin #2 - No video", OwnedDeviceWording.BarText(" Breadbin #2 ", " No video "));
        Assert.Equal("Breadbin #2", OwnedDeviceWording.BarText("Breadbin #2", "  "));
        Assert.Equal("New repair on Breadbin #2", OwnedDeviceWording.NewRepairTitle(" Breadbin #2 "));
    }

    // An exported history's list of repairs: id, title, then in brackets its state, its date words
    // and its worklogs - bracketed, as a title may hold " - " itself.
    [Fact]
    public void A_repair_in_the_history_is_listed_by_what_it_is_and_how_far_it_got()
    {
        Assert.Equal("#12 No video - VIC socket (Open, started 2026-09-28, 3 worklogs)", OwnedDeviceWording.HistoryRepairLine(12, " No video - VIC socket ", "Open", "started 2026-09-28", 3));
        Assert.Equal("#4 (untitled) (Closed, ended 2026-10-04, 1 worklog)", OwnedDeviceWording.HistoryRepairLine(4, "", "Closed", "ended 2026-10-04", 1));
    }

    [Fact]
    public void Every_sentence_is_plain_ASCII_and_calls_it_a_device()
    {
        string[] sentences =
        [
            OwnedDeviceWording.ListHeading,
            OwnedDeviceWording.NoDevicesYet,
            OwnedDeviceWording.NameNeeded,
            OwnedDeviceWording.NoRepairsYet,
            OwnedDeviceWording.NoChangesYet,
            OwnedDeviceWording.NoPhotosYet,
            OwnedDeviceWording.RepairsHeading,
            OwnedDeviceWording.PhotosAdded(1, 1),
            OwnedDeviceWording.PhotosAdded(0, 3),
            OwnedDeviceWording.DeleteQuestion("Breadbin", 2),
            OwnedDeviceWording.BoardSwapped("Commodore 64|250407", "Commodore 64|250425"),
            OwnedDeviceWording.PickerHeading,
            OwnedDeviceWording.NoDevicePicked,
            OwnedDeviceWording.NoDeviceChoice,
            OwnedDeviceWording.NewRepairTitle("Breadbin"),
            OwnedDeviceWording.BarText("Breadbin", "No video"),
            OwnedDeviceWording.HistoryLabel,
            OwnedDeviceWording.NoRepairsRecorded,
            OwnedDeviceWording.HistoryRepairLine(12, "No video", "Open", "started 2026-09-28", 3),
            .. OwnedDeviceStatus.All,
        ];

        Assert.All(sentences, sentence =>
        {
            Assert.All(sentence, character => Assert.True(character < 128, $"[{sentence}] is not plain ASCII"));
            Assert.DoesNotContain("machine", sentence, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("customer", sentence, StringComparison.OrdinalIgnoreCase);
        });
    }
}
