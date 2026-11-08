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
            "Delete [Breadbin #2]? Its photos, comments and links are deleted with it. Its 3 repairs stay in the workbook list, on no device.",
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
        // "Photos" and "Notes" each a headline of their own (owner request, 2026-11-06: "Notes and photos" until then),
        // counted as the worklog editor counts its lists.
        Assert.Equal("Photos", OwnedDeviceWording.PhotosHeading(0));
        Assert.Equal("Photos (1)", OwnedDeviceWording.PhotosHeading(1));
        Assert.Equal("Photos (3)", OwnedDeviceWording.PhotosHeading(3));
        Assert.Equal("Links of interest", OwnedDeviceWording.LinksHeading(0));
        Assert.Equal("Links of interest (2)", OwnedDeviceWording.LinksHeading(2));
    }

    // "Edit device" says why its hardware and board are locked while it has workbooks (owner decision, 2026-11-06).
    [Fact]
    public void A_locked_board_says_why()
    {
        Assert.StartsWith("It has 1 workbook, so its hardware and board cannot change", OwnedDeviceWording.BoardLocked(1), StringComparison.Ordinal);
        Assert.StartsWith("It has 3 workbooks, so", OwnedDeviceWording.BoardLocked(3), StringComparison.Ordinal);
    }

    // The owner's own words for the workbooks on a device (2026-11-06: "rename the 'Repairs' to
    // 'Workbooks associated to this device'") - on its page and in its exported history alike.
    [Fact]
    public void The_workbooks_on_a_device_are_headed_in_the_owners_words()
    {
        Assert.Equal("Workbooks associated to this device", OwnedDeviceWording.RepairsHeading);

        // The page is only read since 2026-11-07, its "New repair" gone - and so its words for it (owner request:
        // 'remove the text ""New repair" starts one on its board."').
        Assert.Equal("No workbooks associated to this device yet.", OwnedDeviceWording.NoRepairsYet);
    }

    // ###########################################################################################
    // The description ("Where it came from" until 2026-11-06, owner request) is several lines - kept, each
    // break a "\n" whatever the box typed, so the page and the PDF break where the user did; shown as it is,
    // with no label. Nothing typed, nothing.
    // ###########################################################################################
    [Fact]
    public void The_description_keeps_its_line_breaks()
    {
        Assert.Equal("A flea market\nCame with its box", OwnedDeviceWording.DescriptionText(" A flea market\r\nCame with its box \n"));
        Assert.Equal(string.Empty, OwnedDeviceWording.DescriptionText(" \r\n "));

        // Markup: a first line nested once keeps the indent that says so (code review, 2026-11-08) - on the page and in
        // the exported history alike.
        Assert.Equal("  - recapped\n- RAM", OwnedDeviceWording.DescriptionText("\r\n  - recapped\r\n- RAM  "));
    }

    // Its condition in words - the exported history's line - within 0 to 5 stars; its comments counted in
    // their heading, as the worklog's lists are; a comment dated as a worklog's.
    [Fact]
    public void The_condition_the_comments_and_their_dates_are_said_as_the_worklog_says_them()
    {
        Assert.Equal("Condition: 0 of 5 stars", OwnedDeviceWording.ConditionLine(0));
        Assert.Equal("Condition: 4 of 5 stars", OwnedDeviceWording.ConditionLine(4));
        Assert.Equal("Condition: 5 of 5 stars", OwnedDeviceWording.ConditionLine(7));

        Assert.Equal("Comments", OwnedDeviceWording.CommentsHeading(0));
        Assert.Equal("Comments (2)", OwnedDeviceWording.CommentsHeading(2));
        Assert.Equal("2026-11-06 09:05", OwnedDeviceWording.CommentDate(new DateTime(2026, 11, 6, 9, 5, 30)));
    }

    // "My devices" with none says where one is added - the worklog bar's button above the tabs (the one
    // beside the list went on 2026-11-06), no workbook needed.
    [Fact]
    public void With_no_device_the_list_names_the_button_that_adds_one()
    {
        Assert.StartsWith("\"Add device\" above the tabs adds a computer you own", OwnedDeviceWording.NoDevicesYet, StringComparison.Ordinal);
        Assert.DoesNotContain("change", OwnedDeviceWording.NoDevicesYet, StringComparison.OrdinalIgnoreCase);
    }

    // With devices, none of them in the Configuration tab's scope: the list says so in that scope's words.
    [Theory]
    [InlineData(MyDevicesScope.CurrentHardware, "No devices for the selected hardware.")]
    [InlineData(MyDevicesScope.CurrentBoard, "No devices for the selected board.")]
    [InlineData(MyDevicesScope.AllBoards, "No devices.")]
    public void With_none_in_the_scope_the_list_says_which_scope(string scope, string expected)
    {
        Assert.Equal(expected, OwnedDeviceWording.NoDevicesInScope(scope));
    }

    // "Add workbook"'s title for a repair started on a device.
    [Fact]
    public void A_repair_started_on_a_device_is_named_after_it()
    {
        Assert.Equal("New repair on Breadbin #2", OwnedDeviceWording.NewRepairTitle(" Breadbin #2 "));
    }

    // The bar's "Device" drop-down's first choice, in the owner's words (2026-11-06).
    [Fact]
    public void The_device_drop_downs_first_choice_is_all_devices()
    {
        Assert.Equal("Show workbooks for all devices", OwnedDeviceWording.AllDevicesChoice);
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
            OwnedDeviceWording.NoDevicesInScope(MyDevicesScope.CurrentHardware),
            OwnedDeviceWording.NoDevicesInScope(MyDevicesScope.CurrentBoard),
            OwnedDeviceWording.NameNeeded,
            OwnedDeviceWording.NoRepairsYet,
            OwnedDeviceWording.DragDeviceTip,
            OwnedDeviceWording.RepairsHeading,
            OwnedDeviceWording.PhotosAdded(1, 1),
            OwnedDeviceWording.PhotosAdded(0, 3),
            OwnedDeviceWording.DeleteQuestion("Breadbin", 2),
            OwnedDeviceWording.DescriptionText("A flea market\nCame with its box"),
            OwnedDeviceWording.ClearChoice,
            OwnedDeviceWording.NoWorkbooksOfDevice,
            OwnedDeviceWording.NoWorkbookOfDeviceYet,
            OwnedDeviceWording.NoWorkbookOfDeviceOnBoard,
            OwnedDeviceWording.CommentsHeading(3),
            OwnedDeviceWording.ConditionHint,
            OwnedDeviceWording.NoDevicesMatch,
            OwnedDeviceWording.LinksHeading(2),
            OwnedDeviceWording.BoardLocked(2),
            OwnedDeviceWording.ConditionLine(3),
            OwnedDeviceWording.PickerHeading,
            OwnedDeviceWording.AllDevicesChoice,
            OwnedDeviceWording.NoDeviceChoice,
            OwnedDeviceWording.NewRepairTitle("Breadbin"),
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
