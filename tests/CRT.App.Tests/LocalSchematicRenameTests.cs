using Avalonia;
using CRT;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// LocalSchematicRename - a board's schematic images renamed, and what THIS machine keeps under their
// names following them (owner decisions, 2026-10-10): the worklogs marked on an image and each
// workbook's image order, the user's thumbnail order, the image last looked at, and the images
// hidden on the Configuration tab. Caught up whenever the board is shown, by the IMAGE FILE each
// name was on the last time; the first time, by the board's published list. Where a single name
// goes is SchematicRenameLog's (CRT.Data.Tests); these cover what is moved on this machine.
//
// Touches UserSettings' static, so this joins the "UserSettings" collection; it also points
// WorklogManager at a temp folder, which is safe in any collection because collections never run
// in parallel (xunit.runner.json) - and it is pointed away again in Dispose.
[Collection("UserSettings")]
public sealed class LocalSchematicRenameTests : IDisposable
{
    private readonly TempWorkspace thisWorkspace = new();

    private const string Hardware = "Commodore 128";
    private const string Board = "310378 Open128";
    private const string BoardKey = "Commodore 128|310378 Open128";

    public LocalSchematicRenameTests()
    {
        this.LoadSettings();
        this.LoadWorklogs();
    }

    public void Dispose()
    {
        // Detach from the temp folder so nothing written later can reach the real files.
        this.LoadSettings();
        this.LoadWorklogs();
        this.thisWorkspace.Dispose();
    }

    private void LoadSettings()
    {
        string path = this.thisWorkspace.Path_(Guid.NewGuid().ToString("N") + ".json");
        File.WriteAllText(path, "{}");
        UserSettings.LoadFrom(path);
    }

    private void LoadWorklogs() =>
        WorklogManager.LoadFrom(this.thisWorkspace.Path_("Workbooks-" + Guid.NewGuid().ToString("N")));

    private static string HiddenKey(string schematic) => CatalogueVisibility.BuildKey(Hardware, Board, schematic);

    private static SchematicImage Image(string name, string file) => new(name, $"Commodore/C128/310378 Open128/{file}");

    // The board shown with these images: CatchUp as Main's board load calls it.
    private static void Show(params SchematicImage[] images) => ShowWith([], images);

    private static void ShowWith(SchematicRenameRecord[] published, params SchematicImage[] images) =>
        LocalSchematicRename.CatchUp(Hardware, Board, images, () => published);

    private static readonly SchematicImage[] AsPublished =
        [Image("Top", "top.png"), Image("Bottom", "bottom.png"), Image("Clock", "clock.png")];

    private static readonly SchematicImage[] TopRenamed =
        [Image("PCB top side", "top.png"), Image("Bottom", "bottom.png"), Image("Clock", "clock.png")];

    // Everything the user keeps under a name: a worklog marked on it, the workbook's and the
    // Schematics tab's orders, the image last looked at, and the image hidden.
    private (WorkbookRecord Workbook, WorklogEntryRecord Entry) KeepThingsUnder(string name)
    {
        WorkbookRecord workbook = WorklogManager.CreateWorkbook(BoardKey, "Open128 repair", "")!;
        WorklogEntryRecord entry = WorklogManager.AddEntry(workbook.Id, name, new Rect(1, 2, 3, 4), "Bad cap", "C12 bulging", "Issue", "Open", ["C12"])!;
        WorklogManager.UpdateWorkbookSchematicOrder(workbook.Id, ["Bottom", name]);

        UserSettings.SetSchematicsOrder(BoardKey, ["Bottom", name, "Clock"]);
        UserSettings.SetLastSchematicForBoard(BoardKey, name);
        UserSettings.SetCatalogueKeyChecked(HiddenKey(name), false);

        return (workbook, entry);
    }

    private static void AssertEverythingIsUnder(string name, WorkbookRecord workbook, int entryId)
    {
        WorklogEntryRecord entry = WorklogManager.GetEntries(workbook.Id).Single(item => item.Id == entryId);

        Assert.Equal(name, entry.SchematicName);
        Assert.Equal(["Bottom", name], WorklogManager.GetWorkbooksForBoard(BoardKey).Single(item => item.Id == workbook.Id).SchematicOrder);
        Assert.Equal(["Bottom", name, "Clock"], UserSettings.GetSchematicsOrder(BoardKey));
        Assert.Equal(name, UserSettings.GetLastSchematicForBoard(BoardKey));
        Assert.False(UserSettings.IsCatalogueKeyChecked(HiddenKey(name)));
    }

    // ------------------------------------------------------------------ following the IMAGE FILE

    // ###########################################################################################
    // *** A RENAME - published, or the user's own draft. *** CRT last showed the board with "Top"
    // on top.png; now top.png is "PCB top side", and everything the user kept under "Top" moves
    // there - the worklog's own text, area and components untouched.
    // ###########################################################################################
    [Fact]
    public void A_renamed_image_takes_everything_kept_under_its_old_name()
    {
        Show(AsPublished);
        (WorkbookRecord workbook, WorklogEntryRecord entry) = this.KeepThingsUnder("Top");

        Show(TopRenamed);

        AssertEverythingIsUnder("PCB top side", workbook, entry.Id);
        Assert.True(UserSettings.IsCatalogueKeyChecked(HiddenKey("Top")));

        WorklogEntryRecord moved = WorklogManager.GetEntries(workbook.Id).Single();
        Assert.Equal("Bad cap", moved.Title);
        Assert.Equal("C12 bulging", moved.Description);
        Assert.Equal(new Rect(1, 2, 3, 4), new Rect(moved.AreaX, moved.AreaY, moved.AreaWidth, moved.AreaHeight));
        Assert.Equal(["C12"], moved.ComponentLabels);
    }

    // ###########################################################################################
    // *** THE OWNER'S CASE: the old name later given to ANOTHER image. *** "Top" (top.png) was
    // renamed, and bottom.png is now called "Top". What was made on "Top" was made on top.png, so
    // it goes with top.png - and what was made on "Bottom" goes with bottom.png, to "Top".
    // ###########################################################################################
    [Fact]
    public void A_name_given_to_another_image_does_not_keep_what_was_made_on_the_first()
    {
        Show(AsPublished);
        (WorkbookRecord topWorkbook, WorklogEntryRecord onTop) = this.KeepThingsUnder("Top");
        WorklogEntryRecord onBottom = WorklogManager.AddEntry(topWorkbook.Id, "Bottom", new Rect(0, 0, 1, 1), "Cold joint", "", "Issue", "Open", [])!;

        Show(Image("PCB top side", "top.png"), Image("Top", "bottom.png"), Image("Clock", "clock.png"));

        List<WorklogEntryRecord> entries = WorklogManager.GetEntries(topWorkbook.Id);
        Assert.Equal("PCB top side", entries.Single(entry => entry.Id == onTop.Id).SchematicName);
        Assert.Equal("Top", entries.Single(entry => entry.Id == onBottom.Id).SchematicName);
        Assert.Equal("PCB top side", UserSettings.GetLastSchematicForBoard(BoardKey));
    }

    // ###########################################################################################
    // *** A RENAME UNDONE. *** The draft discarded or rejected, "view boards as officially
    // published" ticked, or a rename pushed back out of BETA: the board shows the old name again,
    // and everything comes back (code review, 2026-10-10: it used to stay on a name the board no
    // longer had).
    // ###########################################################################################
    [Fact]
    public void A_rename_undone_brings_everything_back()
    {
        Show(AsPublished);
        (WorkbookRecord workbook, WorklogEntryRecord entry) = this.KeepThingsUnder("Top");

        Show(TopRenamed);
        AssertEverythingIsUnder("PCB top side", workbook, entry.Id);

        Show(AsPublished);
        AssertEverythingIsUnder("Top", workbook, entry.Id);
    }

    // Two images swapping names: each keeps its own things - hidden ones included, both at once.
    [Fact]
    public void Two_images_swapping_names_each_keep_their_own_things()
    {
        Show(AsPublished);
        WorkbookRecord workbook = WorklogManager.CreateWorkbook(BoardKey, "Open128 repair", "")!;
        WorklogEntryRecord onTop = WorklogManager.AddEntry(workbook.Id, "Top", new Rect(0, 0, 1, 1), "A", "", "Issue", "Open", [])!;
        WorklogEntryRecord onBottom = WorklogManager.AddEntry(workbook.Id, "Bottom", new Rect(0, 0, 1, 1), "B", "", "Issue", "Open", [])!;
        UserSettings.SetCatalogueKeyChecked(HiddenKey("Top"), false);
        UserSettings.SetCatalogueKeyChecked(HiddenKey("Bottom"), false);

        Show(Image("Bottom", "top.png"), Image("Top", "bottom.png"), Image("Clock", "clock.png"));

        List<WorklogEntryRecord> entries = WorklogManager.GetEntries(workbook.Id);
        Assert.Equal("Bottom", entries.Single(entry => entry.Id == onTop.Id).SchematicName);
        Assert.Equal("Top", entries.Single(entry => entry.Id == onBottom.Id).SchematicName);
        Assert.False(UserSettings.IsCatalogueKeyChecked(HiddenKey("Top")));
        Assert.False(UserSettings.IsCatalogueKeyChecked(HiddenKey("Bottom")));
    }

    // Only this board: another board may well have an image of the same name.
    [Fact]
    public void Another_boards_things_are_left_alone()
    {
        Show(AsPublished);
        WorkbookRecord other = WorklogManager.CreateWorkbook("Commodore 64|250407", "C64 repair", "")!;
        WorklogManager.AddEntry(other.Id, "Top", new Rect(0, 0, 1, 1), "Dead RAM", "", "Issue", "Open", []);
        UserSettings.SetLastSchematicForBoard("Commodore 64|250407", "Top");

        Show(TopRenamed);

        Assert.Equal("Top", WorklogManager.GetEntries(other.Id).Single().SchematicName);
        Assert.Equal("Top", UserSettings.GetLastSchematicForBoard("Commodore 64|250407"));
    }

    // ------------------------------------------------------------------ the FIRST time a board is shown

    // ###########################################################################################
    // *** THE OWNER'S FIRST CASE: a rename published while this CRT had never shown the board -
    // freshly updated, with worklogs from before. *** Nothing remembered, so the board's .json
    // says where "Top" went, and everything follows.
    // ###########################################################################################
    [Fact]
    public void The_first_time_the_boards_published_list_says_where_names_went()
    {
        (WorkbookRecord workbook, WorklogEntryRecord entry) = this.KeepThingsUnder("Top");

        ShowWith([new SchematicRenameRecord("Top", "PCB top side", "Commodore/C128/310378 Open128/top.png", "2026-October-10")], TopRenamed);

        AssertEverythingIsUnder("PCB top side", workbook, entry.Id);
    }

    // The first time, a name the board has stays - which image an old worklog meant cannot be known.
    [Fact]
    public void The_first_time_a_name_the_board_has_stays()
    {
        (WorkbookRecord workbook, WorklogEntryRecord entry) = this.KeepThingsUnder("Top");

        ShowWith(
            [new SchematicRenameRecord("Top", "PCB top side", "Commodore/C128/310378 Open128/top.png", "2026-October-10")],
            Image("PCB top side", "top.png"), Image("Top", "bottom.png"), Image("Clock", "clock.png"));

        AssertEverythingIsUnder("Top", workbook, entry.Id);
    }

    // A board shown for the first time with no renames, or shown as it was last time: nothing moves.
    [Fact]
    public void A_board_shown_as_it_was_moves_nothing()
    {
        (WorkbookRecord workbook, WorklogEntryRecord entry) = this.KeepThingsUnder("Top");

        Show(TopRenamed);
        Show(TopRenamed);

        AssertEverythingIsUnder("Top", workbook, entry.Id);
    }

    // What is remembered is the board as shown, name and file.
    [Fact]
    public void The_board_as_shown_is_remembered()
    {
        Show(TopRenamed);

        Assert.True(SchematicRenameLog.SameImages(TopRenamed, UserSettings.GetSchematicImages(BoardKey)!));
        Assert.Null(UserSettings.GetSchematicImages("Commodore 64|250407"));
    }

    // ###########################################################################################
    // *** A BOARD SHOWN WITH NO IMAGES CHANGES NOTHING (code review, 2026-10-11). *** A new board's
    // draft "viewed as officially published" is shown as an empty board, and so is one whose workbook
    // could not be read. Remembering that forgot the images: a rename made meanwhile was then not
    // followed by its file, and the worklog stayed on a name the board no longer had.
    // ###########################################################################################
    [Fact]
    public void A_board_shown_with_no_images_leaves_what_is_remembered()
    {
        Show(AsPublished);
        (WorkbookRecord workbook, WorklogEntryRecord entry) = this.KeepThingsUnder("Top");

        Show();
        Assert.True(SchematicRenameLog.SameImages(AsPublished, UserSettings.GetSchematicImages(BoardKey)!));

        Show(TopRenamed);
        AssertEverythingIsUnder("PCB top side", workbook, entry.Id);
    }

    // The board's published list is read only when it can be needed - not when the board is shown as
    // it was last time, which is nearly every board load (code review, 2026-10-11).
    [Fact]
    public void The_published_list_is_read_only_when_the_board_changed()
    {
        int reads = 0;
        IReadOnlyList<SchematicRenameRecord> Published()
        {
            reads++;
            return [];
        }

        LocalSchematicRename.CatchUp(Hardware, Board, AsPublished, Published);
        Assert.Equal(1, reads);

        LocalSchematicRename.CatchUp(Hardware, Board, AsPublished, Published);
        Assert.Equal(1, reads);

        LocalSchematicRename.CatchUp(Hardware, Board, TopRenamed, Published);
        Assert.Equal(2, reads);
    }

    // ###########################################################################################
    // *** A WORKLOG THAT CANNOT BE WRITTEN MOVES NOTHING, AND IS TRIED AGAIN (code review,
    // 2026-10-11). *** It kept its old name while the board was remembered as shown, so it was never
    // followed again. Now the worklogs already moved go back and nothing is remembered - shown again,
    // everything moves. Two images swapping names is the case a retry from half-way would get wrong:
    // the worklog already moved would be moved back. The write is stopped by a folder where its
    // temporary file goes, which stops it on every system.
    // ###########################################################################################
    [Fact]
    public void A_worklog_that_cannot_be_written_moves_nothing_and_the_next_showing_moves_everything()
    {
        Show(AsPublished);
        WorkbookRecord workbook = WorklogManager.CreateWorkbook(BoardKey, "Open128 repair", "")!;
        WorklogEntryRecord onTop = WorklogManager.AddEntry(workbook.Id, "Top", new Rect(0, 0, 1, 1), "A", "", "Issue", "Open", [])!;
        WorklogEntryRecord onBottom = WorklogManager.AddEntry(workbook.Id, "Bottom", new Rect(0, 0, 1, 1), "B", "", "Issue", "Open", [])!;
        UserSettings.SetLastSchematicForBoard(BoardKey, "Top");

        string blocked = Path.Combine(
            WorklogManager.GetEntryAttachmentsFolderPath(workbook.Id, onBottom.Id)!,
            AppConfig.WorklogIndexFileName + ".tmp");
        Directory.CreateDirectory(blocked);

        SchematicImage[] swapped = [Image("Bottom", "top.png"), Image("Top", "bottom.png"), Image("Clock", "clock.png")];
        Show(swapped);

        List<WorklogEntryRecord> entries = WorklogManager.GetEntries(workbook.Id);
        Assert.Equal("Top", entries.Single(entry => entry.Id == onTop.Id).SchematicName);
        Assert.Equal("Bottom", entries.Single(entry => entry.Id == onBottom.Id).SchematicName);
        Assert.Equal("Top", UserSettings.GetLastSchematicForBoard(BoardKey));
        Assert.True(SchematicRenameLog.SameImages(AsPublished, UserSettings.GetSchematicImages(BoardKey)!));

        Directory.Delete(blocked);
        Show(swapped);

        entries = WorklogManager.GetEntries(workbook.Id);
        Assert.Equal("Bottom", entries.Single(entry => entry.Id == onTop.Id).SchematicName);
        Assert.Equal("Top", entries.Single(entry => entry.Id == onBottom.Id).SchematicName);
        Assert.Equal("Bottom", UserSettings.GetLastSchematicForBoard(BoardKey));
    }

    // A settings file holding a null for a remembered image - a hand edit - is read round it: the
    // board load does not stop on it (code review, 2026-10-11).
    [Fact]
    public void A_null_in_the_remembered_images_does_not_stop_the_catch_up()
    {
        string path = this.thisWorkspace.Path_(Guid.NewGuid().ToString("N") + ".json");
        File.WriteAllText(
            path,
            "{\"schematicImagesByBoard\": {\"" + BoardKey + "\": " +
            "[{\"name\": null, \"file\": null}, null, {\"name\": \"Top\", \"file\": \"Commodore/C128/310378 Open128/top.png\"}]}}");
        UserSettings.LoadFrom(path);
        (WorkbookRecord workbook, WorklogEntryRecord entry) = this.KeepThingsUnder("Top");

        Show(TopRenamed);

        AssertEverythingIsUnder("PCB top side", workbook, entry.Id);
    }

    // ------------------------------------------------------------------ a list of names

    private static string? TopToPcbTopSide(string name) =>
        SchematicRename.SameName(name, "Top") ? "PCB top side" : null;

    [Fact]
    public void A_moved_name_keeps_its_place_in_a_list()
    {
        Assert.Equal(["Bottom", "PCB top side", "Clock"], LocalSchematicRename.Renamed(["Bottom", "top", "Clock"], TopToPcbTopSide));
    }

    [Fact]
    public void A_list_with_nothing_to_move_is_left_alone()
    {
        Assert.Null(LocalSchematicRename.Renamed(["Bottom", "Clock"], TopToPcbTopSide));
        Assert.Null(LocalSchematicRename.Renamed(null, TopToPcbTopSide));
    }

    // A list that somehow holds both keeps the new name once, where it first stood.
    [Fact]
    public void A_list_already_holding_the_new_name_keeps_it_once()
    {
        Assert.Equal(["PCB top side", "Bottom"], LocalSchematicRename.Renamed(["PCB top side", "Top", "Bottom"], TopToPcbTopSide));
    }
}
