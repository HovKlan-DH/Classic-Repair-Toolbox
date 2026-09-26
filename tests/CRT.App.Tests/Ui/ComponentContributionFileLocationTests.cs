using System.Collections.ObjectModel;
using System.IO;
using System.Reflection;
using Avalonia.Controls;
using Avalonia.Threading;
using Avalonia.VisualTree;
using CRT;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// The component editor's "File location" drop-down on a system that exists ONLY as a draft.
//
// *** REPORTED 2026-09-24. *** A system created with "Add a new system" gets its own empty
// "Scope baseline" folder (DraftSeeder.CreateNewSystem), but the editor's drop-down was built from
// the data root alone - so that folder was never offered, and a baseline image could not be filed
// where published boards keep theirs.
//
// End to end on purpose: the system is created by the real seeder, and the list is read off a real
// file row the window built, which is exactly what the drop-down's ItemsSource binds to. If the
// seeder and the window ever disagree about where that folder is, this fails.
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class ComponentContributionFileLocationTests : IDisposable
{
    private const string Manufacturer = "Test Manu4";
    private const string Hardware = "HW4";
    private const string Board = "Board4";

    private readonly TempWorkspace thisWorkspace = new();

    private string DataRoot => Path.Combine(this.thisWorkspace.Root, "Data");

    private string DraftsRoot => Path.Combine(this.thisWorkspace.Root, "Drafts");

    public ComponentContributionFileLocationTests()
    {
        // DraftManager is a process-wide static the window reads directly, so it is pointed at
        // the workspace before anything is built - and restored in Dispose, so this class does
        // not leak its root into the next one in the collection.
        ComponentContributionFileLocationTests.PointDraftManagerAt(this.DraftsRoot);
    }

    public void Dispose()
    {
        ComponentContributionFileLocationTests.PointDraftManagerAt(string.Empty);
        this.thisWorkspace.Dispose();
    }

    [Fact]
    public void A_NEW_systems_Scope_baseline_folder_is_offered_in_a_file_rows_drop_down()
    {
        // A published board and a shared folder, so the data tree is not empty and the test proves
        // the draft folder is ADDED to it rather than replacing it. The shared folder is the data
        // tree's half that is still offered: since 2026-09-25 another board's folders are not
        // (ContributionFileLocations.WritableBy).
        Directory.CreateDirectory(Path.Combine(this.DataRoot, "Commodore", "C64", "250407", "Scope baseline"));
        Directory.CreateDirectory(Path.Combine(this.DataRoot, "Generic shared files", "Component images"));

        string excelDataFile = NewSystemIdentity.BuildExcelDataFile(
            ComponentContributionFileLocationTests.Manufacturer,
            ComponentContributionFileLocationTests.Hardware,
            ComponentContributionFileLocationTests.Board);

        DraftSeedResult created = DraftSeeder.CreateNewSystem(this.DraftsRoot, new NewSystemRegistration
        {
            HardwareName = ComponentContributionFileLocationTests.Hardware,
            BoardName = ComponentContributionFileLocationTests.Board,
            ExcelDataFile = excelDataFile,
        });

        Assert.True(created.Created, created.Reason);

        // One component with one image row and no file picked yet - the state the report was made
        // from, where the drop-down is the thing about to be used.
        var board = new BoardData
        {
            Components = { new ComponentEntry { BoardLabel = "C1", Region = "PAL" } },
            ComponentImages = { new ComponentImageEntry { BoardLabel = "C1", Region = "PAL", Pin = "1" } },
        };

        UiTest.Run(() =>
        {
            var window = new ComponentContributionWindow();

            window.LoadComponent(
                board,
                this.DataRoot,
                ComponentContributionFileLocationTests.Hardware,
                ComponentContributionFileLocationTests.Board,
                "PAL",
                "C1",
                excelDataFile);

            ContributionComponentImageRow row = Assert.Single(
                ComponentContributionFileLocationTests.ImageRowsOf(window));

            Assert.Contains("Test Manu4/HW4/Board4/Scope baseline", row.AvailableFileLocations);
            Assert.Contains("Generic shared files/Component images", row.AvailableFileLocations);

            // Another board's folder: a new file filed there is refused by the server.
            Assert.DoesNotContain("Commodore/C64/250407/Scope baseline", row.AvailableFileLocations);
        });
    }

    // ###########################################################################################
    // *** THE OWNER'S LIST, AS THE WINDOW SHOWS IT (2026-09-25). *** The drop-down offered every
    // folder of every board of every manufacturer. It now offers this board's own folders, this
    // manufacturer's "Shared files" folders and the "Generic shared files" folders - read off a real
    // row, which is what the drop-down's ItemsSource binds to.
    // ###########################################################################################
    [Fact]
    public void The_drop_down_offers_this_boards_folders_and_the_shared_ones_only()
    {
        this.CreateDataFolders(
            "Commodore/C64/250407/Scope baseline",
            "Commodore/C128/310378/Scope baseline",
            "Commodore/Shared files/Component images",
            "Amstrad/Shared files/Component images",
            "Generic shared files/Component images");

        UiTest.Run(() =>
        {
            // One image row with no file yet, so its list holds the offered folders and nothing of
            // its own.
            ComponentContributionWindow window = this.OpenC64Component(imageFile: string.Empty);

            ContributionComponentImageRow row = Assert.Single(
                ComponentContributionFileLocationTests.ImageRowsOf(window));

            Assert.Equal(
                new[]
                {
                    "Commodore/C64/250407",
                    "Commodore/C64/250407/Scope baseline",
                    "Commodore/Shared files/Component images",
                    "Generic shared files/Component images",
                },
                row.AvailableFileLocations);
        });
    }

    // ###########################################################################################
    // *** THE REPORT (2026-09-25). *** A picture picked from outside the data folder, saved with
    // the "File location" box left empty, was stored as the bare "HotCPU.png" - and the server
    // refused the whole submission later: "[HotCPU.png] belongs to another board". "Save to draft"
    // now refuses it on the spot, marks the row, and lets go the moment a folder is chosen.
    // ###########################################################################################
    [Fact]
    public void A_picked_picture_with_no_file_location_is_refused_and_its_row_marked()
    {
        this.CreateDataFolders("Commodore/Shared files/Component images");

        UiTest.Run(() =>
        {
            ComponentContributionWindow window = this.OpenC64Component();
            ContributionComponentImageRow row = this.AddPickedImageRow(window, fileLocation: string.Empty);

            string? message = ComponentContributionFileLocationTests.ValidateLocations(window);

            Assert.NotNull(message);
            Assert.Contains("Component image #1 [HotCPU.png]", message);
            Assert.Contains("no file location", message);
            Assert.True(row.HasLocationError);
            Assert.True(row.HasFileError);
            Assert.Equal("Choose a file location", row.FileErrorText);

            // Choosing a folder answers it at once, before the next save.
            row.FileLocation = "Commodore/Shared files/Component images";

            Assert.False(row.HasLocationError);
            Assert.False(row.HasFileError);
            Assert.Equal(string.Empty, row.FileErrorText);
            Assert.Null(ComponentContributionFileLocationTests.ValidateLocations(window));
        });
    }

    // Through the real "Save to draft" handler, so the check is proven to be on the save path and
    // not only callable. Refused before anything is written, with the message on the status line.
    [Fact]
    public void Save_to_draft_refuses_a_picked_picture_with_no_file_location()
    {
        UiTest.Run(() =>
        {
            ComponentContributionWindow window = this.OpenC64Component();
            this.AddPickedImageRow(window, fileLocation: string.Empty);

            typeof(ComponentContributionWindow)
                .GetMethod("OnSubmitClick", BindingFlags.Instance | BindingFlags.NonPublic)!
                .Invoke(window, new object?[] { null, new Avalonia.Interactivity.RoutedEventArgs() });

            Dispatcher.UIThread.RunJobs();

            Assert.Contains(
                "Component image #1 [HotCPU.png] has no file location",
                window.FindControl<TextBlock>("StatusTextBlock")!.Text);

            Assert.False(
                Directory.Exists(Path.Combine(this.DraftsRoot, "Commodore")),
                "the refused save wrote a draft");
        });
    }

    // The same refusal the server gives, for a folder that belongs to another board.
    [Fact]
    public void A_picked_picture_filed_in_another_boards_folder_is_refused()
    {
        UiTest.Run(() =>
        {
            ComponentContributionWindow window = this.OpenC64Component();
            ContributionComponentImageRow row = this.AddPickedImageRow(window, "Commodore/C128/310378/Scope baseline");

            string? message = ComponentContributionFileLocationTests.ValidateLocations(window);

            Assert.NotNull(message);
            Assert.Contains("[Commodore/C128/310378/Scope baseline]", message);
            Assert.Contains("not one of this board's folders", message);
            Assert.True(row.HasLocationError);
        });
    }

    // A file picked FROM another board's folder and left there is the published copy itself -
    // citing it unchanged is allowed, here as on the server.
    [Fact]
    public void Another_boards_published_file_left_where_it_is_is_accepted()
    {
        UiTest.Run(() =>
        {
            ComponentContributionWindow window = this.OpenC64Component();
            ContributionComponentImageRow row = this.AddPickedImageRow(window, "Commodore/C128/310378/Scope baseline");
            row.OriginalFilePath = Path.Combine(this.DataRoot, "Commodore", "C128", "310378", "Scope baseline", "HotCPU.png");

            Assert.Null(ComponentContributionFileLocationTests.ValidateLocations(window));
            Assert.False(row.HasLocationError);
        });
    }

    // A row loaded from the board keeps the path it is stored under, which this window does not
    // move, so its location is not the contributor's to fix here.
    [Fact]
    public void A_row_loaded_from_the_board_is_not_checked()
    {
        UiTest.Run(() =>
        {
            ComponentContributionWindow window = this.OpenC64Component(
                imageFile: "Commodore/C128/310378/Scope baseline/U1 pin 1.png");

            Assert.Null(ComponentContributionFileLocationTests.ValidateLocations(window));
            Assert.False(Assert.Single(ComponentContributionFileLocationTests.ImageRowsOf(window)).HasLocationError);
        });
    }

    // Every section with files is checked, not only the images - and the row is named by section.
    [Fact]
    public void A_picked_board_file_with_no_file_location_is_refused_too()
    {
        UiTest.Run(() =>
        {
            ComponentContributionWindow window = this.OpenC64Component();

            var row = new ContributionBoardLocalFileRow
            {
                File = "Service manual.pdf",
                OriginalFilePath = Path.Combine(this.thisWorkspace.Root, "Downloads", "Service manual.pdf"),
            };

            ComponentContributionFileLocationTests.BoardFileRowsOf(window).Add(row);

            string? message = ComponentContributionFileLocationTests.ValidateLocations(window);

            Assert.NotNull(message);
            Assert.Contains("Board file #1 [Service manual.pdf]", message);
            Assert.True(row.HasLocationError);
        });
    }

    // ###########################################################################################
    // The mark has to reach the screen: model -> binding -> style class -> brush, each step silently
    // survivable if broken. Read off the real "File location" box in a shown window.
    // ###########################################################################################
    [Fact]
    public void The_file_location_box_turns_red_on_screen()
    {
        UiTest.Run(() =>
        {
            ComponentContributionWindow window = this.OpenC64Component();
            this.AddPickedImageRow(window, fileLocation: string.Empty);

            window.Show();
            window.FindControl<Expander>("ComponentImagesExpander")!.IsExpanded = true;
            ComponentContributionFileLocationTests.PumpLayout(window);

            ComboBox locationBox = window.FindControl<ItemsControl>("ComponentImageRowsItemsControl")!
                .GetVisualDescendants()
                .OfType<ComboBox>()
                .First();

            Assert.DoesNotContain("LocationError", locationBox.Classes);

            ComponentContributionFileLocationTests.ValidateLocations(window);
            ComponentContributionFileLocationTests.PumpLayout(window);

            Assert.True(window.TryFindResource("Text_Fail_Fg", window.ActualThemeVariant, out object? failBrush));
            Assert.Contains("LocationError", locationBox.Classes);
            Assert.Equal(failBrush, locationBox.BorderBrush);

            window.Close();
        });
    }

    private void CreateDataFolders(params string[] relativeFolders)
    {
        foreach (string relative in relativeFolders)
        {
            Directory.CreateDirectory(Path.Combine(this.DataRoot, relative.Replace('/', Path.DirectorySeparatorChar)));
        }
    }

    // The component editor opened on U6 of the C64 250407, as from the component list. With
    // `imageFile`, U6 carries one image row stored at that path.
    private ComponentContributionWindow OpenC64Component(string? imageFile = null)
    {
        var board = new BoardData
        {
            Components = { new ComponentEntry { BoardLabel = "U6", Region = "PAL" } },
        };

        if (imageFile != null)
        {
            board.ComponentImages.Add(new ComponentImageEntry { BoardLabel = "U6", Region = "PAL", File = imageFile });
        }

        var window = new ComponentContributionWindow();

        window.LoadComponent(
            board,
            this.DataRoot,
            "Commodore 64",
            "250407",
            "PAL",
            "U6",
            "Commodore/C64/250407/Data C64 250407.xlsx");

        return window;
    }

    // An image row as the file picker leaves it: a picture from outside the data folder (a full
    // path in OriginalFilePath) and whatever the "File location" box shows.
    private ContributionComponentImageRow AddPickedImageRow(ComponentContributionWindow window, string fileLocation)
    {
        var row = new ContributionComponentImageRow
        {
            BoardLabel = "U6",
            Region = "PAL",
            File = "HotCPU.png",
            OriginalFilePath = Path.Combine(this.thisWorkspace.Root, "Downloads", "HotCPU.png"),
            FileLocation = fileLocation,
        };

        ComponentContributionFileLocationTests.ImageRowsOf(window).Add(row);

        return row;
    }

    // The status-line message of the first row ValidateNewFileLocations refuses, or null.
    private static string? ValidateLocations(ComponentContributionWindow window)
    {
        var method = typeof(ComponentContributionWindow).GetMethod(
            "ValidateNewFileLocations",
            BindingFlags.Instance | BindingFlags.NonPublic);

        object? problem = method!.Invoke(window, null);

        // The nullable tuple comes back boxed as its underlying (Reveal, Message) value.
        return problem == null ? null : (string)problem.GetType().GetField("Item2")!.GetValue(problem)!;
    }

    private static ObservableCollection<ContributionBoardLocalFileRow> BoardFileRowsOf(ComponentContributionWindow window)
    {
        var field = typeof(ComponentContributionWindow).GetField(
            "thisBoardLocalFileRows",
            BindingFlags.Instance | BindingFlags.NonPublic);

        return (ObservableCollection<ContributionBoardLocalFileRow>)field!.GetValue(window)!;
    }

    // Headless windows do not lay out on their own, and an ItemsControl builds no containers until
    // it has been measured - so the dispatcher is drained and a layout pass forced by hand.
    private static void PumpLayout(Window window)
    {
        Dispatcher.UIThread.RunJobs();
        window.Measure(window.ClientSize);
        window.Arrange(new Avalonia.Rect(window.ClientSize));
        Dispatcher.UIThread.RunJobs();
    }

    private static ObservableCollection<ContributionComponentImageRow> ImageRowsOf(ComponentContributionWindow window)
    {
        var field = typeof(ComponentContributionWindow).GetField(
            "thisComponentImageRows",
            BindingFlags.Instance | BindingFlags.NonPublic);

        return (ObservableCollection<ContributionComponentImageRow>)field!.GetValue(window)!;
    }

    // The same seam ComponentContributionSaveRefreshTests uses - DraftManager's own internal
    // LoadFrom, never Load(), which would resolve the user's real AppData folder.
    private static void PointDraftManagerAt(string draftsRoot)
    {
        var method = typeof(DraftManager).GetMethod(
            "LoadFrom",
            BindingFlags.Static | BindingFlags.NonPublic);

        method!.Invoke(null, new object?[] { draftsRoot });
    }
}
