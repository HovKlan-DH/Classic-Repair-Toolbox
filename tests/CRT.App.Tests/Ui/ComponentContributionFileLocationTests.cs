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
// The component editor's file rows on a CRT 4.0 board: what its "File location" drop-down offers,
// what "Save to draft" refuses, where a picked file is filed, and how an existing file is shown.
//
// *** IN 4.0 A ROW NAMES A FILE OF ITS BOARD (2026-10-25). *** "Component images/U6.png" is the
// board's name for the file; its bytes lie in a store laid out by checksum. The editor used to work
// in the Excel tree's terms - the drop-down listed the data and drafts trees' folders, a file picked
// from the data folder took that folder as its location, and an existing image was looked for as a
// path under the data root - and in a 4.0 tree each of those lost a file or showed nothing.
//
// Read off the real rows the window builds, which is what the drop-down's ItemsSource binds to.
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class ComponentContributionFileLocationTests : IDisposable
{
    private const string BoardIdentity = "Commodore/C64/250407/board.json";

    // A minimal valid 1x1 PNG (the same bytes SchematicThumbnailSetTests uses).
    private static readonly byte[] OnePixelPng =
    {
        0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,
        0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52,
        0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01,
        0x08, 0x06, 0x00, 0x00, 0x00, 0x1F, 0x15, 0xC4,
        0x89, 0x00, 0x00, 0x00, 0x0D, 0x49, 0x44, 0x41,
        0x54, 0x78, 0x9C, 0x63, 0x00, 0x01, 0x00, 0x00,
        0x05, 0x00, 0x01, 0x0D, 0x0A, 0x2D, 0xB4, 0x00,
        0x00, 0x00, 0x00, 0x49, 0x45, 0x4E, 0x44, 0xAE,
        0x42, 0x60, 0x82
    };

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

    // ###########################################################################################
    // *** THE DROP-DOWN OFFERS THE BOARD'S OWN FOLDERS, NOT THE DATA TREE'S. *** The data and
    // drafts trees hold the store's checksum folders and CRT's bookkeeping in 4.0, and the old list
    // was built from them; none may be offered. The board's own "Truth tables" folder is, beside the
    // four every board keeps its files in.
    // ###########################################################################################
    [Fact]
    public void The_drop_down_offers_the_boards_own_folders_and_never_the_data_trees()
    {
        this.CreateFolders(this.DataRoot, "Files/ab", "Commodore/C64/250407", "MiniPro/IC tests");
        this.CreateFolders(this.DraftsRoot, "Commodore/C64/250407/Files/cd", "Commodore/C64/250407/.crt-pre-4.0");

        UiTest.Run(() =>
        {
            ComponentContributionWindow window = this.OpenC64Component(imageFile: "Truth tables/7400.png");

            ContributionComponentImageRow row = Assert.Single(
                ComponentContributionFileLocationTests.ImageRowsOf(window));

            Assert.Equal(
                new[]
                {
                    "Board local files",
                    "Component images",
                    "Component local files",
                    "Scope baseline",
                    "Truth tables",
                },
                row.AvailableFileLocations);
        });
    }

    // ###########################################################################################
    // *** REPORTED 2026-09-24, STILL HELD. *** A board created with "Add a new board" exists only
    // as a draft and has no files yet - and must still offer "Scope baseline", where every board
    // keeps its baselines, so the first one can be filed there.
    // ###########################################################################################
    [Fact]
    public void A_new_board_with_no_files_offers_the_Scope_baseline_folder()
    {
        UiTest.Run(() =>
        {
            ComponentContributionWindow window = this.OpenC64Component(imageFile: string.Empty);

            ContributionComponentImageRow row = Assert.Single(
                ComponentContributionFileLocationTests.ImageRowsOf(window));

            Assert.Contains("Scope baseline", row.AvailableFileLocations);
        });
    }

    // ###########################################################################################
    // *** THE REPORT (2026-09-25). *** A picture picked from outside the data folder, saved with
    // the "File location" box left empty, was stored as the bare "HotCPU.png". "Save to draft"
    // refuses it on the spot, marks the row, and lets go the moment a folder is chosen.
    // ###########################################################################################
    [Fact]
    public void A_picked_picture_with_no_file_location_is_refused_and_its_row_marked()
    {
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
            row.FileLocation = "Component images";

            Assert.False(row.HasLocationError);
            Assert.False(row.HasFileError);
            Assert.Equal(string.Empty, row.FileErrorText);
            Assert.Null(ComponentContributionFileLocationTests.ValidateLocations(window));
        });
    }

    // ###########################################################################################
    // *** "SAVE TO DRAFT" FILES A PICKED PICTURE WITH NO FOLDER IN THE USUAL ONE (CRT 4.0 step 1,
    // 2026-10-25 - "Unfinished work is fine. Only Submit asks for problems to be fixed"). *** It used
    // to refuse the save; the owner's report (2026-09-25) was the file stored as a bare name at the
    // top of the board, which this still never does. Through the real handler, so the filling is
    // proven to be on the save path: the row then names "Component images", and nothing is refused.
    // ###########################################################################################
    [Fact]
    public void Save_to_draft_files_a_picked_picture_with_no_file_location_in_the_usual_folder()
    {
        UiTest.Run(() =>
        {
            ComponentContributionWindow window = this.OpenC64Component();
            ContributionComponentImageRow row = this.AddPickedImageRow(window, fileLocation: string.Empty);

            typeof(ComponentContributionWindow)
                .GetMethod("OnSubmitClick", BindingFlags.Instance | BindingFlags.NonPublic)!
                .Invoke(window, new object?[] { null, new Avalonia.Interactivity.RoutedEventArgs() });

            Dispatcher.UIThread.RunJobs();

            Assert.Equal("Component images", row.FileLocation);
            Assert.False(row.HasLocationError);
            Assert.DoesNotContain("no file location", window.FindControl<TextBlock>("StatusTextBlock")!.Text ?? string.Empty);
        });
    }

    // The usual folder of each kind of row (ContributionFileLocations.DefaultFolderFor) - a scope
    // baseline's is its own.
    [Theory]
    [InlineData("ComponentPicture", "Component images")]
    [InlineData("ScopeBaseline", "Scope baseline")]
    [InlineData("ComponentFile", "Component local files")]
    [InlineData("BoardFile", "Board local files")]
    public void Each_kind_of_file_row_has_a_usual_folder_the_drop_down_offers(string kind, string folder)
    {
        Assert.Equal(folder, ContributionFileLocations.DefaultFolderFor(Enum.Parse<ContributionFileLocations.FileRowKind>(kind)));
        Assert.Contains(folder, ContributionFileLocations.UsualFolders);
        Assert.True(ContributionFileLocations.IsWritableFolder(folder));
    }

    // A folder that cannot hold a file of the board - the store - is refused with its name.
    [Fact]
    public void A_picked_picture_filed_in_the_store_is_refused()
    {
        UiTest.Run(() =>
        {
            ComponentContributionWindow window = this.OpenC64Component();
            ContributionComponentImageRow row = this.AddPickedImageRow(window, "Files/ab");

            string? message = ComponentContributionFileLocationTests.ValidateLocations(window);

            Assert.NotNull(message);
            Assert.Contains("[Files/ab]", message);
            Assert.Contains("cannot hold a file of the board", message);
            Assert.True(row.HasLocationError);
            Assert.Equal("Choose another file location", row.FileErrorText);
        });
    }

    // ###########################################################################################
    // *** A FILE PICKED FROM INSIDE THE DATA FOLDER KEEPS THE ROW'S FOLDER (changed 2026-10-25). ***
    // It used to take the data-tree folder it was picked from as its location - the published copy,
    // cited where it lay. In 4.0 that folder is the store's ("Files/ab"): the file was copied into
    // the draft under a name no draft write lists, and lost. Now it is the board's new file, filed
    // where the row says - and a row that says nothing files it in the pictures' usual folder (CRT
    // 4.0 step 1, 2026-10-25: unfinished work is fine, and a picked file is never a bare name).
    // ###########################################################################################
    [Fact]
    public void A_file_picked_from_inside_the_data_folder_is_filed_where_the_row_says()
    {
        BoardDocument.StoredFileRow stored = StoredFiles.Put(this.DataRoot, "Component images/U6.png", ComponentContributionFileLocationTests.OnePixelPng);
        string picked = StoredFiles.PathOf(this.DataRoot, stored);

        UiTest.Run(() =>
        {
            ComponentContributionWindow window = this.OpenC64Component();
            ContributionComponentImageRow row = this.AddPickedImageRow(window, "Component images");

            ComponentContributionFileLocationTests.Pick(window, row, picked);

            Assert.Equal("Component images", row.FileLocation);
            Assert.Equal(picked, row.OriginalFilePath);
            Assert.DoesNotContain(row.AvailableFileLocations, folder => folder.StartsWith("Files", StringComparison.OrdinalIgnoreCase));
            Assert.Null(ComponentContributionFileLocationTests.ValidateLocations(window));

            // Picked into a row with no folder, it is not "used where it lies" any more: it goes in
            // the pictures' usual folder.
            row.FileLocation = string.Empty;
            ComponentContributionFileLocationTests.Pick(window, row, picked);

            Assert.Equal("Component images", row.FileLocation);
            Assert.Null(ComponentContributionFileLocationTests.ValidateLocations(window));
        });
    }

    // A row loaded from the board keeps the name it is stored under, which this window does not
    // move, so its location is not the contributor's to fix here - even one a new file could not
    // be filed in.
    [Fact]
    public void A_row_loaded_from_the_board_is_not_checked()
    {
        UiTest.Run(() =>
        {
            ComponentContributionWindow window = this.OpenC64Component(imageFile: "Files/ab/U1 pin 1.png");

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
    // *** AN EXISTING IMAGE IS SHOWN BY ITS NAME (2026-10-25). *** The preview looked for the name
    // as a path under the data root, where a 4.0 tree has nothing - its bytes lie under their
    // checksum - so no existing image of any board was previewed. Found through the published
    // board's file list, as the component popup finds it.
    // ###########################################################################################
    [Fact]
    public void An_existing_image_is_previewed_through_the_published_boards_file_list()
    {
        BoardDocument.StoredFileRow stored = StoredFiles.Put(this.DataRoot, "Component images/U6.png", ComponentContributionFileLocationTests.OnePixelPng);

        var board = new BoardData
        {
            Components = { new ComponentEntry { BoardLabel = "U6", Region = "PAL" } },
            ComponentImages = { new ComponentImageEntry { BoardLabel = "U6", Region = "PAL", File = stored.Name } },
        };

        StoredFiles.WriteBoard(
            Path.Combine(this.DataRoot, "Commodore", "C64", "250407", DataTreeLayout.BoardFileName),
            board,
            stored);

        UiTest.Run(() =>
        {
            var window = new ComponentContributionWindow();
            window.LoadComponent(board, this.DataRoot, "Commodore 64", "250407", "PAL", "U6", ComponentContributionFileLocationTests.BoardIdentity);

            ContributionComponentImageRow row = Assert.Single(ComponentContributionFileLocationTests.ImageRowsOf(window));

            Assert.Equal("Component images", row.FileLocation);
            Assert.NotNull(row.PreviewImage);

            window.Close();
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

    private void CreateFolders(string root, params string[] relativeFolders)
    {
        foreach (string relative in relativeFolders)
        {
            Directory.CreateDirectory(Path.Combine(root, relative.Replace('/', Path.DirectorySeparatorChar)));
        }
    }

    // The component editor opened on U6 of the C64 250407, as from the component list. With
    // `imageFile`, U6 carries one image row naming that file.
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
            ComponentContributionFileLocationTests.BoardIdentity);

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

    // What the file picker does with the file it was given - the window's own step after the dialog.
    private static void Pick(ComponentContributionWindow window, IContributionFileRow row, string path)
    {
        typeof(ComponentContributionWindow)
            .GetMethod("ApplySelectedFilePathToRow", BindingFlags.Instance | BindingFlags.NonPublic)!
            .Invoke(window, new object?[] { row, path });
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
