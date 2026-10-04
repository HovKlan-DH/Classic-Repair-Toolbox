using Avalonia.Controls;
using CRT;
using ClassicRepairToolbox.Tests.Maintainer;
using Handlers.DataHandling;
using Handlers.MaintainerHandling;

namespace ClassicRepairToolbox.Tests.Ui.Maintainer;

// ###########################################################################################
// "Unused files" on the Account screen, UnusedFilesView - a window of its own until 2026-09-27. Its
// list comes from the server, answered here by AnsweringHttpHandler (test rule 6).
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class UnusedFilesViewTests
{
    private static readonly ReviewSession Session =
        new("token", DateTimeOffset.UtcNow.AddDays(1), 1, "admin@example.com", "Admin");

    [Fact]
    public void Unused_files_builds_with_its_remove_button_off()
    {
        UiTest.Run(() =>
        {
            var view = new UnusedFilesView();

            Assert.False(view.FindControl<Button>("RemoveButton")!.IsEnabled);
            Assert.False(view.FindControl<CheckBox>("LookedThroughCheckBox")!.IsEnabled);
            Assert.Equal(2, view.FindControl<ComboBox>("TreeComboBox")!.ItemCount);

            view.Clear();
            Assert.False(view.FindControl<Button>("RemoveButton")!.IsEnabled);
        });
    }

    // ###########################################################################################
    // *** THE LIST IS THE FILE TREE (owner request, 2026-10-04: "the exact same tree-view like it does
    // in 'Systems' and 'Files' ... including visualization of images and opening of files"). *** Every
    // folder open, each file with its size, a source of bytes to show and open them from - and the
    // tree's own line off, since the summary above says the count and the size.
    // ###########################################################################################
    [Fact]
    public async Task The_unused_files_are_drawn_as_the_file_tree_with_their_sizes()
    {
        await UiTest.RunAsync(async () =>
        {
            const string Json =
                "{\"tree\":\"beta\",\"isComplete\":true,\"problems\":[],\"masterCount\":2,\"boardWorkbookCount\":22,\"fileCount\":10961," +
                "\"files\":[{\"path\":\"Commodore/Shared files/Component images/6510.jpg\",\"sizeBytes\":46182}," +
                "{\"path\":\"Generic shared files/Component images/7408.jpg\",\"sizeBytes\":57446}]," +
                "\"publicDataUrl\":\"https://example.org/app-data-BETA/Data/\"}";

            var view = new UnusedFilesView();
            view.Initialize(
                new ReviewApiClient("https://review.invalid", new HttpClient(new AnsweringHttpHandler(_ => AnsweringHttpHandler.Json(Json)))),
                UnusedFilesViewTests.Session);

            await view.LoadAsync();

            FileTreeView tree = view.FileTreeForTests;

            Assert.Equal(
                ["Commodore", "Shared files", "Component images", "6510.jpg", "Generic shared files", "Component images", "7408.jpg"],
                tree.RowsForTests.Select(row => row.Name));
            Assert.Equal("45.1 KB", tree.RowsForTests.Single(row => row.Name == "6510.jpg").Size);
            Assert.Equal("56.1 KB", tree.RowsForTests.Single(row => row.Name == "7408.jpg").Size);
            Assert.NotNull(tree.Files);
            Assert.False(tree.ShowSummary);
            Assert.Equal("Remove 2 files (101.2 KB)", view.FindControl<Button>("RemoveButton")!.Content);

            view.Clear();
            Assert.Empty(tree.RowsForTests);
        });
    }
}
