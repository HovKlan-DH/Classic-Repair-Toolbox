using Avalonia.Controls;
using Avalonia.Platform.Storage;
using CRT;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// StorageItemPaths - the local paths of dropped or pasted items, the one copy the Feedback tab and
// BoardFilesWindow share (code review, 2026-10-16). Real storage items come from the headless
// window's storage provider.
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class StorageItemPathsTests
{
    [Fact]
    public async Task Files_on_this_computer_give_their_paths_and_nothing_gives_none()
    {
        await UiTest.RunAsync(async () =>
        {
            using var workspace = new TempWorkspace();
            string first = Path.Combine(workspace.Root, "a.txt");
            string second = Path.Combine(workspace.Root, "b.png");
            File.WriteAllText(first, "a");
            File.WriteAllText(second, "b");

            var window = new Window();
            window.Show();

            try
            {
                IStorageFile? a = await window.StorageProvider.TryGetFileFromPathAsync(new Uri(first));
                IStorageFile? b = await window.StorageProvider.TryGetFileFromPathAsync(new Uri(second));

                Assert.Equal([first, second], StorageItemPaths.LocalPaths([a!, b!]).Select(Path.GetFullPath));
                Assert.Empty(StorageItemPaths.LocalPaths(null));
                Assert.Empty(StorageItemPaths.LocalPaths([]));
            }
            finally
            {
                window.Close();
            }
        });
    }
}
