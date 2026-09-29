using Avalonia.Controls;
using CRT;
using ClassicRepairToolbox.Tests.Maintainer;

namespace ClassicRepairToolbox.Tests.Ui.Maintainer;

// ###########################################################################################
// "Unused files" on the Admin screen, UnusedFilesView - a window of its own until 2026-09-27. Its
// list comes only from the server, so this pins what can be pinned without one: it builds, and with
// no list nothing can be removed.
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class UnusedFilesViewTests
{
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
}
