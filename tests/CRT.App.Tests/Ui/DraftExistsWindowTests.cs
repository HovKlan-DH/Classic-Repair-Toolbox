using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Input;
using Avalonia.Threading;
using CRT;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// DraftExistsWindow - "Edit board as draft" on a board that already has a draft (owner request,
// 2026-10-09: "it should be informed via popup, that there cannot be two draft for same board").
// It names the board, and Enter opens the draft - the step the person came for - while Escape and
// Close leave everything as it is.
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class DraftExistsWindowTests
{
    private static DraftExistsWindow Shown()
    {
        var window = new DraftExistsWindow();
        window.Initialize("Commodore 64 / 250407");
        window.Show();
        Dispatcher.UIThread.RunJobs();
        return window;
    }

    [Fact]
    public void It_names_the_board_and_offers_the_draft_there_is()
    {
        UiTest.Run(() =>
        {
            DraftExistsWindow window = Shown();

            Assert.Equal("Commodore 64 / 250407", window.GetControl<TextBlock>("BoardNameText").Text);
            Assert.Equal("Open the draft", window.GetControl<Button>("OpenDraftButton").Content);
            Assert.Equal("Close", window.GetControl<Button>("CloseButton").Content);

            window.Close();
        });
    }

    // ###########################################################################################
    // Enter follows the focus (code review, 2026-10-09): "Open the draft" is the default, so Enter
    // opens the draft with it focused or with no button focused - but with "Close" focused, Enter
    // closes, as the focused button says. A window-wide Enter handler opened the draft regardless.
    // Escape closes from anywhere.
    // ###########################################################################################
    [Theory]
    [InlineData(null, Key.Enter, PhysicalKey.Enter, true)]
    [InlineData("OpenDraftButton", Key.Enter, PhysicalKey.Enter, true)]
    [InlineData("CloseButton", Key.Enter, PhysicalKey.Enter, false)]
    [InlineData("OpenDraftButton", Key.Escape, PhysicalKey.Escape, false)]
    public void Enter_answers_as_the_focused_button_says_and_Escape_closes(string? focused, Key key, PhysicalKey physical, bool opens)
    {
        UiTest.Run(() =>
        {
            DraftExistsWindow window = Shown();

            if (focused is not null)
                window.GetControl<Button>(focused).Focus();

            window.KeyPress(key, RawInputModifiers.None, physical, keySymbol: null);
            Dispatcher.UIThread.RunJobs();

            Assert.False(window.IsVisible);
            Assert.Equal(opens, window.OpenChosen);
        });
    }
}
