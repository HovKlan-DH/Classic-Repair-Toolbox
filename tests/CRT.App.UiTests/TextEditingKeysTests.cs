using Avalonia;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Input;
using Avalonia.Input.Platform;
using Avalonia.Threading;
using Handlers.Theming;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// SHIFT+DELETE CUTS (owner report, 2026-11-07: "when I mark some text in a textarea and I press SHIFT+DELETE, then it
// normally cuts the text and copies it to clipboard - but it will not add it to clipboard, when done inside app?").
// Avalonia 12.1.4 lists only Ctrl+X as Cut, so the key deleted the selection and put nothing on the clipboard;
// TextEditingKeys adds it to the platform's keys at start-up, for every text box.
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class TextEditingKeysTests
{
    [Fact]
    public void Shift_Delete_is_one_of_the_keys_that_cut_and_is_added_once()
    {
        UiTest.Run(() =>
        {
            PlatformHotkeyConfiguration keys = Application.Current!.PlatformSettings!.HotkeyConfiguration;

            TextEditingKeys.UseShiftDeleteCut(keys);
            TextEditingKeys.UseShiftDeleteCut(keys);

            Assert.Single(keys.Cut, gesture => gesture.Key == Key.Delete && gesture.KeyModifiers == KeyModifiers.Shift);
            Assert.Contains(keys.Cut, gesture => gesture.Key == Key.X);
        });
    }

    [Fact]
    public async Task Shift_Delete_in_a_text_box_puts_the_selection_on_the_clipboard_and_takes_it_out()
    {
        await UiTest.RunAsync(async () =>
        {
            var box = new TextBox { Text = "Recapped the PSU", AcceptsReturn = true };
            var window = new Window { Content = box, Width = 400, Height = 200 };
            window.Show();

            try
            {
                Assert.True(box.Focus());
                box.SelectionStart = 9;
                box.SelectionEnd = 16;

                window.KeyPress(Key.Delete, RawInputModifiers.Shift, PhysicalKey.Delete, keySymbol: null);
                window.KeyRelease(Key.Delete, RawInputModifiers.Shift, PhysicalKey.Delete, keySymbol: null);

                // The cut writes the clipboard, then takes the text out - awaited inside the box.
                for (int pass = 0; pass < 10 && box.Text != "Recapped "; pass++)
                {
                    Dispatcher.UIThread.RunJobs();
                    await Task.Delay(10);
                }

                Assert.Equal("Recapped ", box.Text);
                Assert.Equal("the PSU", await window.Clipboard!.TryGetTextAsync());
            }
            finally
            {
                window.Close();
            }
        });
    }
}
