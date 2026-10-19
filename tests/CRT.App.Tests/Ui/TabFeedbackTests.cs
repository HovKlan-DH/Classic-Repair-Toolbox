using Avalonia.Controls;
using Avalonia.Input;
using CRT;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// TabFeedback on screen - what a send carries (FeedbackAttachmentNamesTests holds the tab's
// contract with the server, FeedbackZipTests its zip).
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class TabFeedbackTests
{
    // A minimal valid 1x1 PNG.
    private static readonly byte[] OnePixelPng =
    [
        0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52,
        0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1F, 0x15, 0xC4,
        0x89, 0x00, 0x00, 0x00, 0x0D, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9C, 0x63, 0x00, 0x01, 0x00, 0x00,
        0x05, 0x00, 0x01, 0x0D, 0x0A, 0x2D, 0xB4, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4E, 0x44, 0xAE,
        0x42, 0x60, 0x82
    ];

    // ###########################################################################################
    // Code review, 2026-10-16: Send read the field and the attachments at the click, while a
    // paste of files was still decoding its pictures - and the Clear after a successful send then
    // threw away what arrived. What is sent is read once everything on its way in has arrived: the
    // picture in the text, and the file that is no picture in the attachments.
    // ###########################################################################################
    [Fact]
    public async Task What_is_sent_waits_for_a_paste_still_putting_things_in()
    {
        await UiTest.RunAsync(async () =>
        {
            using var workspace = new TempWorkspace();
            string shot = Path.Combine(workspace.Root, "shot.png");
            string notes = Path.Combine(workspace.Root, "notes.txt");
            File.WriteAllBytes(shot, OnePixelPng);
            File.WriteAllText(notes, "not a picture");

            var tab = new TabFeedback();
            var window = new Window { Content = tab, Width = 900, Height = 700 };
            window.Show();

            try
            {
                FeedbackComposer composer = tab.GetControl<FeedbackComposer>("FeedbackText");
                TextBox part = composer.PartsForTests[0];
                part.Text = "Words.";
                part.Focus();
                part.CaretIndex = part.Text.Length;

                TopLevel top = TopLevel.GetTopLevel(window)!;
                var transfer = new DataTransfer();

                foreach (string path in new[] { shot, notes })
                {
                    var item = new DataTransferItem();
                    item.SetFile((await top.StorageProvider.TryGetFileFromPathAsync(new Uri(path)))!);
                    transfer.Add(item);
                }

                await top.Clipboard!.SetDataAsync(transfer);
                part.Paste();

                FeedbackToSend toSend = await tab.GatherAsync();

                Assert.Equal("Words.\n[shot.png]", toSend.Text);
                Assert.Equal(["shot.png"], toSend.Pictures.Select(picture => picture.Name));
                Assert.Equal([notes], toSend.Attachments);
            }
            finally
            {
                window.Close();
            }
        });
    }
}
