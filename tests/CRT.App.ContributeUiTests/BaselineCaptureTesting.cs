using Avalonia;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Threading;
using CRT;
using Handlers.DataHandling;
using Handlers.Theming;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// What the tests of "Capture from oscilloscope pin-by-pin" share (`using static`) - ComponentPaneBaselineRunTests (the pane:
// offering the window, opening and closing it, "Capture again", the captures it takes and puts into the draft) and
// BaselineCaptureWindowTests (the window itself): a draft of U27, a 74LS08 with baselines of pins 1 and 3, in a workspace of its
// own (Draft); the pane on its "Oscilloscope baseline" page with a fake oscilloscope (FakeBaselineScope), and the window opened
// from it.
// ###########################################################################################
internal static class BaselineCaptureTesting
{
    internal const string BoardKey = "Commodore/C64/250407/board.json";

    internal static readonly byte[] Png = FakeBaselineScope.Png;

    // U27, a 74LS08 in its case, its pinout partly typed, and baselines of pins 1 and 3 - pin 1's measured with the machine
    // at the READY prompt. U26 another chip, C1 a part with no case and no pins.
    internal static BoardData Published() => new()
    {
        Components =
        [
            new ComponentEntry { BoardLabel = "U27", FriendlyName = "AND", TechnicalNameOrValue = "74LS08", Category = "IC", Package = "DIP-14" },
            new ComponentEntry { BoardLabel = "U26", FriendlyName = "OR", TechnicalNameOrValue = "74LS32", Category = "IC", Package = "DIP-14" },
            new ComponentEntry { BoardLabel = "C1", FriendlyName = "Decoupling", TechnicalNameOrValue = "100nF", Category = "Capacitor" },
        ],
        ComponentPins =
        [
            new ComponentPinEntry { BoardLabel = "U27", Pin = "1", Name = "1A", Direction = "Input" },
            new ComponentPinEntry { BoardLabel = "U27", Pin = "2", Name = "1B", Direction = "Input" },
            new ComponentPinEntry { BoardLabel = "U27", Pin = "3", Name = "1Y", Direction = "Output" },
            new ComponentPinEntry { BoardLabel = "U27", Pin = "7", Name = "GND", Direction = "Ground" },
            new ComponentPinEntry { BoardLabel = "U27", Pin = "14", Name = "VCC", Direction = "Power" },
        ],
        ComponentImages =
        [
            new ComponentImageEntry { BoardLabel = "U27", Pin = "1", Name = "1A", File = "Scope baseline/U27_1.png", TimeDiv = "50us", VoltsDiv = "1V", TriggerLevelVolts = "1.5V", MachineState = "At the READY prompt" },
            new ComponentImageEntry { BoardLabel = "U27", Pin = "3", Name = "1Y", File = "Scope baseline/U27_3.png", TimeDiv = "1us", VoltsDiv = "1V", TriggerLevelVolts = "1.5V" },
        ],
    };

    // ###########################################################################################
    // The board as a draft, in a workspace of its own - a test class holds one and lets go of it as it is disposed.
    // ###########################################################################################
    internal sealed class Draft : IDisposable
    {
        private readonly TempWorkspace thisWorkspace = new();

        private string DraftsRoot => Path.Combine(this.thisWorkspace.Root, "Drafts");

        private string BoardFile => DraftFolderLayout.GetWorkbookPath(this.DraftsRoot, BoardKey);

        private string DraftFolder => Path.GetDirectoryName(this.BoardFile)!;

        public void Dispose() => this.thisWorkspace.Dispose();

        // The draft - pin 1's earlier baseline the picture `earlier`, when given.
        internal DraftSession Open(byte[]? earlier = null)
        {
            Directory.CreateDirectory(this.DraftFolder);
            BoardFiles.Write(this.BoardFile, BoardDocumentMapping.ToDocument(
                new BoardContent(Published(), [], []),
                name => name,
                [this.Stored("Scope baseline/U27_1.png", earlier ?? Png), this.Stored("Scope baseline/U27_3.png", Png)]));

            return DraftSession.Open(this.DraftsRoot, BoardKey)!;
        }

        // `name`'s bytes in the draft's own store, and its entry.
        private BoardDocument.StoredFileRow Stored(string name, byte[] bytes)
        {
            string sha = Convert.ToHexStringLower(System.Security.Cryptography.SHA256.HashData(bytes));
            string path = Path.Combine(this.DraftFolder, DataTreeLayout.StoredFilePath(sha, name).Replace('/', Path.DirectorySeparatorChar));
            Directory.CreateDirectory(Path.GetDirectoryName(path)!);
            File.WriteAllBytes(path, bytes);

            return new BoardDocument.StoredFileRow { Name = name, Sha256 = sha, Size = bytes.Length };
        }
    }

    // A PNG of `width` x `height` in one color - an oscilloscope's screen, as far as a comparison goes.
    internal static byte[] Screen(int width, int height, SkiaSharp.SKColor color)
    {
        using var bitmap = new SkiaSharp.SKBitmap(width, height);
        bitmap.Erase(color);
        using SkiaSharp.SKData data = bitmap.Encode(SkiaSharp.SKEncodedImageFormat.Png, 100);
        return data.ToArray();
    }

    // The pane on U27's "Oscilloscope baseline" page, the oscilloscope connected.
    internal static (ComponentPane Pane, Window Window) Shown(DraftSession session, FakeBaselineScope scope)
    {
        var pane = new ComponentPane { AutoSave = false };
        var window = new Window { Content = pane, Width = 1100, Height = 1400 };
        window.Show();

        pane.Show(session, "U27", "PAL", Published());
        pane.ShowPage(ComponentPage.Baselines);
        pane.Scope = scope;
        pane.ShowScopeConnected(true);
        Dispatcher.UIThread.RunJobs();

        return (pane, window);
    }

    // "Capture from oscilloscope pin-by-pin" pressed - the pictures it opens noted in `viewers`, when given.
    internal static BaselineCaptureWindow Open(ComponentPane pane, List<WorklogPhotoViewerWindow>? viewers = null)
    {
        BaselineCaptureWindow window = pane.OpenBaselineCaptureWindow()!;
        Assert.NotNull(window);
        window.ShowPictureOverrideForTests = viewer => (viewers ?? []).Add(viewer);
        Dispatcher.UIThread.RunJobs();
        return window;
    }

    internal static T Get<T>(ComponentPane pane, string name) where T : Control => pane.GetControl<T>(name);

    internal static T In<T>(BaselineCaptureWindow window, string name) where T : Control => window.GetControl<T>(name);

    internal static void Click(Button button) => button.RaiseEvent(new RoutedEventArgs(Button.ClickEvent));

    internal static IReadOnlyList<ComponentImageEntry> Baselines(DraftSession session) =>
        ComponentPaneEdits.BaselinesOf(session.Content.Board, "U27", "PAL");

    internal static List<string> Picked(BaselineCaptureWindow window) =>
        window.TargetsForTests.Where(target => target.IsPicked).Select(target => target.Step.PinText).ToList();

    internal static string Said(BaselineCaptureWindow window, string name) => BoldRuns.TextOf(In<TextBlock>(window, name));

    // A baseline of U27 added to the draft by another editor - the table, say - as another editor's change.
    internal static void AddBaseline(DraftSession session, ComponentImageEntry baseline) =>
        session.Apply("Another editor", content => content with
        {
            Board = content.Board.WithComponentSections(componentImages: [.. content.Board.ComponentImages, baseline]),
        });

    // The bold pieces of a note with formatting (RichNote) - the viewer's line under a picture.
    internal static List<string> BoldOfNote(string markup)
    {
        RichNoteText note = RichNote.Parse(markup);
        var pieces = new List<string>();
        int start = -1;

        for (int index = 0; index <= note.Text.Length; index++)
        {
            bool bold = index < note.Text.Length && note.Styles[index].HasFlag(RichNoteStyle.Bold);

            if (bold && start < 0)
            {
                start = index;
            }
            else if (!bold && start >= 0)
            {
                pieces.Add(note.Text[start..index]);
                start = -1;
            }
        }

        return pieces;
    }

    internal static async Task Until(Func<bool> condition, string what)
    {
        for (int attempt = 0; attempt < 300 && !condition(); attempt++)
        {
            Dispatcher.UIThread.RunJobs();
            await Task.Delay(10);
        }

        Assert.True(condition(), what);
    }

    // The window's pictures and the pane's thumbnails read - their files may go then (Windows refuses a file being read).
    internal static async Task Settled(ComponentPane pane, BaselineCaptureWindow window)
    {
        await window.ImagesReadForTests;
        Dispatcher.UIThread.RunJobs();
        await window.ImagesReadForTests;
        await pane.ThumbnailsLoadedForTests;
    }

    // A real press and release of the left button in the middle of `control`, on its window as last drawn.
    internal static void Press(BaselineCaptureWindow window, Control control)
    {
        AvaloniaHeadlessPlatform.ForceRenderTimerTick();
        Dispatcher.UIThread.RunJobs();

        Point point = control.TranslatePoint(new Point(control.Bounds.Width / 2, control.Bounds.Height / 2), window)!.Value;

        window.MouseDown(point, MouseButton.Left);
        window.MouseUp(point, MouseButton.Left);
        Dispatcher.UIThread.RunJobs();
    }
}
