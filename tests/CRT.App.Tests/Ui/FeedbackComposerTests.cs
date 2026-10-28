using Avalonia;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Threading;
using Avalonia.VisualTree;
using CRT;
using Handlers.DataHandling;
using Handlers.Geometry;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// FeedbackComposer - the Feedback tab's text with pictures in it (owner request, 2026-10-15). What
// must hold: a picture goes in where the caret is and the text goes on under it; Backspace under
// it, Delete above it and its bin take it away and join the text again; Up and Down cross it;
// Ctrl+A then Delete clears everything; what is sent is the text with a marker line per picture;
// a pasted picture's number is never reused; dropped files become pictures or attachments.
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class FeedbackComposerTests
{
    private static FeedbackPicture Picture(string name) =>
        new(name, null, new PixelSize(640, 480), () => Task.FromResult(new byte[] { 1, 2, 3 }));

    // A picture with a bitmap to show, so it takes up room as a real one does.
    private static FeedbackPicture Shown(string name, int width, int height) =>
        new(name, new Avalonia.Media.Imaging.WriteableBitmap(new PixelSize(width, height), new Vector(96, 96)), new PixelSize(width, height), () => Task.FromResult(new byte[] { 1 }));

    private static Point CentreOf(Window window, Control control) =>
        control.TranslatePoint(new Point(control.Bounds.Width / 2, control.Bounds.Height / 2), window)!.Value;

    // A corner handle pressed, dragged by (dx, dy) and let go, as a hand would.
    private static void DragHandle(Window window, Border handle, double dx, double dy, Action? whileDragging = null)
    {
        Point from = CentreOf(window, handle);
        Point to = from + new Point(dx, dy);

        window.MouseDown(from, MouseButton.Left);
        Dispatcher.UIThread.RunJobs();
        window.MouseMove(from + new Point(dx / 2, dy / 2), RawInputModifiers.LeftMouseButton);
        window.MouseMove(to, RawInputModifiers.LeftMouseButton);
        Dispatcher.UIThread.RunJobs();
        whileDragging?.Invoke();
        window.MouseUp(to, MouseButton.Left);
        Dispatcher.UIThread.RunJobs();
    }

    // A minimal valid 1x1 PNG (the bytes SchematicThumbnailSetTests and WorkbooksBoardPreviewTests use).
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

    private static (Window Window, FeedbackComposer Composer) Show()
    {
        var composer = new FeedbackComposer { PlaceholderText = "Describe it" };
        var window = new Window { Content = composer, Width = 700, Height = 500 };
        window.Show();
        Dispatcher.UIThread.RunJobs();
        return (window, composer);
    }

    // The composer beside another control that can take the focus away from every part.
    private static (Window Window, FeedbackComposer Composer, Button Elsewhere) ShowBeside()
    {
        var composer = new FeedbackComposer { PlaceholderText = "Describe it" };
        var elsewhere = new Button { Content = "elsewhere" };
        var window = new Window { Content = new StackPanel { Children = { composer, elsewhere } }, Width = 700, Height = 500 };
        window.Show();
        Dispatcher.UIThread.RunJobs();
        return (window, composer, elsewhere);
    }

    // Types into the focused part by putting the caret there first, as a user's click would.
    private static void PutCaret(TextBox part, int caret)
    {
        part.Focus();
        part.CaretIndex = caret;
        Dispatcher.UIThread.RunJobs();
    }

    private static void Press(Window window, Key key, RawInputModifiers modifiers = RawInputModifiers.None)
    {
        PhysicalKey physical = key switch
        {
            Key.Back => PhysicalKey.Backspace,
            Key.Delete => PhysicalKey.Delete,
            Key.Up => PhysicalKey.ArrowUp,
            Key.Down => PhysicalKey.ArrowDown,
            Key.A => PhysicalKey.A,
            _ => PhysicalKey.None,
        };

        window.KeyPress(key, modifiers, physical, keySymbol: null);
        Dispatcher.UIThread.RunJobs();
    }

    private static string[] Texts(FeedbackComposer composer) =>
        composer.PartsForTests.Select(part => part.Text ?? string.Empty).ToArray();

    [Fact]
    public void It_starts_as_one_empty_part_showing_the_placeholder()
    {
        UiTest.Run(() =>
        {
            var (_, composer) = Show();

            Assert.True(composer.IsEmpty);
            Assert.Equal("", composer.ComposeText());
            TextBox only = Assert.Single(composer.PartsForTests);
            Assert.Equal("Describe it", only.PlaceholderText);
            Assert.Empty(composer.Pictures);
        });
    }

    [Fact]
    public void A_picture_goes_in_at_the_caret_and_the_text_goes_on_under_it()
    {
        UiTest.Run(() =>
        {
            var (_, composer) = Show();
            TextBox first = composer.PartsForTests[0];
            first.Text = "one two";
            PutCaret(first, 4);

            composer.InsertPicture(Picture("pasted-image-1.png"));

            Assert.Equal(["one ", "two"], Texts(composer));
            Assert.Equal("one \n[pasted-image-1.png]\ntwo", composer.ComposeText());
            Assert.Single(composer.Pictures);
            Assert.True(composer.PartsForTests[1].IsFocused);
            Assert.Equal(0, composer.PartsForTests[1].CaretIndex);
            Assert.False(composer.IsEmpty);

            // The placeholder belongs to an empty field only.
            Assert.Null(first.PlaceholderText);
        });
    }

    [Fact]
    public void With_no_part_focused_a_picture_goes_at_the_end()
    {
        UiTest.Run(() =>
        {
            var (_, composer, elsewhere) = ShowBeside();
            composer.PartsForTests[0].Text = "typed";
            elsewhere.Focus();
            Dispatcher.UIThread.RunJobs();
            Assert.DoesNotContain(composer.PartsForTests, part => part.IsFocused);

            composer.InsertPicture(Picture("a.png"));

            Assert.Equal(["typed", ""], Texts(composer));
            Assert.Equal("typed\n[a.png]", composer.ComposeText());
        });
    }

    [Fact]
    public void Backspace_at_the_start_of_the_part_under_a_picture_removes_it_and_joins_the_text()
    {
        UiTest.Run(() =>
        {
            var (window, composer) = Show();
            composer.PartsForTests[0].Text = "one two";
            PutCaret(composer.PartsForTests[0], 4);
            composer.InsertPicture(Picture("a.png"));
            Dispatcher.UIThread.RunJobs();

            Press(window, Key.Back);

            TextBox only = Assert.Single(composer.PartsForTests);
            Assert.Equal("one two", only.Text);
            Assert.Equal(4, only.CaretIndex);
            Assert.True(only.IsFocused);
            Assert.Empty(composer.Pictures);
            Assert.Equal("one two", composer.ComposeText());
        });
    }

    [Fact]
    public void Backspace_anywhere_else_is_the_parts_own()
    {
        UiTest.Run(() =>
        {
            var (window, composer) = Show();
            composer.PartsForTests[0].Text = "one two";
            PutCaret(composer.PartsForTests[0], 4);
            composer.InsertPicture(Picture("a.png"));
            PutCaret(composer.PartsForTests[1], 2);

            Press(window, Key.Back);

            Assert.Equal(["one ", "to"], Texts(composer));
            Assert.Single(composer.Pictures);
        });
    }

    [Fact]
    public void Delete_at_the_end_of_the_part_above_a_picture_removes_it()
    {
        UiTest.Run(() =>
        {
            var (window, composer) = Show();
            composer.PartsForTests[0].Text = "one two";
            PutCaret(composer.PartsForTests[0], 4);
            composer.InsertPicture(Picture("a.png"));
            PutCaret(composer.PartsForTests[0], 4);

            Press(window, Key.Delete);

            TextBox only = Assert.Single(composer.PartsForTests);
            Assert.Equal("one two", only.Text);
            Assert.Equal(4, only.CaretIndex);
            Assert.Empty(composer.Pictures);
        });
    }

    [Fact]
    public void The_bin_on_a_picture_removes_it()
    {
        UiTest.Run(() =>
        {
            var (_, composer) = Show();
            composer.PartsForTests[0].Text = "ab";
            PutCaret(composer.PartsForTests[0], 1);
            composer.InsertPicture(Picture("a.png"));
            Dispatcher.UIThread.RunJobs();

            Button bin = composer.GetVisualDescendants().OfType<Button>().Single(button => button.Classes.Contains("RemovePicture"));
            Assert.False(bin.Focusable);
            bin.RaiseEvent(new RoutedEventArgs(Button.ClickEvent));

            Assert.Equal(["ab"], Texts(composer));
            Assert.Empty(composer.Pictures);
        });
    }

    [Fact]
    public void Up_on_the_first_line_goes_to_the_end_of_the_part_above_and_Down_on_the_last_to_the_start_of_the_one_below()
    {
        UiTest.Run(() =>
        {
            var (window, composer) = Show();
            composer.PartsForTests[0].Text = "above";
            PutCaret(composer.PartsForTests[0], 5);
            composer.InsertPicture(Picture("a.png"));
            composer.PartsForTests[1].Text = "below";
            PutCaret(composer.PartsForTests[1], 2);

            Press(window, Key.Up);

            Assert.True(composer.PartsForTests[0].IsFocused);
            Assert.Equal(5, composer.PartsForTests[0].CaretIndex);

            Press(window, Key.Down);

            Assert.True(composer.PartsForTests[1].IsFocused);
            Assert.Equal(0, composer.PartsForTests[1].CaretIndex);
        });
    }

    [Fact]
    public void Up_and_Down_inside_a_part_stay_the_parts_own()
    {
        UiTest.Run(() =>
        {
            var (window, composer) = Show();
            composer.PartsForTests[0].Text = "above";
            PutCaret(composer.PartsForTests[0], 5);
            composer.InsertPicture(Picture("a.png"));
            composer.PartsForTests[1].Text = "first\nsecond";
            PutCaret(composer.PartsForTests[1], 8);

            Press(window, Key.Up);

            Assert.True(composer.PartsForTests[1].IsFocused);
            Assert.InRange(composer.PartsForTests[1].CaretIndex, 0, 5);
        });
    }

    [Fact]
    public void Ctrl_A_then_Delete_clears_the_whole_field()
    {
        UiTest.Run(() =>
        {
            var (window, composer) = Show();
            composer.PartsForTests[0].Text = "one";
            PutCaret(composer.PartsForTests[0], 3);
            composer.InsertPicture(Picture("a.png"));
            composer.PartsForTests[1].Text = "two";
            PutCaret(composer.PartsForTests[1], 1);

            Press(window, Key.A, RawInputModifiers.Control);

            Assert.All(composer.PartsForTests, part => Assert.Equal(part.Text, part.SelectedText));

            Press(window, Key.Delete);

            Assert.True(composer.IsEmpty);
            Assert.Equal("Describe it", composer.PartsForTests[0].PlaceholderText);

            // The caret was in the part that went: the field keeps the keyboard (code review,
            // 2026-10-16 - the next keys went nowhere).
            Assert.True(composer.PartsForTests[0].IsFocused);
        });
    }

    // Code review, 2026-10-16: Ctrl+A, then away from the field and back with the keyboard, and a
    // Backspace meant for one character erased the whole feedback.
    [Fact]
    public void Ctrl_A_ends_when_the_focus_leaves_the_field()
    {
        UiTest.Run(() =>
        {
            var (window, composer, elsewhere) = ShowBeside();
            composer.PartsForTests[0].Text = "one";
            PutCaret(composer.PartsForTests[0], 3);
            composer.InsertPicture(Picture("a.png"));
            composer.PartsForTests[1].Text = "two";
            PutCaret(composer.PartsForTests[1], 3);
            Press(window, Key.A, RawInputModifiers.Control);

            elsewhere.Focus();
            Dispatcher.UIThread.RunJobs();
            PutCaret(composer.PartsForTests[1], 3);
            Press(window, Key.Back);

            Assert.Single(composer.Pictures);
            Assert.Equal(["one", "tw"], Texts(composer));
        });
    }

    // After Ctrl+A then Delete the field is an ordinary one again: a Backspace under a picture put
    // in since takes that picture, not everything.
    [Fact]
    public void After_Ctrl_A_then_Delete_the_field_is_an_ordinary_one_again()
    {
        UiTest.Run(() =>
        {
            var (window, composer) = Show();
            composer.PartsForTests[0].Text = "one";
            PutCaret(composer.PartsForTests[0], 3);
            composer.InsertPicture(Picture("a.png"));
            Press(window, Key.A, RawInputModifiers.Control);
            Press(window, Key.Delete);

            composer.PartsForTests[0].Text = "kept";
            PutCaret(composer.PartsForTests[0], 4);
            composer.InsertPicture(Picture("b.png"));
            Press(window, Key.Back);

            Assert.Empty(composer.Pictures);
            Assert.Equal(["kept"], Texts(composer));
        });
    }

    // A single part keeps the TextBox's own Ctrl+A - and a key other than Delete or Backspace after
    // the composer's Ctrl+A is that part's own again.
    [Fact]
    public void Ctrl_A_followed_by_typing_only_changes_the_part_typed_in()
    {
        UiTest.Run(() =>
        {
            var (window, composer) = Show();
            composer.PartsForTests[0].Text = "one";
            PutCaret(composer.PartsForTests[0], 3);
            composer.InsertPicture(Picture("a.png"));
            composer.PartsForTests[1].Text = "two";
            PutCaret(composer.PartsForTests[1], 1);

            Press(window, Key.A, RawInputModifiers.Control);
            Press(window, Key.Right);
            Press(window, Key.Delete);

            Assert.Single(composer.Pictures);
            Assert.Equal("one", composer.PartsForTests[0].Text);
        });
    }

    [Fact]
    public void Clear_leaves_one_empty_part_and_starts_the_numbering_again()
    {
        UiTest.Run(() =>
        {
            var (_, composer) = Show();
            composer.PartsForTests[0].Text = "text";
            PutCaret(composer.PartsForTests[0], 4);
            composer.InsertPicture(Picture("pasted-image-1.png"));
            composer.PartsForTests[1].Text = "more";

            composer.Clear();

            Assert.True(composer.IsEmpty);
            Assert.Single(composer.PartsForTests);
            Assert.Equal("Describe it", composer.PartsForTests[0].PlaceholderText);
        });
    }

    // Removing picture 1 and pasting again gives 3, never 1 again - the markers in the text that
    // is sent must not name two pictures alike.
    [Fact]
    public async Task A_pasted_pictures_number_is_never_handed_out_twice_until_Clear()
    {
        await UiTest.RunAsync(async () =>
        {
            var (_, composer) = Show();

            // A pasted bitmap is the composer's from then on (it goes with the picture), so each
            // paste gets a bitmap of its own, as each paste from the clipboard does.
            static Avalonia.Media.Imaging.Bitmap Small() =>
                new Avalonia.Media.Imaging.WriteableBitmap(new PixelSize(4, 4), new Vector(96, 96));

            await composer.AddPastedBitmapAsync(Small());
            await composer.AddPastedBitmapAsync(Small());
            composer.RemovePicture(composer.Pictures[0]);
            await composer.AddPastedBitmapAsync(Small());

            // Number 3, never 1 again - and it sits ABOVE picture 2: removing picture 1 left the
            // caret at the join, which is where the third paste goes.
            Assert.Equal(["pasted-image-3.png", "pasted-image-2.png"], composer.Pictures.Select(picture => picture.Name));

            composer.Clear();
            await composer.AddPastedBitmapAsync(Small());

            Assert.Equal(["pasted-image-1.png"], composer.Pictures.Select(picture => picture.Name));
        });
    }

    [Fact]
    public async Task Dropped_files_become_pictures_under_their_own_names_and_other_files_go_to_the_attachments()
    {
        await UiTest.RunAsync(async () =>
        {
            using var workspace = new TempWorkspace();
            string shot = Path.Combine(workspace.Root, "shot.png");
            string twin = Path.Combine(workspace.Root, "twin", "shot.png");
            string notes = Path.Combine(workspace.Root, "notes.txt");
            Directory.CreateDirectory(Path.GetDirectoryName(twin)!);
            File.WriteAllBytes(shot, OnePixelPng);
            File.WriteAllBytes(twin, OnePixelPng);
            File.WriteAllText(notes, "not a picture");

            var (_, composer) = Show();
            IReadOnlyList<string>? forAttachment = null;
            composer.FilesForAttachment += (_, paths) => forAttachment = paths;

            await composer.AddFilesAsync([shot, notes, twin]);

            Assert.Equal(["shot.png", "shot (2).png"], composer.Pictures.Select(picture => picture.Name));
            Assert.NotNull(forAttachment);
            Assert.Equal([notes], forAttachment);
            Assert.Equal("[shot.png]\n[shot (2).png]", composer.ComposeText());
            Assert.Equal(OnePixelPng, await composer.Pictures[0].ReadBytesAsync());
        });
    }

    // Enter and then Ctrl+V (owner report, 2026-10-15): the picture takes the line the Enter made,
    // so no empty line shows above it - and gets it back, caret and all, when the picture goes.
    [Fact]
    public void A_picture_pasted_after_Enter_takes_that_line_and_gives_it_back_when_removed()
    {
        UiTest.Run(() =>
        {
            var (window, composer) = Show();
            composer.PartsForTests[0].Text = "see:\n";
            PutCaret(composer.PartsForTests[0], 5);

            composer.InsertPicture(Picture("a.png"));

            Assert.Equal(["see:", ""], Texts(composer));
            Assert.Equal("see:\n[a.png]", composer.ComposeText());

            Press(window, Key.Back);

            TextBox only = Assert.Single(composer.PartsForTests);
            Assert.Equal("see:\n", only.Text);
            Assert.Equal(5, only.CaretIndex);
        });
    }

    // ------------------------------------------------------------------ the size in the mail

    // Owner request, 2026-10-15: "show the inserted image with the real size it will be shown in mail".
    [Fact]
    public void A_picture_is_shown_at_the_size_it_will_have_in_the_mail()
    {
        UiTest.Run(() =>
        {
            var (window, composer) = Show();

            composer.InsertPicture(Shown("small.png", 400, 300));
            Dispatcher.UIThread.RunJobs();

            Image image = composer.PictureControlsForTests[0].Image;
            Assert.Equal(400, image.Bounds.Width, 1);
            Assert.Equal(300, image.Bounds.Height, 1);
            Assert.Equal("{\"small.png\":400}", composer.PictureWidthsText());
        });
    }

    // A screenshot wider than the mail's column is shown at the column's width - or the field's,
    // when that is narrower - and goes to the mail at the column's width.
    [Fact]
    public void A_wide_picture_is_as_wide_as_the_mail_allows_and_no_wider_than_the_field()
    {
        UiTest.Run(() =>
        {
            var (window, composer) = Show();

            composer.InsertPicture(Shown("wide.png", 3840, 2160));
            Dispatcher.UIThread.RunJobs();

            FeedbackPicture picture = Assert.Single(composer.Pictures);
            Assert.Equal(FeedbackComposition.MaxPictureWidth, picture.DisplayWidth);
            Assert.Equal(551, picture.DisplayHeight);

            // The window is 700 wide: the field is narrower than the mail, so the picture fills it.
            Image image = composer.PictureControlsForTests[0].Image;
            Assert.True(image.Bounds.Width < 700, $"{image.Bounds.Width}");
            Assert.Equal(image.Bounds.Width * 2160 / 3840, image.Bounds.Height, 1);
        });
    }

    // Owner request, 2026-10-15: "give the option to resize in the corners".
    [Fact]
    public void Dragging_a_corner_in_makes_the_picture_smaller_in_the_field_and_in_the_mail()
    {
        UiTest.Run(() =>
        {
            var (window, composer) = Show();
            composer.InsertPicture(Shown("shot.png", 400, 300));
            Dispatcher.UIThread.RunJobs();

            (Image image, IReadOnlyList<Border> handles) = composer.PictureControlsForTests[0];
            Border bottomRight = handles.Single(handle => Equals(handle.Tag, PictureCorner.BottomRight));
            TextBlock? size = null;

            DragHandle(window, bottomRight, -100, 0, () =>
            {
                // The size it will have in the mail is written on it while dragging.
                size = composer.GetVisualDescendants().OfType<TextBlock>().Single(text => text.Parent is Border { IsVisible: true } border && border.Classes.Contains("PictureSize"));
                Assert.Equal("300 x 225", size.Text);
            });

            Assert.Equal(300, Assert.Single(composer.Pictures).DisplayWidth);
            Assert.Equal(300, image.Bounds.Width, 1);
            Assert.Equal(225, image.Bounds.Height, 1);
            Assert.Equal("{\"shot.png\":300}", composer.PictureWidthsText());
            Assert.False(((Border)size!.Parent!).IsVisible);
        });
    }

    // A left corner grows the picture as it is dragged left, and nothing makes it wider than itself.
    [Fact]
    public void A_picture_is_never_dragged_wider_than_itself()
    {
        UiTest.Run(() =>
        {
            var (window, composer) = Show();
            composer.InsertPicture(Shown("shot.png", 400, 300));
            Dispatcher.UIThread.RunJobs();

            Border topLeft = composer.PictureControlsForTests[0].Handles.Single(handle => Equals(handle.Tag, PictureCorner.TopLeft));

            DragHandle(window, topLeft, 80, 0);
            Assert.Equal(320, composer.Pictures[0].DisplayWidth);

            DragHandle(window, topLeft, -300, 0);
            Assert.Equal(400, composer.Pictures[0].DisplayWidth);
        });
    }

    // A dropped picture's own size comes from its header, so a wide one is shown at the mail's
    // column at once - its own size kept for the resizing.
    [Fact]
    public async Task A_dropped_wide_picture_keeps_its_own_size_and_starts_at_the_mails_column()
    {
        await UiTest.RunAsync(async () =>
        {
            using var workspace = new TempWorkspace();
            string wide = Path.Combine(workspace.Root, "wide.png");
            byte[] png = (byte[])OnePixelPng.Clone();
            png[18] = 0x07; png[19] = 0xD0;   // IHDR width 2000
            png[23] = 0x64;                   // IHDR height 100
            File.WriteAllBytes(wide, png);

            var (_, composer) = Show();
            await composer.AddFilesAsync([wide]);

            FeedbackPicture picture = Assert.Single(composer.Pictures);
            Assert.Equal(new PixelSize(2000, 100), picture.NaturalSize);
            Assert.Equal(FeedbackComposition.MaxPictureWidth, picture.DisplayWidth);
        });
    }

    // ------------------------------------------------------------------ code review, 2026-10-16

    // A pasted picture's PNG is encoded from the very bitmap shown, on the pool. Removed while that
    // ran (its bin, Backspace, Ctrl+A then Delete), the bitmap was disposed under the encoder,
    // which read freed memory and could end CRT. It is now let go only once the encoding is done.
    //
    // *** KNOWN INTERMITTENT (2026-10-28). *** Failed once in a full run - "Expected: 4, 4 / Actual:
    // 1, 1" on the PixelSize assertion - and passed alone and in three full runs after it; the run's
    // seed was not kept. A 1x1 bitmap is what the headless platform hands out for one DECODED from a
    // stream, which nothing on this path does, so the shown bitmap read was another than the one
    // pasted. Whoever picks this up: replay with the seed of a failing run (CLAUDE.md, "Tests").
    [Fact]
    public async Task A_pasted_picture_removed_while_its_PNG_is_encoded_is_let_go_only_once_the_encoding_is_done()
    {
        await UiTest.RunAsync(async () =>
        {
            var (_, composer) = Show();
            using var hold = new ManualResetEventSlim();
            composer.BeforeEncodingForTests = () => hold.Wait(TimeSpan.FromSeconds(30));

            await composer.AddPastedBitmapAsync(new Avalonia.Media.Imaging.WriteableBitmap(new PixelSize(4, 4), new Vector(96, 96)));
            FeedbackPicture picture = Assert.Single(composer.Pictures);
            Avalonia.Media.Imaging.Bitmap shown = picture.Thumbnail!;

            composer.RemovePicture(picture);
            Dispatcher.UIThread.RunJobs();

            // Still there for the encoder, which then finishes without reading a disposed bitmap.
            Assert.Equal(new PixelSize(4, 4), shown.PixelSize);
            hold.Set();
            await picture.ReadBytesAsync();

            await WaitUntil(() => IsDisposed(shown));
        });
    }

    // Excel, LibreOffice Calc and Numbers put a picture of copied cells on the clipboard beside
    // their text - and the text is what was meant. It was pasted as a screenshot of the cells.
    [Fact]
    public async Task Text_copied_with_a_picture_beside_it_is_pasted_as_text()
    {
        await UiTest.RunAsync(async () =>
        {
            var (window, composer) = Show();
            var item = new DataTransferItem();
            item.SetText("U8\t6510");
            item.SetBitmap(new Avalonia.Media.Imaging.WriteableBitmap(new PixelSize(4, 4), new Vector(96, 96)));
            await SetClipboard(window, item);

            TextBox part = composer.PartsForTests[0];
            PutCaret(part, 0);
            part.Paste();

            await WaitUntil(() => part.Text == "U8\t6510");
            Assert.Empty(composer.Pictures);
        });
    }

    // A screenshot tool puts a picture alone on the clipboard: that is still a picture.
    [Fact]
    public async Task A_picture_alone_on_the_clipboard_is_pasted_as_a_picture()
    {
        await UiTest.RunAsync(async () =>
        {
            var (window, composer) = Show();
            var item = new DataTransferItem();
            item.SetBitmap(new Avalonia.Media.Imaging.WriteableBitmap(new PixelSize(4, 4), new Vector(96, 96)));
            await SetClipboard(window, item);

            TextBox part = composer.PartsForTests[0];
            PutCaret(part, 0);
            part.Paste();

            await WaitUntil(() => composer.Pictures.Count == 1);
            Assert.Equal("pasted-image-1.png", composer.Pictures[0].Name);
        });
    }

    // Finder's Copy puts the file, its name as text and its icon as a picture on the clipboard -
    // the file is what was copied.
    [Fact]
    public async Task A_file_copied_with_its_name_and_icon_beside_it_is_pasted_as_the_file()
    {
        await UiTest.RunAsync(async () =>
        {
            using var workspace = new TempWorkspace();
            string shot = Path.Combine(workspace.Root, "shot.png");
            File.WriteAllBytes(shot, OnePixelPng);

            var (window, composer) = Show();
            var item = new DataTransferItem();
            Avalonia.Platform.Storage.IStorageFile? file = await TopLevel.GetTopLevel(window)!.StorageProvider.TryGetFileFromPathAsync(new Uri(shot));
            Assert.NotNull(file);
            item.SetFile(file);
            item.SetText("shot.png");
            item.SetBitmap(new Avalonia.Media.Imaging.WriteableBitmap(new PixelSize(4, 4), new Vector(96, 96)));
            await SetClipboard(window, item);

            TextBox part = composer.PartsForTests[0];
            PutCaret(part, 0);
            part.Paste();

            await WaitUntil(() => composer.Pictures.Count == 1);
            Assert.Equal("shot.png", composer.Pictures[0].Name);
            Assert.Equal(OnePixelPng, await composer.Pictures[0].ReadBytesAsync());
        });
    }

    // A dropped file can be called "pasted-image-1.png" too: the next paste takes the next free
    // number, or the zip held two entries of one name and the mail showed the first at both markers.
    [Fact]
    public async Task A_pasted_picture_never_takes_the_name_of_a_dropped_one()
    {
        await UiTest.RunAsync(async () =>
        {
            using var workspace = new TempWorkspace();
            string dropped = Path.Combine(workspace.Root, "pasted-image-1.png");
            File.WriteAllBytes(dropped, OnePixelPng);

            var (_, composer) = Show();
            await composer.AddFilesAsync([dropped]);
            await composer.AddPastedBitmapAsync(new Avalonia.Media.Imaging.WriteableBitmap(new PixelSize(4, 4), new Vector(96, 96)));

            Assert.Equal(["pasted-image-1.png", "pasted-image-2.png"], composer.Pictures.Select(picture => picture.Name));
        });
    }

    // A marker cannot hold brackets, so "Screenshot [scope].png" was shown in CRT and never in the
    // mail. It travels as "Screenshot (scope).png", and its marker reads back to that name.
    [Fact]
    public async Task A_dropped_picture_with_brackets_in_its_name_travels_under_a_name_its_marker_can_carry()
    {
        await UiTest.RunAsync(async () =>
        {
            using var workspace = new TempWorkspace();
            string dropped = Path.Combine(workspace.Root, "Screenshot [scope].png");
            File.WriteAllBytes(dropped, OnePixelPng);

            var (_, composer) = Show();
            await composer.AddFilesAsync([dropped]);

            FeedbackPicture picture = Assert.Single(composer.Pictures);
            Assert.Equal("Screenshot (scope).png", picture.Name);
            Assert.True(FeedbackComposition.TryParseMarker(composer.ComposeText(), out string named));
            Assert.Equal(picture.Name, named);
        });
    }

    // A paste over a selection replaces it, as in any text box.
    [Fact]
    public void A_picture_pasted_over_selected_text_replaces_it()
    {
        UiTest.Run(() =>
        {
            var (_, composer) = Show();
            TextBox first = composer.PartsForTests[0];
            first.Text = "see PLACEHOLDER here";
            first.Focus();
            first.SelectionStart = 4;
            first.SelectionEnd = 15;
            Dispatcher.UIThread.RunJobs();

            composer.InsertPicture(Picture("a.png"));

            Assert.Equal(["see ", " here"], Texts(composer));
            Assert.Equal("see \n[a.png]\n here", composer.ComposeText());
        });
    }

    // A phone photo is stored on its side with an EXIF orientation saying so, and the mail's
    // reader turns it upright. The field shows it upright too, at the size it will have in the
    // mail - 4032 x 3024 stored, orientation 6, is a picture 3024 wide and 4032 tall.
    [Fact]
    public async Task A_dropped_photo_stored_on_its_side_has_the_size_it_is_shown_at_in_the_mail()
    {
        await UiTest.RunAsync(async () =>
        {
            using var workspace = new TempWorkspace();
            string photo = Path.Combine(workspace.Root, "board.jpg");
            File.WriteAllBytes(photo, JpegExifOrientationTests.Jpeg(width: 4032, height: 3024, orientation: 6, bigEndian: false));

            var (_, composer) = Show();
            await composer.AddFilesAsync([photo]);

            FeedbackPicture picture = Assert.Single(composer.Pictures);
            Assert.Equal(new PixelSize(3024, 4032), picture.NaturalSize);
            Assert.Equal(FeedbackComposition.MaxPictureWidth, picture.DisplayWidth);
            Assert.Equal(1307, picture.DisplayHeight);
        });
    }

    // A paste from the right-click menu may come while the menu has the focus: the picture still
    // goes where the part's caret is, as the menu's text paste does - not at the end.
    [Fact]
    public async Task A_picture_pasted_into_a_part_without_the_focus_goes_where_its_caret_is()
    {
        await UiTest.RunAsync(async () =>
        {
            var (window, composer, elsewhere) = ShowBeside();
            TextBox part = composer.PartsForTests[0];
            part.Text = "one two";
            PutCaret(part, 4);
            elsewhere.Focus();
            Dispatcher.UIThread.RunJobs();

            var item = new DataTransferItem();
            item.SetBitmap(new Avalonia.Media.Imaging.WriteableBitmap(new PixelSize(4, 4), new Vector(96, 96)));
            await SetClipboard(window, item);
            part.Paste();

            await WaitUntil(() => composer.Pictures.Count == 1);
            Assert.Equal(["one ", "two"], Texts(composer));
        });
    }

    // A paste still putting its picture in (the bitmap is mended off the UI thread first) is what
    // a send waits for - code review, 2026-10-16.
    [Fact]
    public async Task A_paste_still_putting_a_picture_in_is_waited_for()
    {
        await UiTest.RunAsync(async () =>
        {
            var (window, composer) = Show();
            var item = new DataTransferItem();
            item.SetBitmap(new Avalonia.Media.Imaging.WriteableBitmap(new PixelSize(4, 4), new Vector(96, 96)));
            await SetClipboard(window, item);

            TextBox part = composer.PartsForTests[0];
            PutCaret(part, 0);
            part.Paste();
            Assert.Empty(composer.Pictures);

            await composer.WhenSettledAsync();

            Assert.Single(composer.Pictures);
        });
    }

    // Code review, 2026-10-16: a click in the empty space beside a picture in the middle of the
    // text sent the caret to the very end. It goes to the line under that picture.
    [Fact]
    public void A_click_beside_a_picture_puts_the_caret_on_the_line_under_it()
    {
        UiTest.Run(() =>
        {
            var (window, composer) = Show();
            composer.PartsForTests[0].Text = "one";
            PutCaret(composer.PartsForTests[0], 3);
            composer.InsertPicture(Shown("a.png", 200, 100));
            composer.PartsForTests[1].Text = "two";
            PutCaret(composer.PartsForTests[1], 3);
            composer.InsertPicture(Shown("b.png", 200, 100));
            composer.PartsForTests[2].Text = "three";
            Dispatcher.UIThread.RunJobs();

            Image image = composer.PictureControlsForTests[0].Image;
            Point beside = image.TranslatePoint(new Point(image.Bounds.Width + 150, image.Bounds.Height / 2), window)!.Value;
            window.MouseDown(beside, MouseButton.Left);
            window.MouseUp(beside, MouseButton.Left);
            Dispatcher.UIThread.RunJobs();

            Assert.True(composer.PartsForTests[1].IsFocused);
            Assert.Equal(0, composer.PartsForTests[1].CaretIndex);
        });
    }

    // Below the last part, a click still puts the caret at the end of it.
    [Fact]
    public void A_click_below_the_text_puts_the_caret_at_its_end()
    {
        UiTest.Run(() =>
        {
            var (window, composer) = Show();
            composer.PartsForTests[0].Text = "one";
            PutCaret(composer.PartsForTests[0], 0);
            composer.InsertPicture(Shown("a.png", 200, 100));
            composer.PartsForTests[1].Text = "two";
            PutCaret(composer.PartsForTests[0], 0);

            Point below = composer.TranslatePoint(new Point(100, composer.Bounds.Height - 10), window)!.Value;
            window.MouseDown(below, MouseButton.Left);
            window.MouseUp(below, MouseButton.Left);
            Dispatcher.UIThread.RunJobs();

            Assert.True(composer.PartsForTests[1].IsFocused);
            Assert.Equal(3, composer.PartsForTests[1].CaretIndex);
        });
    }

    // The dispatcher run until `done` holds - a paste reads the clipboard, and a let-go bitmap is
    // disposed, asynchronously.
    private static async Task WaitUntil(Func<bool> done)
    {
        for (int attempt = 0; attempt < 300 && !done(); attempt++)
        {
            Dispatcher.UIThread.RunJobs();
            await Task.Delay(10);
        }

        Assert.True(done());
    }

    private static bool IsDisposed(Avalonia.Media.Imaging.Bitmap bitmap)
    {
        try
        {
            _ = bitmap.PixelSize;
            return false;
        }
        catch (ObjectDisposedException)
        {
            return true;
        }
    }

    private static async Task SetClipboard(Window window, DataTransferItem item)
    {
        var transfer = new DataTransfer();
        transfer.Add(item);
        await TopLevel.GetTopLevel(window)!.Clipboard!.SetDataAsync(transfer);
    }

    // The composer's border is the text box: a missing theme key would draw NO border, silently.
    [Fact]
    public void The_border_brushes_it_binds_to_exist_in_the_theme()
    {
        UiTest.Run(() =>
        {
            var (_, composer) = Show();

            // The resize handles use the focused border's brush and the field's background.
            foreach (string key in new[] { "TextControlBorderBrush", "TextControlBorderBrushFocused", "TextControlBorderThemeThickness", "TextControlBorderThemeThicknessFocused", "Form_Bg", "Feedback_TrashIcon_Fg" })
            {
                Assert.True(composer.TryFindResource(key, Avalonia.Styling.ThemeVariant.Light, out object? value) && value is not null, key);
            }

            Assert.NotNull(composer.GetControl<Border>("Frame").BorderBrush);
        });
    }
}
