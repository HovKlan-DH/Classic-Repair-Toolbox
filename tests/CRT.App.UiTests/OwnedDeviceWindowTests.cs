using Avalonia;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Threading;
using Avalonia.VisualTree;
using CRT;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// "Add device" and "Edit device" laid out as the worklog's editor is (owner request, 2026-11-06: "I think this modal
// needs to be more like the 'Worklog' modal, where it has these same possibilities: Name + Hardware + Board + Status +
// Region + Serial + 'Description' as-of today - Comments - Photos"), with its condition in stars ('can we add a
// "Condition" and then have 0 to 5 stars?') - OwnedDeviceWindow. An existing device's comments and photos are saved
// as they change (owner decision, 2026-11-06: "At once, as a worklog does"); a new device's wait for "Add device".
// The form's small windows - a comment, a photo, the viewer - stand in through its seams; the field rules, region
// and keys are WorkbooksDevicesTests'.
//
// HeadlessUi: it builds windows, and points WorklogManager - the Workbooks folder devices share - at a folder of
// its own per test.
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class OwnedDeviceWindowTests : IDisposable
{
    private const string Board = "Commodore 64|250407";

    private static readonly HardwareBoardEntry[] Boards = [new() { HardwareName = "Commodore 64", BoardName = "250407" }];

    private readonly TempWorkspace thisWorkspace = new();

    public OwnedDeviceWindowTests()
    {
        WorklogManager.LoadFrom(this.thisWorkspace.Path_("Workbooks-" + Guid.NewGuid().ToString("N")));
    }

    public void Dispose()
    {
        WorklogManager.LoadFrom(this.thisWorkspace.Path_("Detached-" + Guid.NewGuid().ToString("N")));
        this.thisWorkspace.Dispose();
    }

    private string Picture(string folder, string name)
    {
        string path = Path.Combine(this.thisWorkspace.Root, "Picked", folder, name);
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        File.WriteAllBytes(path, [1, 2, 3]);
        return path;
    }

    private static OwnedDeviceWindow NewForm()
    {
        var window = new OwnedDeviceWindow();
        window.Initialize(Boards, Board, boardRegions: _ => Task.FromResult<IReadOnlyList<string>>([]));
        return window;
    }

    private static OwnedDeviceWindow EditForm(OwnedDeviceRecord device)
    {
        var window = new OwnedDeviceWindow();
        window.InitializeForEdit(device, Boards, _ => Task.FromResult<IReadOnlyList<string>>([]));
        return window;
    }

    private static OwnedDeviceRecord AddDevice() =>
        OwnedDeviceStore.Create(new OwnedDeviceDetails("Breadbin #2", Board, "", "", "", OwnedDeviceStatus.Working))!;

    private static List<WorklogCommentRow> CommentRows(OwnedDeviceWindow window) =>
        window.GetControl<ItemsControl>("CommentsList").Items.OfType<WorklogCommentRow>().ToList();

    private static List<WorklogAttachmentRow> PhotoRows(OwnedDeviceWindow window) =>
        window.GetControl<ItemsControl>("PhotosList").Items.OfType<WorklogAttachmentRow>().ToList();

    private static int LitStars(OwnedDeviceWindow window) =>
        Assert.IsAssignableFrom<Panel>(window.GetControl<Decorator>("ConditionStarsHost").Child).Children.Count(star => star.Classes.Contains("Lit"));

    // ###########################################################################################
    // Laid out as the worklog's editor: what it is and its comments on the left, its photos on the right, a splitter
    // between - "Add comment" and "Add photo" green, as they add.
    // ###########################################################################################
    [Fact]
    public void Comments_stand_on_the_left_under_the_fields_and_photos_on_the_right()
    {
        UiTest.Run(() =>
        {
            OwnedDeviceWindow window = NewForm();
            window.Show();

            try
            {
                Dispatcher.UIThread.RunJobs();
                window.UpdateLayout();

                Button addComment = window.GetControl<Button>("AddCommentButton");
                Button addPhoto = window.GetControl<Button>("AddPhotoButton");
                TextBox description = window.GetControl<TextBox>("DescriptionTextBox");

                Assert.Contains("Accept", addComment.Classes);
                Assert.Contains("Accept", addPhoto.Classes);
                Assert.Single(window.GetVisualDescendants().OfType<GridSplitter>());

                double X(Control control) => control.TranslatePoint(default, window)!.Value.X;
                double Y(Control control) => control.TranslatePoint(default, window)!.Value.Y;

                Assert.True(X(addComment) < X(addPhoto), "\"Add comment\" is not left of \"Add photo\"");
                Assert.True(X(window.GetControl<TextBox>("NameTextBox")) < X(addPhoto));
                Assert.True(Y(description) < Y(addComment), "the comments are not under the fields");
                Assert.Equal("Comments", window.GetControl<TextBlock>("CommentsHeadingText").Text);
                Assert.True(window.GetControl<TextBlock>("NoCommentsText").IsVisible);
                Assert.True(window.GetControl<TextBlock>("NoPhotosText").IsVisible);
            }
            finally
            {
                window.Close();
            }
        });
    }

    // ###########################################################################################
    // AN EXISTING DEVICE'S COMMENTS ARE SAVED AS THEY CHANGE (owner decision, 2026-11-06): added, changed and deleted
    // in the worklog's own small window, newest first - and the form called off keeps them, saying the device changed.
    // ###########################################################################################
    [Fact]
    public async Task An_existing_devices_comments_are_saved_at_once_and_calling_the_form_off_keeps_them()
    {
        OwnedDeviceRecord device = AddDevice();

        await UiTest.RunAsync(async () =>
        {
            OwnedDeviceWindow window = EditForm(device);
            var answers = new Queue<string?>(["Recapped", "New PLA fitted", "PSU recapped"]);
            var asked = new List<string?>();
            window.AskCommentOverrideForTests = text =>
            {
                asked.Add(text);
                return Task.FromResult(answers.Dequeue());
            };

            await window.AddCommentAsync();
            Assert.True(window.ChangedOnDisk);
            Assert.Equal(["Recapped"], OwnedDeviceStore.Get(device.Id)!.Comments.Select(comment => comment.Text));

            await window.AddCommentAsync();
            Assert.Equal(2, CommentRows(window).Count);
            Assert.Equal("Comments (2)", window.GetControl<TextBlock>("CommentsHeadingText").Text);
            Assert.False(window.GetControl<TextBlock>("NoCommentsText").IsVisible);

            // Changed in the same window, on what it says.
            int first = OwnedDeviceStore.Get(device.Id)!.Comments.Single(comment => comment.Text == "Recapped").Id;
            await window.EditCommentAsync(first);
            Assert.Equal([null, null, "Recapped"], asked);
            Assert.Contains("PSU recapped", OwnedDeviceStore.Get(device.Id)!.Comments.Select(comment => comment.Text));

            window.DeleteComment(first);
            Assert.Equal(["New PLA fitted"], OwnedDeviceStore.Get(device.Id)!.Comments.Select(comment => comment.Text));
            Assert.Equal(["New PLA fitted"], CommentRows(window).Select(row => row.Text));

            // Called off: nothing of the fields saved, the comment kept.
            window.GetControl<TextBox>("NameTextBox").Text = "Renamed";
            window.Show();
            window.GetVisualDescendants().OfType<Button>().Single(button => Equals(button.Content, "Cancel"))
                .RaiseEvent(new RoutedEventArgs(Button.ClickEvent));

            Assert.Null(window.Saved);
            OwnedDeviceRecord kept = OwnedDeviceStore.Get(device.Id)!;
            Assert.Equal("Breadbin #2", kept.Name);
            Assert.Equal(["New PLA fitted"], kept.Comments.Select(comment => comment.Text));
        });
    }

    // ###########################################################################################
    // An existing device's photos are saved as they change too, each with its comment - the worklog's own "Add photo"
    // window - changed and deleted from its row.
    // ###########################################################################################
    [Fact]
    public async Task An_existing_devices_photos_are_added_changed_and_deleted_at_once()
    {
        OwnedDeviceRecord device = AddDevice();
        string front = this.Picture("a", "front.jpg");

        await UiTest.RunAsync(async () =>
        {
            OwnedDeviceWindow window = EditForm(device);
            var answers = new Queue<WorklogAddPhotoWindow.PhotoResult?>(
            [
                new WorklogAddPhotoWindow.PhotoResult(front, " Before cleaning "),
                new WorklogAddPhotoWindow.PhotoResult(null, "After cleaning"),
            ]);
            window.AskPhotoOverrideForTests = _ => Task.FromResult(answers.Dequeue());

            await window.AddPhotoAsync();

            WorklogAttachmentRecord photo = Assert.Single(OwnedDeviceStore.Get(device.Id)!.Photos);
            Assert.Equal(("front.jpg", "Before cleaning"), (photo.FileName, photo.Comment));
            Assert.Equal(["Before cleaning"], PhotoRows(window).Select(row => row.Comment));
            Assert.Equal("Photos (1)", window.GetControl<TextBlock>("PhotosHeadingText").Text);
            Assert.True(window.ChangedOnDisk);

            await window.EditPhotoAsync(photo.Id);
            Assert.Equal("After cleaning", Assert.Single(OwnedDeviceStore.Get(device.Id)!.Photos).Comment);

            window.DeletePhoto(photo.Id);
            Assert.Empty(OwnedDeviceStore.Get(device.Id)!.Photos);
            Assert.Empty(PhotoRows(window));
            Assert.True(window.GetControl<TextBlock>("NoPhotosText").IsVisible);
        });
    }

    // ###########################################################################################
    // A NEW DEVICE HAS NO FOLDER YET: the form holds its comments and photos - listed, changed and deleted there - and
    // "Add device" writes them with it, the photo copied in only then. Called off, nothing is on disk.
    // ###########################################################################################
    [Fact]
    public async Task A_new_devices_comments_and_photos_wait_for_Add_device()
    {
        string front = this.Picture("a", "front.jpg");
        string back = this.Picture("b", "back.jpg");

        await UiTest.RunAsync(async () =>
        {
            OwnedDeviceWindow window = NewForm();
            window.AskCommentOverrideForTests = _ => Task.FromResult<string?>("Bought at a flea market");
            var photos = new Queue<WorklogAddPhotoWindow.PhotoResult?>(
            [
                new WorklogAddPhotoWindow.PhotoResult(front, "Front"),
                new WorklogAddPhotoWindow.PhotoResult(back, "Back"),
            ]);
            window.AskPhotoOverrideForTests = _ => Task.FromResult(photos.Dequeue());

            await window.AddCommentAsync();
            await window.AddPhotoAsync();
            await window.AddPhotoAsync();
            window.DeletePhoto(PhotoRows(window).Single(row => row.Comment == "Back").Id);

            Assert.Equal(["Bought at a flea market"], CommentRows(window).Select(row => row.Text));
            Assert.Equal(["Front"], PhotoRows(window).Select(row => row.Comment));
            Assert.Empty(OwnedDeviceStore.GetAll());
            Assert.False(window.ChangedOnDisk);

            window.GetControl<TextBox>("NameTextBox").Text = "Breadbin #2";
            window.Show();
            window.GetControl<Button>("SaveButton").RaiseEvent(new RoutedEventArgs(Button.ClickEvent));

            // Its photo is copied in off the UI thread (code review, 2026-11-08) - the save waits for it.
            await window.SavingForTests;

            OwnedDeviceRecord added = Assert.Single(OwnedDeviceStore.GetAll());
            Assert.Equal(added.Id, window.Saved!.Id);
            Assert.Equal(["Bought at a flea market"], added.Comments.Select(comment => comment.Text));
            WorklogAttachmentRecord photo = Assert.Single(added.Photos);
            Assert.Equal(("front.jpg", "Front"), (photo.FileName, photo.Comment));
            Assert.NotNull(OwnedDeviceStore.PhotoPath(added.Id, "front.jpg"));
            Assert.Equal((1, 0), (window.PhotosAdded, window.PhotosNotAdded));
        });
    }

    // ###########################################################################################
    // *** PHOTOS ARE COPIED IN, AND THEIR THUMBNAILS READ, OFF THE UI THREAD (code review, 2026-11-08; CLAUDE.md: "Photos
    // are copied in, and a device deleted, off the UI thread under the BusyOverlay"). *** Six camera photos copied and
    // decoded on the UI thread froze CRT's window with nothing said. A gate held shut in each step proves it: run on the
    // UI thread, the step would hold the form until the gate gave up, and the form would answer nothing meanwhile -
    // here it lists the photo as being read, and its copy waits under the form's "please wait", which nothing closes.
    // ###########################################################################################
    [Fact]
    public async Task Photos_are_copied_in_and_their_thumbnails_read_off_the_UI_thread()
    {
        string front = this.Picture("a", "front.jpg");
        string back = this.Picture("b", "back.jpg");

        await UiTest.RunAsync(async () =>
        {
            using var reading = new ManualResetEventSlim();
            using var copying = new ManualResetEventSlim();

            OwnedDeviceWindow window = NewForm();
            window.WhileThumbnailIsReadForTests = () => reading.Wait(TimeSpan.FromSeconds(10));
            window.AskPhotoOverrideForTests = _ => Task.FromResult<WorklogAddPhotoWindow.PhotoResult?>(new(front, "Front"));

            await window.AddPhotoAsync();

            // Listed at once, its thumbnail still being read - neither picture nor "missing" shown meanwhile.
            WorklogAttachmentRow row = Assert.Single(PhotoRows(window));
            Assert.True(row.IsThumbnailPending);
            Assert.False(row.HasNoThumbnail);

            reading.Set();

            for (int attempt = 0; attempt < 300 && row.IsThumbnailPending; attempt++)
            {
                Dispatcher.UIThread.RunJobs();
                await Task.Delay(10);
            }

            Assert.False(row.IsThumbnailPending, "The thumbnail was not read in time");

            // Read, one of the two shows: the picture, or that there is none.
            Assert.NotEqual(row.HasThumbnail, row.HasNoThumbnail);

            // A new device's photo copied in as it is added - under the form's "please wait", which nothing closes.
            window.WhilePhotosAreCopiedForTests = () => copying.Wait(TimeSpan.FromSeconds(10));
            window.GetControl<TextBox>("NameTextBox").Text = "Breadbin #2";
            window.Show();

            Task saving = window.SaveAsync();

            Assert.False(saving.IsCompleted, "The photos were copied in on the UI thread");
            Assert.True(BusyOverlay.For(window)!.IsBusy);

            window.Close();
            Assert.True(window.IsVisible, "The form closed while its photos were copied in");

            copying.Set();
            await saving;

            Assert.False(BusyOverlay.For(window)!.IsBusy);
            Assert.False(window.IsVisible);
            Assert.Equal((1, 0), (window.PhotosAdded, window.PhotosNotAdded));
            WorklogAttachmentRecord photo = Assert.Single(Assert.Single(OwnedDeviceStore.GetAll()).Photos);
            Assert.Equal(("front.jpg", "Front"), (photo.FileName, photo.Comment));

            // An existing device's photo added: the same.
            OwnedDeviceWindow edit = EditForm(Assert.Single(OwnedDeviceStore.GetAll()));
            using var copyingAgain = new ManualResetEventSlim();
            edit.WhilePhotosAreCopiedForTests = () => copyingAgain.Wait(TimeSpan.FromSeconds(10));
            edit.AskPhotoOverrideForTests = _ => Task.FromResult<WorklogAddPhotoWindow.PhotoResult?>(new(back, "Back"));

            Task adding = edit.AddPhotoAsync();

            Assert.False(adding.IsCompleted, "The photo was copied in on the UI thread");
            Assert.True(BusyOverlay.For(edit)!.IsBusy);

            copyingAgain.Set();
            await adding;

            Assert.Equal(["Front", "Back"], Assert.Single(OwnedDeviceStore.GetAll()).Photos.Select(each => each.Comment));
        });
    }

    [Fact]
    public async Task A_new_device_called_off_leaves_nothing_on_disk()
    {
        string front = this.Picture("a", "front.jpg");

        await UiTest.RunAsync(async () =>
        {
            OwnedDeviceWindow window = NewForm();
            window.AskCommentOverrideForTests = _ => Task.FromResult<string?>("A comment");
            window.AskPhotoOverrideForTests = _ => Task.FromResult<WorklogAddPhotoWindow.PhotoResult?>(new(front, ""));

            await window.AddCommentAsync();
            await window.AddPhotoAsync();
            window.GetControl<TextBox>("NameTextBox").Text = "Breadbin #2";
            window.Show();
            window.GetVisualDescendants().OfType<Button>().Single(button => Equals(button.Content, "Cancel"))
                .RaiseEvent(new RoutedEventArgs(Button.ClickEvent));

            Assert.Null(window.Saved);
            Assert.Empty(OwnedDeviceStore.GetAll());
        });
    }

    // ###########################################################################################
    // *** ESCAPE IN AN OPEN DROP-DOWN CLOSES THE LIST, NEVER THE FORM (code review, 2026-11-08). *** The form's own keys
    // come first, and Escape cancelled it whatever had the keyboard: a list opened to pick a status, closed again with
    // Escape, threw away everything typed - for a new device its comments, links and photos held for "Add device" too,
    // with nothing asked. Enter already left an open list alone.
    // ###########################################################################################
    [Fact]
    public async Task Escape_in_an_open_drop_down_closes_the_list_and_the_form_stays()
    {
        await UiTest.RunAsync(async () =>
        {
            OwnedDeviceWindow window = NewForm();
            window.AskCommentOverrideForTests = _ => Task.FromResult<string?>("Bought at a flea market");
            await window.AddCommentAsync();
            window.Show();

            // The form puts the keyboard in its name box as it opens - settled first.
            Dispatcher.UIThread.RunJobs();

            try
            {
                ComboBox status = window.GetControl<ComboBox>("StatusComboBox");
                status.Focus();
                status.IsDropDownOpen = true;
                Dispatcher.UIThread.RunJobs();

                // Open, the list has the keyboard, on its item picked (ComboBox.TryFocusSelectedItem) - in the list's
                // popup, so the key goes there, as a person's would.
                Control picked = status.ContainerFromIndex(status.SelectedIndex)!;
                Assert.True(picked.IsFocused, "The open list's item picked does not have the keyboard");
                TopLevel list = TopLevel.GetTopLevel(picked)!;

                list.KeyPress(Key.Escape, RawInputModifiers.None, PhysicalKey.Escape, keySymbol: null);
                Dispatcher.UIThread.RunJobs();

                Assert.True(window.IsVisible, "Escape in an open list closed the form");
                Assert.False(status.IsDropDownOpen);
                Assert.Equal(["Bought at a flea market"], CommentRows(window).Select(row => row.Text));

                // With no list open, Escape cancels the form, as ever.
                window.KeyPress(Key.Escape, RawInputModifiers.None, PhysicalKey.Escape, keySymbol: null);
                Dispatcher.UIThread.RunJobs();

                Assert.False(window.IsVisible);
                Assert.Null(window.Saved);
            }
            finally
            {
                if (window.IsVisible)
                {
                    window.Close();
                }
            }
        });
    }

    // ###########################################################################################
    // A comment opened and closed with nothing changed is no change - its first line nested once too (code review,
    // 2026-11-08: compared with the stored comment trimmed whole, "  - recapped" never read as unchanged, and the device
    // was written again, the form saying it changed).
    // ###########################################################################################
    [Fact]
    public async Task A_comment_closed_unchanged_writes_nothing_its_first_line_nested_too()
    {
        OwnedDeviceRecord device = AddDevice();
        OwnedDeviceStore.AddComment(device.Id, "  - recapped", new DateTime(2026, 11, 8, 9, 0, 0));

        await UiTest.RunAsync(async () =>
        {
            OwnedDeviceWindow window = EditForm(OwnedDeviceStore.Get(device.Id)!);
            window.AskCommentOverrideForTests = text => Task.FromResult(text);

            await window.EditCommentAsync(Assert.Single(CommentRows(window)).Id);

            Assert.False(window.ChangedOnDisk);
        });
    }

    // A comment saved with its first line nested once is listed so, not trimmed to the first level (code review,
    // 2026-11-08).
    [Fact]
    public void A_comment_is_listed_with_the_indent_that_says_what_its_first_line_is()
    {
        OwnedDeviceRecord device = AddDevice();
        OwnedDeviceStore.AddComment(device.Id, "  - replaced C14\n- U8", new DateTime(2026, 11, 8, 9, 0, 0));

        UiTest.Run(() =>
        {
            OwnedDeviceWindow window = EditForm(OwnedDeviceStore.Get(device.Id)!);

            Assert.Equal(["  - replaced C14\n- U8"], CommentRows(window).Select(row => row.Text));
        });
    }

    // ###########################################################################################
    // *** A NOTE'S FIRST LINE KEEPS THE INDENT THAT SAYS WHAT IT IS (code review, 2026-11-08). *** The description was
    // saved as its markup trimmed, and a first line nested once - "  - recapped" - reopened as a bullet of the first
    // level. Blank lines around it and white space at its end still go.
    // ###########################################################################################
    [Fact]
    public void A_description_starting_with_a_nested_bullet_keeps_its_level()
    {
        UiTest.Run(() =>
        {
            OwnedDeviceWindow window = NewForm();
            window.GetControl<TextBox>("NameTextBox").Text = "Breadbin #2";
            window.GetControl<RichNoteTextBox>("DescriptionTextBox").Markup = "\n  - recapped\n- RAM\n  ";

            Assert.Equal("  - recapped\n- RAM", window.Details().Description);
        });
    }

    // ###########################################################################################
    // ITS CONDITION, 0 TO 5 STARS (owner request, 2026-11-06): a star clicked - by the pointer, on the star - sets
    // that many, the one set clicked again none; saved with the device, and shown again when it is edited.
    // ###########################################################################################
    [Fact]
    public void A_star_clicked_sets_the_condition_and_clicked_again_takes_it_away()
    {
        UiTest.Run(() =>
        {
            OwnedDeviceWindow window = NewForm();
            window.Show();

            try
            {
                Dispatcher.UIThread.RunJobs();
                window.UpdateLayout();
                Assert.Equal((0, 0), (window.Condition, LitStars(window)));

                Control third = Assert.IsAssignableFrom<Panel>(window.GetControl<Decorator>("ConditionStarsHost").Child)
                    .Children.Single(star => Equals(star.Tag, 3));
                Point at = third.TranslatePoint(new Point(third.Bounds.Width / 2, third.Bounds.Height / 2), window)!.Value;
                window.MouseDown(at, MouseButton.Left);
                window.MouseUp(at, MouseButton.Left);

                Assert.Equal((3, 3), (window.Condition, LitStars(window)));

                window.ClickStar(3);
                Assert.Equal((0, 0), (window.Condition, LitStars(window)));

                window.ClickStar(5);
                window.GetControl<TextBox>("NameTextBox").Text = "Breadbin #2";
                window.GetControl<Button>("SaveButton").RaiseEvent(new RoutedEventArgs(Button.ClickEvent));

                OwnedDeviceRecord saved = Assert.Single(OwnedDeviceStore.GetAll());
                Assert.Equal(5, saved.Condition);

                OwnedDeviceWindow edit = EditForm(saved);
                Assert.Equal((5, 5), (edit.Condition, LitStars(edit)));
            }
            finally
            {
                if (window.IsVisible)
                {
                    window.Close();
                }
            }
        });
    }

    private static List<WorklogLinkRow> LinkRows(OwnedDeviceWindow window) =>
        window.GetControl<ItemsControl>("LinksList").Items.OfType<WorklogLinkRow>().ToList();

    // ###########################################################################################
    // LINKS OF INTEREST, as a worklog has them (owner request, 2026-11-06): the worklog's own small window, added,
    // changed and deleted - an existing device's saved at once, a new device's with it - and a row clicked opens its
    // address.
    // ###########################################################################################
    [Fact]
    public async Task Links_of_interest_are_saved_at_once_on_a_device_and_with_a_new_one()
    {
        OwnedDeviceRecord device = AddDevice();

        await UiTest.RunAsync(async () =>
        {
            OwnedDeviceWindow edit = EditForm(device);
            var answers = new Queue<(string Headline, string Url)?>([("Service manual", "https://zimmers.net"), ("Manual", "https://zimmers.net/c64")]);
            var asked = new List<(string Headline, string Url)?>();
            edit.AskLinkOverrideForTests = link =>
            {
                asked.Add(link);
                return Task.FromResult(answers.Dequeue());
            };

            await edit.AddLinkAsync();
            Assert.True(edit.ChangedOnDisk);
            Assert.Equal([("Service manual", "https://zimmers.net")], OwnedDeviceStore.Get(device.Id)!.Links.Select(link => (link.Headline, link.Url)));
            Assert.Equal("Links of interest (1)", edit.GetControl<TextBlock>("LinksHeadingText").Text);
            Assert.False(edit.GetControl<TextBlock>("NoLinksText").IsVisible);

            int id = Assert.Single(LinkRows(edit)).Id;
            await edit.EditLinkAsync(id);
            Assert.Equal([null, ("Service manual", "https://zimmers.net")], asked);
            Assert.Equal("Manual", Assert.Single(OwnedDeviceStore.Get(device.Id)!.Links).Headline);

            edit.DeleteLink(id);
            Assert.Empty(OwnedDeviceStore.Get(device.Id)!.Links);
            Assert.True(edit.GetControl<TextBlock>("NoLinksText").IsVisible);

            // A new device's wait for "Add device".
            OwnedDeviceWindow added = NewForm();
            added.AskLinkOverrideForTests = _ => Task.FromResult<(string Headline, string Url)?>(("Wiki", "https://c64-wiki.com"));
            await added.AddLinkAsync();
            Assert.Equal(["Wiki"], LinkRows(added).Select(row => row.Headline));
            Assert.Single(OwnedDeviceStore.GetAll());

            added.GetControl<TextBox>("NameTextBox").Text = "Plus/4";
            added.Show();
            added.GetControl<Button>("SaveButton").RaiseEvent(new RoutedEventArgs(Button.ClickEvent));

            Assert.Equal([("Wiki", "https://c64-wiki.com")], OwnedDeviceStore.Get(added.Saved!.Id)!.Links.Select(link => (link.Headline, link.Url)));
        });
    }

    // A link's row clicked opens its address - the one way out of CRT, as a worklog's link.
    [Fact]
    public async Task A_link_row_clicked_opens_its_address()
    {
        OwnedDeviceRecord device = AddDevice();
        OwnedDeviceStore.AddLink(device.Id, "Service manual", "https://zimmers.net");

        await UiTest.RunAsync(async () =>
        {
            OwnedDeviceWindow window = EditForm(OwnedDeviceStore.Get(device.Id)!);
            var opened = new List<string>();
            window.OpenLinkOverrideForTests = opened.Add;
            window.Show();

            try
            {
                Dispatcher.UIThread.RunJobs();
                window.UpdateLayout();
                await Task.Yield();

                Border row = window.GetControl<ItemsControl>("LinksList").GetVisualDescendants().OfType<Border>()
                    .First(border => Equals(border.Tag, "https://zimmers.net"));
                Point at = row.TranslatePoint(new Point(row.Bounds.Width / 2, row.Bounds.Height / 2), window)!.Value;
                window.MouseDown(at, MouseButton.Left);
                window.MouseUp(at, MouseButton.Left);

                Assert.Equal(["https://zimmers.net"], opened);
            }
            finally
            {
                window.Close();
            }
        });
    }

    // A link typed without its scheme opens as the web address it is (code review, 2026-11-08: tried as a file, refused).
    [Fact]
    public async Task A_link_row_typed_without_its_scheme_opens_as_a_web_address()
    {
        OwnedDeviceRecord device = AddDevice();
        OwnedDeviceStore.AddLink(device.Id, "Wiki", "www.c64-wiki.com");

        await UiTest.RunAsync(async () =>
        {
            OwnedDeviceWindow window = EditForm(OwnedDeviceStore.Get(device.Id)!);
            var opened = new List<string>();
            window.OpenLinkOverrideForTests = opened.Add;
            window.Show();

            try
            {
                Dispatcher.UIThread.RunJobs();
                window.UpdateLayout();
                await Task.Yield();

                Border row = window.GetControl<ItemsControl>("LinksList").GetVisualDescendants().OfType<Border>()
                    .First(border => Equals(border.Tag, "www.c64-wiki.com"));
                Point at = row.TranslatePoint(new Point(row.Bounds.Width / 2, row.Bounds.Height / 2), window)!.Value;
                window.MouseDown(at, MouseButton.Left);
                window.MouseUp(at, MouseButton.Left);

                Assert.Equal(["https://www.c64-wiki.com"], opened);
            }
            finally
            {
                window.Close();
            }
        });
    }

    // ###########################################################################################
    // THE PHOTOS ARE DRAGGED INTO ANOTHER ORDER, as a worklog's are (owner request, 2026-11-06: "I would like to be able to
    // rearrange the photos, just like it is possible to do in the worklogs"): by the space beside the thumbnail, a dashed
    // slot where it lands, the order saved at once on a device.
    // ###########################################################################################
    [Fact]
    public void A_photo_dragged_up_takes_its_new_place_and_the_order_is_saved()
    {
        OwnedDeviceRecord device = AddDevice();
        OwnedDeviceStore.AddPhotos(device.Id,
        [
            new OwnedDevicePhotoSource(this.Picture("a", "one.jpg"), "One"),
            new OwnedDevicePhotoSource(this.Picture("a", "two.jpg"), "Two"),
            new OwnedDevicePhotoSource(this.Picture("a", "three.jpg"), "Three"),
        ]);

        UiTest.Run(() =>
        {
            OwnedDeviceWindow window = EditForm(OwnedDeviceStore.Get(device.Id)!);
            window.Show();

            try
            {
                Dispatcher.UIThread.RunJobs();
                window.UpdateLayout();

                ItemsControl list = window.GetControl<ItemsControl>("PhotosList");
                // The handle: the panel beside the thumbnail, in the row's second column.
                List<Panel> handles = list.GetVisualDescendants().OfType<Panel>()
                    .Where(panel => panel.DataContext is WorklogAttachmentRow && panel.Parent is Grid && Grid.GetColumn(panel) == 1)
                    .ToList();
                Assert.Equal(3, handles.Count);

                Control third = handles[2];
                Point from = third.TranslatePoint(new Point(third.Bounds.Width / 2, third.Bounds.Height - 4), window)!.Value;
                Control first = list.ContainerFromIndex(0)!;
                Point top = first.TranslatePoint(new Point(10, 2), window)!.Value;

                window.MouseDown(from, MouseButton.Left);
                Dispatcher.UIThread.RunJobs();
                window.MouseMove(top, RawInputModifiers.LeftMouseButton);
                Dispatcher.UIThread.RunJobs();
                Assert.True(window.IsDraggingPhotoForTests);
                window.MouseUp(top, MouseButton.Left);
                Dispatcher.UIThread.RunJobs();

                Assert.Equal(["Three", "One", "Two"], PhotoRows(window).Select(row => row.Comment));
                Assert.Equal(["three.jpg", "one.jpg", "two.jpg"], OwnedDeviceStore.Get(device.Id)!.Photos
                    .OrderBy(photo => photo.DisplayOrder).Select(photo => photo.FileName));
                Assert.True(window.ChangedOnDisk);
            }
            finally
            {
                window.Close();
            }
        });
    }

    // ###########################################################################################
    // HARDWARE AND BOARD ARE LOCKED WHILE THE DEVICE HAS WORKBOOKS (owner decision, 2026-11-06, asked: "Locked while it has
    // workbooks" - a workbook is associated only to a device of its own board), saying why; with none they change.
    // ###########################################################################################
    [Fact]
    public void Hardware_and_board_are_locked_while_the_device_has_workbooks()
    {
        OwnedDeviceRecord device = AddDevice();

        UiTest.Run(() =>
        {
            OwnedDeviceWindow locked = new();
            locked.InitializeForEdit(device, Boards, _ => Task.FromResult<IReadOnlyList<string>>([]), workbookCount: 2);

            Assert.False(locked.GetControl<ComboBox>("HardwareComboBox").IsEnabled);
            Assert.False(locked.GetControl<ComboBox>("BoardComboBox").IsEnabled);
            Assert.True(locked.GetControl<TextBlock>("BoardLockedText").IsVisible);
            Assert.Equal(OwnedDeviceWording.BoardLocked(2), locked.GetControl<TextBlock>("BoardLockedText").Text);

            OwnedDeviceWindow free = EditForm(device);
            Assert.True(free.GetControl<ComboBox>("HardwareComboBox").IsEnabled);
            Assert.True(free.GetControl<ComboBox>("BoardComboBox").IsEnabled);
            Assert.False(free.GetControl<TextBlock>("BoardLockedText").IsVisible);
        });
    }

    // The stars light up to the one under the pointer, as a rating does, and back to the condition set as it leaves.
    [Fact]
    public void The_stars_light_up_to_the_one_under_the_pointer()
    {
        UiTest.Run(() =>
        {
            OwnedDeviceWindow window = NewForm();
            window.Show();

            try
            {
                Dispatcher.UIThread.RunJobs();
                window.UpdateLayout();
                window.ClickStar(2);

                Panel stars = Assert.IsAssignableFrom<Panel>(window.GetControl<Decorator>("ConditionStarsHost").Child);
                Control fourth = stars.Children.Single(star => Equals(star.Tag, 4));
                Point over = fourth.TranslatePoint(new Point(fourth.Bounds.Width / 2, fourth.Bounds.Height / 2), window)!.Value;

                window.MouseMove(over);
                Assert.Equal(4, LitStars(window));

                window.MouseMove(new Point(2, 2));
                Assert.Equal(2, LitStars(window));
                Assert.Equal(2, window.Condition);
            }
            finally
            {
                window.Close();
            }
        });
    }

    // A photo's thumbnail clicked opens it whole in the worklog's viewer, which zooms (owner request, 2026-11-06) -
    // a new device's from where it was picked.
    [Fact]
    public async Task A_photo_clicked_opens_the_viewer_on_it()
    {
        string front = this.Picture("a", "front.jpg");

        await UiTest.RunAsync(async () =>
        {
            OwnedDeviceWindow window = NewForm();
            window.AskPhotoOverrideForTests = _ => Task.FromResult<WorklogAddPhotoWindow.PhotoResult?>(new(front, "The front"));
            WorklogPhotoViewerWindow? shown = null;
            window.ShowPhotoViewerOverrideForTests = viewer => shown = viewer;

            await window.AddPhotoAsync();
            window.ViewPhoto(Assert.Single(PhotoRows(window)).Id);

            Assert.NotNull(shown);
            Assert.Equal("front.jpg", shown!.Title);
            Assert.Equal("The front", shown.GetControl<TextBlock>("CommentText").Text);
        });
    }
}
