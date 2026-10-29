using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Layout;
using Avalonia.Media;
using Avalonia.Media.Imaging;
using Avalonia.Platform.Storage;
using Avalonia.Threading;
using Handlers.DataHandling;
using Handlers.Theming;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace CRT
{
    // ###########################################################################################
    // "MY DEVICES" ON THE WORKBOOKS TAB (CRT 4.0 step 4 - the plan: "Your own machines, each with
    // its repairs"; called a Device wherever a person reads it). The computers the user owns are
    // listed above the workbooks, every one whatever board is shown; a click shows the device's
    // PAGE in place of the workbook:
    //
    //   - what it is - name, status, hardware and board, region, serial number, where it came from -
    //     with "Edit device" and "Delete device" (its repairs stay, on no device);
    //   - its REPAIRS - the workbooks naming it, on any board - with their worklogs' time and cost
    //     added up (OwnedDeviceRepairs), each opening its workbook, and "New repair", which creates
    //     one on the device's board and opens it;
    //   - the CHANGES made to it outside a repair, with a line to add one (a board swapped in "Edit
    //     device" is recorded there by itself);
    //   - its NOTES (saved when the box is left) and PHOTOS.
    //
    // A workbook card clicked, or a repair opened, brings the workbook back. Under a workbook's
    // title "On my device" puts it on a device with its board, or takes it off - existing workbooks
    // stay as they are until someone does (the plan: "Nothing is forced").
    //
    // Everything is read and written through OwnedDeviceStore and WorklogManager; the rules are
    // OwnedDeviceRepairs' and the words OwnedDeviceWording's.
    //
    // Part of the TabWorkbooks partial class - see TabWorkbooks.axaml.cs for the file map.
    // ###########################################################################################
    public partial class TabWorkbooks
    {
        // The device whose page is shown in place of the workbook - null while the workbook is.
        private int? thisSelectedDeviceId;

        // The devices as the last refresh read them.
        private List<OwnedDeviceRecord> thisDevices = new();

        // Set while the code fills "On my device", so its own selection is no pick.
        private bool thisApplyingWorkbookDevice;

        // ###########################################################################################
        // WHOSE NOTES THE NOTES BOX HOLDS - null while it holds none (code review, 2026-10-28). A
        // device card is a plain border, and clicking one takes no focus from the box: the next
        // device's page kept the first one's typed notes (the box was not refilled under the caret),
        // and leaving the box saved them as the SECOND device's. They are saved to the device they
        // were typed for, before the box shows another's.
        // ###########################################################################################
        private int? thisNotesDeviceId;

        // ###########################################################################################
        // WHAT A REFRESH NEED NOT DO AGAIN while a device's page is open (code review, 2026-10-28 -
        // each search keystroke and worklog save re-read every workbook and decoded every photo):
        //   - every workbook, read ONCE per refresh pass (null: not read yet this pass);
        //   - each photo decoded once, kept by its path, length and write time - a photo replaced is a
        //     new one; never disposed by hand (one still drawn takes the render thread down);
        //   - the board pane behind the page, rebuilt only once the page goes.
        // ###########################################################################################
        private List<WorkbookRecord>? thisAllWorkbooksThisPass;

        private readonly Dictionary<string, Task<Bitmap?>> thisDevicePhotos = new(StringComparer.OrdinalIgnoreCase);

        private bool thisBoardPreviewsWaitForPage;

        private List<WorkbookRecord> AllWorkbooksThisPass => this.thisAllWorkbooksThisPass ??= WorklogManager.GetAllWorkbooks();

        // The board pane is rebuilt by a refresh only while it is on screen - else when the page goes.
        private bool ShouldRebuildBoardPreviews()
        {
            this.thisBoardPreviewsWaitForPage = this.thisSelectedDeviceId is not null;
            return !this.thisBoardPreviewsWaitForPage;
        }

        // Stand-ins for the windows and the picker the device screens open - null in the app.
        internal Func<Task<IReadOnlyList<string>>>? PickDevicePhotosOverrideForTests { get; set; }

        internal Func<OwnedDeviceRecord, int, Task<bool>>? ConfirmDeleteDeviceOverrideForTests { get; set; }

        internal int? SelectedDeviceIdForTests => this.thisSelectedDeviceId;

        internal bool IsDevicePageShown => this.DevicePageScroller?.IsVisible == true;

        // ###########################################################################################
        // "My devices" drawn - every device a card, the one whose page is shown marked - from the
        // devices the caller read (Main.RefreshWorklogBar's), else read again. A device that has gone
        // (deleted elsewhere) takes its page with it.
        // ###########################################################################################
        private void RefreshDeviceList(List<OwnedDeviceRecord>? devices = null)
        {
            if (this.DeviceListPanel == null)
            {
                return;
            }

            this.thisDevices = devices ?? OwnedDeviceStore.GetAll();

            if (this.thisSelectedDeviceId is { } shown && this.thisDevices.All(device => device.Id != shown))
            {
                this.thisSelectedDeviceId = null;
            }

            this.DeviceListPanel.Children.Clear();

            foreach (OwnedDeviceRecord device in this.thisDevices)
            {
                this.DeviceListPanel.Children.Add(this.BuildDeviceCard(device, device.Id == this.thisSelectedDeviceId));
            }

            this.DeviceListHeadingText.Text = OwnedDeviceWording.ListHeading;
            this.NoDevicesText.Text = OwnedDeviceWording.NoDevicesYet;
            this.NoDevicesText.IsVisible = this.thisDevices.Count == 0;
        }

        // ###########################################################################################
        // The right-hand side as it should be now: the device's page while one is picked, else the
        // workbook - called at the end of every refresh.
        // ###########################################################################################
        private void ApplyDeviceView()
        {
            if (this.DevicePageScroller == null || this.WorkbookViewGrid == null)
            {
                return;
            }

            OwnedDeviceRecord? device = this.thisDevices.FirstOrDefault(each => each.Id == this.thisSelectedDeviceId);

            // The page going: the notes typed on it saved first, to the device they were typed for.
            if (device is null)
            {
                this.SaveDeviceNotes();
                this.thisNotesDeviceId = null;
            }

            this.DevicePageScroller.IsVisible = device is not null;
            this.WorkbookViewGrid.IsVisible = device is null;

            // The board pane, not rebuilt behind the page, rebuilt as it comes back.
            if (device is null && this.thisBoardPreviewsWaitForPage)
            {
                this.thisBoardPreviewsWaitForPage = false;
                this.RefreshBoardPreviews();
            }

            // One card marked: the device's while its page is shown, else the workbook's shown.
            foreach (Control card in this.WorkbookListPanel.Children.OfType<Control>())
            {
                card.Classes.Set("Selected", device is null && card.Tag is WorkbookRecord workbook && workbook.Id == this.thisSelectedWorkbookId);
            }

            if (device is not null)
            {
                this.FillDevicePage(device);
            }
        }

        // A device card clicked: its page, in place of the workbook.
        internal void SelectDevice(int deviceId)
        {
            // What "Add photos" said was of the device shown before.
            if (this.thisSelectedDeviceId != deviceId)
            {
                this.ShowDevicePhotosNote(null);
            }

            this.thisSelectedDeviceId = deviceId;
            this.ShowDeviceAgain();
        }

        // ###########################################################################################
        // One card in "My devices": the name with its status, the board under it.
        // ###########################################################################################
        private Border BuildDeviceCard(OwnedDeviceRecord device, bool isSelected)
        {
            var name = new TextBlock
            {
                Text = device.Name,
                FontSize = 11,
                FontWeight = FontWeight.Bold,
                TextWrapping = TextWrapping.Wrap,
                VerticalAlignment = VerticalAlignment.Center,
            };

            Border status = WorklogInfoPillBuilder.BuildDeviceStatusPill(device.Status, fontSize: 10);
            status.HorizontalAlignment = HorizontalAlignment.Right;

            var header = new Grid { ColumnDefinitions = new ColumnDefinitions("*,Auto") };
            Grid.SetColumn(status, 1);
            header.Children.Add(name);
            header.Children.Add(status);

            var board = new TextBlock
            {
                Text = BoardTextOf(device.BoardKey),
                FontSize = 10,
                TextWrapping = TextWrapping.Wrap,
            };
            board.Classes.Add("WorkbooksMuted");

            var body = new StackPanel { Spacing = 3 };
            body.Children.Add(header);
            body.Children.Add(board);

            var card = new Border { Tag = device, Child = body };
            card.Classes.Add("WorkbooksJobCard");

            if (isSelected)
            {
                card.Classes.Add("Selected");
            }

            card.PointerPressed += (_, _) => this.SelectDevice(device.Id);
            return card;
        }

        // A board key as the two drop-downs name it - "Commodore 64 - 250407 (long board)".
        private static string BoardTextOf(string boardKey)
        {
            (string hardware, string board) = Main.FormatBoardKeyAsFullNames(boardKey);

            return string.IsNullOrWhiteSpace(board)
                ? OwnedDeviceWording.BoardText(boardKey)
                : $"{hardware} - {board}";
        }

        // ###########################################################################################
        // THE DEVICE'S PAGE, filled from the device and every workbook naming it.
        // ###########################################################################################
        private void FillDevicePage(OwnedDeviceRecord device)
        {
            this.DeviceNameText.Text = device.Name;
            this.DeviceStatusHost.Child = WorklogInfoPillBuilder.BuildDeviceStatusPill(device.Status, fontSize: 11);

            this.DeviceDetailText.Text = OwnedDeviceWording.DetailLine(device, BoardTextOf(device.BoardKey));
            this.DeviceOriginText.Text = OwnedDeviceWording.OriginLine(device.Origin);
            this.DeviceOriginText.IsVisible = this.DeviceOriginText.Text.Length > 0;

            this.FillDeviceRepairs(device);
            this.FillDeviceChanges(device);
            this.FillDevicePhotos(device);
        }

        private void FillDeviceRepairs(OwnedDeviceRecord device)
        {
            List<WorkbookRecord> repairs = OwnedDeviceRepairs.Of(device.Id, this.AllWorkbooksThisPass);
            OwnedDeviceRepairs.History history = OwnedDeviceRepairs.Summarize(repairs, this.GetEntriesForThisPass);

            this.DeviceRepairsHeadingText.Text = OwnedDeviceWording.RepairsHeading;
            ApplyStatRuns(this.DeviceRepairsSummaryText, OwnedDeviceRepairs.HeadlineStats(history, UserSettings.WorklogCurrencyCode));
            this.DeviceRepairsSummaryText.IsVisible = repairs.Count > 0;

            this.NoDeviceRepairsText.Text = OwnedDeviceWording.NoRepairsYet;
            this.NoDeviceRepairsText.IsVisible = repairs.Count == 0;
            this.NewRepairButton.IsEnabled = device.BoardKey.Trim().Length > 0;

            this.DeviceRepairsPanel.Children.Clear();

            foreach (WorkbookRecord repair in repairs)
            {
                this.DeviceRepairsPanel.Children.Add(this.BuildDeviceRepairRow(repair, device));
            }
        }

        // ###########################################################################################
        // One repair on a device's page: its state, "#12 No video", how many worklogs and since when -
        // and its board when the device has another one now (a board swapped since).
        // ###########################################################################################
        private Border BuildDeviceRepairRow(WorkbookRecord repair, OwnedDeviceRecord device)
        {
            var title = new TextBlock
            {
                Text = $"#{repair.Id} · {(string.IsNullOrWhiteSpace(repair.Title) ? "(untitled)" : repair.Title)}",
                FontSize = 11,
                FontWeight = FontWeight.SemiBold,
                TextWrapping = TextWrapping.Wrap,
                VerticalAlignment = VerticalAlignment.Center,
            };

            Border state = WorklogInfoPillBuilder.BuildStatePill(repair.Status, fontSize: 10.0);
            state.HorizontalAlignment = HorizontalAlignment.Right;

            var header = new Grid { ColumnDefinitions = new ColumnDefinitions("*,Auto") };
            Grid.SetColumn(state, 1);
            header.Children.Add(title);
            header.Children.Add(state);

            var body = new StackPanel { Spacing = 2 };
            body.Children.Add(header);

            var meta = new TextBlock { Text = BuildWorkbookMetaText(repair), FontSize = 10, TextWrapping = TextWrapping.Wrap };
            meta.Classes.Add("WorkbooksMuted");
            body.Children.Add(meta);

            if (!string.Equals(repair.BoardKey, device.BoardKey, StringComparison.OrdinalIgnoreCase))
            {
                var board = new TextBlock { Text = "On " + BoardTextOf(repair.BoardKey), FontSize = 10, TextWrapping = TextWrapping.Wrap };
                board.Classes.Add("WorkbooksMuted");
                body.Children.Add(board);
            }

            var row = new Border { Tag = repair, Child = body };
            row.Classes.Add("WorkbooksJobCard");
            row.PointerPressed += (_, _) => this.OpenRepair(repair);
            return row;
        }

        // ###########################################################################################
        // A repair opened from a device's page: the workbook shown and activated, on its own board -
        // the same path as clicking its card.
        // ###########################################################################################
        internal void OpenRepair(WorkbookRecord repair)
        {
            this.thisSelectedDeviceId = null;
            this.SelectWorkbook(repair.Id, repair.BoardKey);
            this.ShowDeviceAgain();
        }

        private void FillDeviceChanges(OwnedDeviceRecord device)
        {
            this.DeviceChangesHeadingText.Text = OwnedDeviceWording.ChangesHeading(device.Changes.Count);
            this.NoDeviceChangesText.Text = OwnedDeviceWording.NoChangesYet;
            this.NoDeviceChangesText.IsVisible = device.Changes.Count == 0;

            this.ChangeDatePicker.SelectedDate ??= DateTime.Now.Date;

            this.DeviceChangesPanel.Children.Clear();

            foreach (OwnedDeviceChangeRecord change in device.Changes.OrderBy(change => change.Date).ThenBy(change => change.Id))
            {
                var date = new TextBlock { Text = OwnedDeviceWording.ChangeDate(change.Date), FontSize = 11, VerticalAlignment = VerticalAlignment.Center };
                date.Classes.Add("WorkbooksMuted");

                var text = new TextBlock { Text = change.Text, FontSize = 11, TextWrapping = TextWrapping.Wrap, VerticalAlignment = VerticalAlignment.Center };

                var remove = new Button { Content = "Remove", FontSize = 10, Padding = new Thickness(6, 1), Tag = change };
                remove.Classes.Add("Danger");
                remove.Click += (_, _) => this.RemoveDeviceChange(change.Id);

                var row = new Grid { ColumnDefinitions = new ColumnDefinitions("80,8,*,8,Auto") };
                Grid.SetColumn(text, 2);
                Grid.SetColumn(remove, 4);
                row.Children.Add(date);
                row.Children.Add(text);
                row.Children.Add(remove);

                this.DeviceChangesPanel.Children.Add(row);
            }
        }

        private void FillDevicePhotos(OwnedDeviceRecord device)
        {
            this.DevicePhotosHeadingText.Text = OwnedDeviceWording.PhotosHeading(device.Photos.Count);

            // Another device's notes in the box: saved to THEIR device, then this one's shown. This
            // device's own: never refilled under the caret - a refresh would throw the typing away.
            if (this.thisNotesDeviceId != device.Id)
            {
                this.SaveDeviceNotes();
                this.DeviceNotesTextBox.Text = device.Notes;
                this.thisNotesDeviceId = device.Id;
            }
            else if (!this.DeviceNotesTextBox.IsFocused)
            {
                this.DeviceNotesTextBox.Text = device.Notes;
            }

            this.NoDevicePhotosText.Text = OwnedDeviceWording.NoPhotosYet;
            this.NoDevicePhotosText.IsVisible = device.Photos.Count == 0;

            this.DevicePhotosPanel.Children.Clear();

            foreach (WorklogAttachmentRecord photo in device.Photos.OrderBy(photo => photo.DisplayOrder).ThenBy(photo => photo.Id))
            {
                var image = new Image { Width = 160, Height = 120, Stretch = Stretch.Uniform };
                this.ShowDevicePhoto(image, OwnedDeviceStore.PhotoPath(device.Id, photo.FileName));

                var remove = new Button { Content = "Remove", FontSize = 10, Padding = new Thickness(6, 1), HorizontalAlignment = HorizontalAlignment.Right, Tag = photo };
                remove.Classes.Add("Danger");
                int deviceId = device.Id;
                remove.Click += (_, _) => this.RemoveDevicePhoto(deviceId, photo.Id);

                var tile = new StackPanel { Spacing = 3, Tag = photo };
                tile.Children.Add(new Border
                {
                    BorderBrush = ThemeResources.ResolveBrush("Workbooks_Separator", Brushes.Gray),
                    BorderThickness = new Thickness(1),
                    Child = image,
                });
                tile.Children.Add(remove);

                this.DevicePhotosPanel.Children.Add(tile);
            }
        }

        // ###########################################################################################
        // A photo decoded small OFF the UI thread and shown once read - ONCE per photo, kept by where it
        // lies, its length and its write time (thisDevicePhotos); never disposed by hand (one still
        // drawn takes the render thread down). A photo that cannot be read leaves its frame empty.
        // ###########################################################################################
        private void ShowDevicePhoto(Image image, string? path)
        {
            if (path is null)
            {
                return;
            }

            string key;

            try
            {
                var file = new FileInfo(path);
                key = $"{path}|{file.Length}|{file.LastWriteTimeUtc.Ticks}";
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                return;
            }

            if (!this.thisDevicePhotos.TryGetValue(key, out Task<Bitmap?>? load))
            {
                load = Task.Run(() => DecodeDevicePhoto(path));
                this.thisDevicePhotos[key] = load;
            }

            if (load.IsCompletedSuccessfully)
            {
                image.Source = load.Result;
                return;
            }

            _ = load.ContinueWith(
                done => Dispatcher.UIThread.Post(() => image.Source = done.Result),
                TaskContinuationOptions.OnlyOnRanToCompletion);
        }

        private static Bitmap? DecodeDevicePhoto(string path)
        {
            try
            {
                using FileStream stream = File.OpenRead(path);
                return Bitmap.DecodeToWidth(stream, 320);
            }
            catch (Exception ex)
            {
                Logger.Warning($"Could not show a device photo: [{path}] [{ex.Message}]");
                return null;
            }
        }

        // How many photos have been decoded since the tab was built - for the tests.
        internal int DevicePhotosDecodedForTests => this.thisDevicePhotos.Count;

        // ---- what the page's buttons do ------------------------------------------------------

        // After any change to a device: the list, the page and the worklog bar as they now are.
        private void RefreshAfterDeviceChange()
        {
            if (this.MainWindow is { } main)
            {
                main.RefreshWorklogBar();
            }
            else
            {
                this.RefreshWorkbooks();
            }
        }

        private IReadOnlyList<HardwareBoardEntry> BoardsForDevices =>
            this.BoardsOverrideForTests ?? DataManager.HardwareBoards;

        // The boards a device can be on - DataManager's, or a test's.
        internal IReadOnlyList<HardwareBoardEntry>? BoardsOverrideForTests { get; set; }

        private async void OnAddDeviceClick(object? sender, RoutedEventArgs e)
        {
            // Nothing may leave an async void handler: App lets such an exception end CRT.
            try
            {
                if (TopLevel.GetTopLevel(this) is not Window owner)
                {
                    return;
                }

                var dialog = new OwnedDeviceWindow();
                dialog.Initialize(this.BoardsForDevices, this.BoardKeyOverrideForTests ?? this.MainWindow?.GetCurrentBoardKey());

                if (await dialog.ShowDialog<OwnedDeviceRecord?>(owner) is { } added)
                {
                    this.thisSelectedDeviceId = added.Id;
                    this.RefreshAfterDeviceChange();
                }
            }
            catch (Exception ex)
            {
                Logger.Critical($"Adding a device failed - [{ex}]");
            }
        }

        private async void OnEditDeviceClick(object? sender, RoutedEventArgs e)
        {
            try
            {
                if (TopLevel.GetTopLevel(this) is not Window owner || this.SelectedDevice is not { } device)
                {
                    return;
                }

                var dialog = new OwnedDeviceWindow();
                dialog.InitializeForEdit(device, this.BoardsForDevices);

                if (await dialog.ShowDialog<OwnedDeviceRecord?>(owner) is not null)
                {
                    this.RefreshAfterDeviceChange();
                }
            }
            catch (Exception ex)
            {
                Logger.Critical($"Editing a device failed - [{ex}]");
            }
        }

        private async void OnDeleteDeviceClick(object? sender, RoutedEventArgs e)
        {
            try
            {
                await this.DeleteSelectedDeviceAsync();
            }
            catch (Exception ex)
            {
                Logger.Critical($"Deleting a device failed - [{ex}]");
            }
        }

        // ###########################################################################################
        // The device whose page is shown deleted, once confirmed - its repairs stay, on no device,
        // and the workbook comes back in its place. Said when its folder could not be removed.
        // ###########################################################################################
        internal async Task DeleteSelectedDeviceAsync()
        {
            if (this.SelectedDevice is not { } device)
            {
                return;
            }

            int repairs = OwnedDeviceRepairs.Of(device.Id, WorklogManager.GetAllWorkbooks()).Count;

            if (!await this.ConfirmDeleteDeviceAsync(device, repairs))
            {
                return;
            }

            // Off the UI thread under CRT's please-wait - a folder of camera photos takes a while.
            if (!await BusyOverlay.RunLocalAsync(this, CrtWaitWording.DeletingDevice, () => Task.Run(() => OwnedDeviceStore.Delete(device.Id))))
            {
                if (TopLevel.GetTopLevel(this) is Window owner)
                {
                    await ShowWorkbookActionFailedAsync(
                        owner,
                        "Delete device",
                        $"Could not delete [{device.Name}] - see the log for details.\n\nOne of its photos may be open in another program.");
                }

                return;
            }

            this.thisSelectedDeviceId = null;
            this.RefreshAfterDeviceChange();
        }

        private async Task<bool> ConfirmDeleteDeviceAsync(OwnedDeviceRecord device, int repairs)
        {
            if (this.ConfirmDeleteDeviceOverrideForTests is { } confirm)
            {
                return await confirm(device, repairs);
            }

            if (TopLevel.GetTopLevel(this) is not Window owner)
            {
                return false;
            }

            var dialog = new DeleteDeviceWindow();
            dialog.Initialize(device, repairs);
            return await dialog.ShowDialog<bool?>(owner) == true;
        }

        private OwnedDeviceRecord? SelectedDevice =>
            this.thisSelectedDeviceId is { } id ? OwnedDeviceStore.Get(id) : null;

        // ###########################################################################################
        // "New repair": a workbook on the device's board, on the device - "Create new workbook"'s own
        // form, the device preselected - then opened, its board selected (the plan: "New repair on a
        // machine creates a workbook for that machine's board, and selects the board").
        // ###########################################################################################
        private async void OnNewRepairClick(object? sender, RoutedEventArgs e)
        {
            try
            {
                if (TopLevel.GetTopLevel(this) is not Window owner || this.MainWindow is not { } main ||
                    this.SelectedDevice is not { } device || device.BoardKey.Length == 0)
                {
                    return;
                }

                // Main's one way of opening the form - a device added in it is shown however it closes.
                if (await main.ShowCreateWorkbookAsync(owner, device.BoardKey, device.Id) is { } created)
                {
                    this.OpenRepair(created);
                }
            }
            catch (Exception ex)
            {
                Logger.Critical($"Starting a repair on a device failed - [{ex}]");
            }
        }

        private void OnAddChangeClick(object? sender, RoutedEventArgs e) => this.AddDeviceChange();

        private void OnChangeTextKeyDown(object? sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                this.AddDeviceChange();
                e.Handled = true;
            }
        }

        // What the change line holds, as a change of the device - the line emptied once it is in.
        internal void AddDeviceChange()
        {
            if (this.thisSelectedDeviceId is not { } id || string.IsNullOrWhiteSpace(this.ChangeTextBox.Text))
            {
                return;
            }

            DateTime date = this.ChangeDatePicker.SelectedDate ?? DateTime.Now.Date;

            if (OwnedDeviceStore.AddChange(id, date, this.ChangeTextBox.Text) is not null)
            {
                this.ChangeTextBox.Text = string.Empty;
                this.ShowDeviceAgain();
            }
        }

        private void RemoveDeviceChange(int changeId)
        {
            if (this.thisSelectedDeviceId is { } id && OwnedDeviceStore.RemoveChange(id, changeId) is not null)
            {
                this.ShowDeviceAgain();
            }
        }

        // ###########################################################################################
        // The devices read again, then the page drawn - after a change made on the page itself. The
        // page is drawn from the list, so drawing it alone showed the device as it was BEFORE the
        // change (a change added did not appear).
        // ###########################################################################################
        private void ShowDeviceAgain()
        {
            // A click, not a refresh pass: the workbooks read as they now are.
            this.thisAllWorkbooksThisPass = null;
            this.thisEntriesReadThisPass.Clear();

            this.RefreshDeviceList();
            this.ApplyDeviceView();
        }

        // The notes saved as the box is left - only when they changed.
        private void OnDeviceNotesLostFocus(object? sender, RoutedEventArgs e) => this.SaveDeviceNotes();

        // The box's notes, to the device they were typed for (thisNotesDeviceId) - only when changed.
        internal void SaveDeviceNotes()
        {
            if (this.thisNotesDeviceId is { } id && OwnedDeviceStore.Get(id) is { } device &&
                !string.Equals(device.Notes, this.DeviceNotesTextBox.Text?.Trim() ?? string.Empty, StringComparison.Ordinal))
            {
                OwnedDeviceStore.SetNotes(device.Id, this.DeviceNotesTextBox.Text);
            }
        }

        private async void OnAddDevicePhotosClick(object? sender, RoutedEventArgs e)
        {
            try
            {
                IReadOnlyList<string> picked = await this.PickDevicePhotosAsync();

                if (picked.Count > 0)
                {
                    await this.AddDevicePhotosAsync(picked);
                }
            }
            catch (Exception ex)
            {
                Logger.Critical($"Adding photos to a device failed - [{ex}]");
            }
        }

        // Copied in off the UI thread under CRT's please-wait - a dozen camera photos from a USB stick
        // froze the window (code review, 2026-10-28).
        internal async Task AddDevicePhotosAsync(IReadOnlyList<string> paths)
        {
            if (this.thisSelectedDeviceId is not { } id)
            {
                return;
            }

            IReadOnlyList<WorklogAttachmentRecord> added = await BusyOverlay.RunLocalAsync(
                this,
                CrtWaitWording.AddingDevicePhotos,
                () => Task.Run(() => OwnedDeviceStore.AddPhotos(id, paths)));

            if (added.Count > 0)
            {
                this.ShowDeviceAgain();
            }

            // Said either way - what could not be added is no longer left without a word.
            this.ShowDevicePhotosNote(OwnedDeviceWording.PhotosAdded(added.Count, paths.Count - added.Count));
        }

        private void ShowDevicePhotosNote(string? text)
        {
            this.DevicePhotosNoteText.Text = text ?? string.Empty;
            this.DevicePhotosNoteText.IsVisible = !string.IsNullOrEmpty(text);
        }

        private async Task<IReadOnlyList<string>> PickDevicePhotosAsync()
        {
            if (this.PickDevicePhotosOverrideForTests is { } pick)
            {
                return await pick();
            }

            if (TopLevel.GetTopLevel(this)?.StorageProvider is not { } storage)
            {
                return [];
            }

            IReadOnlyList<IStorageFile> files = await storage.OpenFilePickerAsync(new FilePickerOpenOptions
            {
                Title = "Pick photos of the device",
                AllowMultiple = true,
                FileTypeFilter = [ComponentPane.PictureFileType],
            });

            return files.Select(file => file.TryGetLocalPath()).OfType<string>().ToList();
        }

        private void RemoveDevicePhoto(int deviceId, int photoId)
        {
            if (OwnedDeviceStore.RemovePhoto(deviceId, photoId))
            {
                this.ShowDeviceAgain();
            }
        }

        // ###########################################################################################
        // "On my device" under a workbook's title: None, or a device with the workbook's board (or the
        // one it is on). Shown while there is a device to pick.
        // ###########################################################################################
        private void ApplyWorkbookDevice(WorkbookRecord? workbook)
        {
            if (this.WorkbookDevicePanel == null)
            {
                return;
            }

            List<OwnedDeviceRecord> choices = workbook is null ? [] : OwnedDeviceRepairs.ChoicesFor(workbook, this.thisDevices);
            this.WorkbookDevicePanel.IsVisible = choices.Count > 0;

            if (workbook is null || choices.Count == 0)
            {
                return;
            }

            var items = new List<CreateWorkbookWindow.DeviceChoice> { new(null, OwnedDeviceWording.NoDeviceChoice) };
            items.AddRange(choices.Select(device => new CreateWorkbookWindow.DeviceChoice(device.Id, device.Name)));

            this.thisApplyingWorkbookDevice = true;

            try
            {
                this.WorkbookDeviceComboBox.ItemsSource = items;
                this.WorkbookDeviceComboBox.SelectedItem =
                    items.FirstOrDefault(item => item.Id is { } id && id == workbook.OwnedDeviceId) ?? items[0];
            }
            finally
            {
                this.thisApplyingWorkbookDevice = false;
            }
        }

        private void OnWorkbookDeviceSelectionChanged(object? sender, SelectionChangedEventArgs e)
        {
            if (this.thisApplyingWorkbookDevice || this.WorkbookDeviceComboBox.SelectedItem is not CreateWorkbookWindow.DeviceChoice choice)
            {
                return;
            }

            // After the drop-down has finished setting its selection - the refresh replaces its items.
            Dispatcher.UIThread.Post(() => this.PutShownWorkbookOnDevice(choice.Id));
        }

        // The workbook shown, put on a device (or on none) - then everything drawn again.
        internal void PutShownWorkbookOnDevice(int? deviceId)
        {
            if (this.thisHeaderWorkbook is not { } workbook || workbook.OwnedDeviceId == deviceId)
            {
                return;
            }

            if (WorklogManager.SetWorkbookDevice(workbook.Id, deviceId) is not null)
            {
                this.RefreshAfterDeviceChange();
            }
        }

        // The line a workbook card gets for the device its repair is on - null for none.
        private TextBlock? BuildWorkbookDeviceLine(WorkbookRecord workbook)
        {
            if (OwnedDeviceRepairs.DeviceOf(workbook, this.thisDevices) is not { } device)
            {
                return null;
            }

            var line = new TextBlock { Text = "On " + device.Name, FontSize = 10, TextWrapping = TextWrapping.Wrap };
            line.Classes.Add("WorkbooksMuted");
            return line;
        }
    }
}
