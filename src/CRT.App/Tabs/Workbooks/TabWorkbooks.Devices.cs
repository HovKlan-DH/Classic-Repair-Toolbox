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
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace CRT
{
    // ###########################################################################################
    // "MY DEVICES" ON THE WORKBOOKS TAB (CRT 4.0 step 4 - the plan: "Your own machines, each with
    // its repairs"; called a Device wherever a person reads it). The computers the user owns are
    // listed - in Main's left panel while the tab is shown (ShowDevicesIn), as the Configuration tab's
    // scope has them (MyDevicesScope), each card with its status and how many workbooks it has, the
    // count last (owner request, 2026-11-06: "Just as in the 'Device' drop-down list").
    //
    // A CARD CLICKED PICKS THE DEVICE - the worklog bar's "Device" drop-down with it, one choice
    // (UserSettings.WorklogBarDeviceId; owner request, 2026-11-06: "it will only show the workbooks for
    // this device in the 2 workbooks lists") - so the workbook list holds only its workbooks, and CRT
    // works only in them (Main.PickDevice); "Clear" at the list's foot lets go of it. A workbook
    // clicked picks the device it is on ("I want that device selected in both the drop-down and in the
    // 'My devices' panel"), or all devices for one on none. The card marked is the device picked.
    //
    // The device picked has its PAGE in place of the workbook once its card is clicked - ONLY READ (owner
    // request, 2026-11-07: "it should not allow for direct edits or 'Add X' ... the user should go to the
    // 'Edit device' button"), each part shown only while it holds something:
    //
    //   - what it is - name, status, condition in stars, hardware and board, region, serial number,
    //     and its description in the lines it was typed in - with "Edit device" and "Delete device"
    //     (its repairs stay, on no device);
    //   - "WORKBOOKS ASSOCIATED TO THIS DEVICE" - the workbooks naming it, on any board - with their
    //     worklogs' time and cost added up (OwnedDeviceRepairs), each opening its workbook; with none,
    //     a line saying so ("New repair" went on 2026-11-07 - "Add workbook" above the tabs starts one);
    //   - its LINKS OF INTEREST and COMMENTS, and its PHOTOS, a photo clicked opened to zoom in. Its
    //     notes went the same day - its description says it (OwnedDeviceStore.FoldNotes).
    //
    // There is no list of changes made to it (owner request, 2026-11-06: "this is exactly what the
    // workbooks will cover"). The worklog bar's "Add device" adds one on any hardware and board - no
    // workbook needed - and picks it, its page open.
    //
    // A workbook card clicked, or a repair opened, brings the workbook back. Under a workbook's
    // title "Associate workbook to device" ("On my device" until 2026-11-06) puts it on a device with
    // its board, or takes it off - existing workbooks
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
        // WHAT A REFRESH NEED NOT DO AGAIN while a device's page is open (code review, 2026-10-28 -
        // each search keystroke and worklog save re-read every workbook and decoded every photo):
        //   - every workbook, read ONCE per refresh pass (null: not read yet this pass);
        //   - each photo decoded once, kept by its path, length and write time - a photo replaced is a
        //     new one (DevicePhotoThumbnails, the device form's too);
        //   - the board pane behind the page, rebuilt only once the page goes.
        // ###########################################################################################
        private List<WorkbookRecord>? thisAllWorkbooksThisPass;

        private bool thisBoardPreviewsWaitForPage;

        private List<WorkbookRecord> AllWorkbooksThisPass => this.thisAllWorkbooksThisPass ??= WorklogManager.GetAllWorkbooks();

        // The board pane is rebuilt by a refresh only while it is on screen - else when the page goes.
        private bool ShouldRebuildBoardPreviews()
        {
            this.thisBoardPreviewsWaitForPage = this.thisSelectedDeviceId is not null;
            return !this.thisBoardPreviewsWaitForPage;
        }

        // Stand-ins for the windows the device screens open - null in the app.
        internal Func<OwnedDeviceRecord, int, Task<bool>>? ConfirmDeleteDeviceOverrideForTests { get; set; }

        internal int? SelectedDeviceIdForTests => this.thisSelectedDeviceId;

        internal bool IsDevicePageShown => this.DevicePageScroller?.IsVisible == true;

        // ###########################################################################################
        // The device picked - the worklog bar's "Device" drop-down's, saved and checked by Main before the
        // tab is drawn (OwnedDeviceRepairs.BarFilter) - while it is a device the last refresh read; null
        // for all devices.
        // ###########################################################################################
        internal int? PickedDeviceId =>
            UserSettings.WorklogBarDeviceId is { } id && this.thisDevices.Any(device => device.Id == id) ? id : null;

        // ###########################################################################################
        // "My devices" drawn - a card for each device the Configuration tab's scope lists (owner request,
        // 2026-11-06: "Show all devices", "... for selected hardware", "... for selected board";
        // MyDevicesScope), the one whose page is shown marked - from the devices the caller read
        // (Main.RefreshWorklogBar's), else read again. Every device is kept for the page and the
        // lookups; only the cards follow the scope. A device that has gone (deleted elsewhere) takes
        // its page with it. With none listed the list says why: none at all, or none in the scope.
        // ###########################################################################################
        private void RefreshDeviceList(List<OwnedDeviceRecord>? devices = null)
        {
            if (this.DeviceListPanel == null)
            {
                return;
            }

            this.thisDevices = devices ?? OwnedDeviceStore.GetAll();

            // A device's page is the device picked's: a device gone, or another picked - in the bar, by a
            // workbook clicked, by "Clear" - takes the page with it.
            if (this.thisSelectedDeviceId is { } shown && this.PickedDeviceId != shown)
            {
                this.thisSelectedDeviceId = null;
            }

            string boardShown = this.BoardKeyOverrideForTests ?? this.MainWindow?.GetCurrentBoardKey() ?? string.Empty;

            // In the order "Sorting of devices" asks for (owner request, 2026-11-07; OwnedDeviceSorting).
            List<OwnedDeviceRecord> inScope = SortDevices(this.thisDevices
                .Where(device => MyDevicesScope.Holds(Scope, device.BoardKey, boardShown)));

            // "Find a previous repair" searches the devices too (owner request, 2026-11-06: "searching a serial/unique in
            // the filter in 'My Devices' should filter and highlight info in this list also") - WorklogSearchIndex.ForDevice.
            List<OwnedDeviceRecord> listed = this.thisSearchQuery.IsEmpty
                ? inScope
                : inScope.Where(device => this.thisSearchQuery.Matches(WorklogSearchIndex.ForDevice(device, BoardTextOf(device.BoardKey)))).ToList();

            int? picked = this.PickedDeviceId;
            bool byHand = UserSettings.DeviceSorting == OwnedDeviceSorting.Manual;

            // Changed in place: the drag holds the very collection (ListRowDrag), and a drag in progress is let go of
            // first, so no card is left drawn as a gap.
            this.thisDeviceDrag?.Reset();
            this.thisDeviceCards.Clear();

            foreach (OwnedDeviceRecord device in listed)
            {
                this.thisDeviceCards.Add(this.BuildDeviceCard(
                    device,
                    OwnedDeviceRepairs.Of(device.Id, this.AllWorkbooksThisPass).Count,
                    device.Id == picked,
                    byHand));
            }

            this.DeviceListHeadingText.Text = OwnedDeviceWording.ListHeading;

            // How many are listed, in a badge after the headline (owner request, 2026-11-07).
            this.DeviceCountHost.Child = WorklogInfoPillBuilder.BuildCountPill(listed.Count, fontSize: 10);
            this.NoDevicesText.Text = this.thisDevices.Count == 0
                ? OwnedDeviceWording.NoDevicesYet
                : inScope.Count > 0 ? OwnedDeviceWording.NoDevicesMatch : OwnedDeviceWording.NoDevicesInScope(Scope);
            this.NoDevicesText.IsVisible = listed.Count == 0;

            // "Clear" stays in sight, greyed out with nothing to let go of.
            this.ClearDeviceButton.Content = OwnedDeviceWording.ClearChoice;
            this.ClearDeviceButton.IsEnabled = picked is not null || this.thisSelectedDeviceId is not null;
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

        // ###########################################################################################
        // A device card clicked: the device PICKED - the bar's "Device" drop-down with it, its workbooks only,
        // its repair worked on (Main.PickDevice) - and its page in place of the workbook. A tab built on its
        // own, as a test builds it, picks it itself.
        // ###########################################################################################
        internal void SelectDevice(int deviceId)
        {
            // The device picked already: its page again, the workbook worked in left as it is - a workbook of it
            // clicked since is not swapped for the repair a pick works on.
            if (this.PickedDeviceId == deviceId)
            {
                this.ShowDevicePage(deviceId);
                return;
            }

            if (this.MainWindow is { } main && this.thisDevices.FirstOrDefault(device => device.Id == deviceId) is { } picked)
            {
                main.PickDevice(picked, showPage: true);
                return;
            }

            UserSettings.WorklogBarDeviceId = deviceId;
            this.ShowDevicePage(deviceId);
            this.RefreshWorkbooks();
        }

        // The device's page in place of the workbook - the device picked's (Main.PickDevice).
        internal void ShowDevicePage(int deviceId)
        {
            // What "Add device" said of its photos is of that device alone.
            if (this.thisPhotosNote is { } said && said.DeviceId != deviceId)
            {
                this.thisPhotosNote = null;
            }

            this.thisSelectedDeviceId = deviceId;
            this.ShowDeviceAgain();
        }

        // ###########################################################################################
        // "Clear" under "My devices" (owner request, 2026-11-06: "clicking that will unselect the device and by
        // then show all again"): no device picked, the workbook back in place of a device's page.
        // ###########################################################################################
        private void OnClearDeviceClick(object? sender, RoutedEventArgs e) => this.ClearDevicePick();

        internal void ClearDevicePick()
        {
            this.thisSelectedDeviceId = null;

            if (this.MainWindow is { } main)
            {
                main.ClearDevicePick();
                return;
            }

            UserSettings.WorklogBarDeviceId = null;
            this.RefreshWorkbooks();
        }

        // ###########################################################################################
        // One card in "My devices" (owner requests, 2026-11-06 and 2026-11-07): its lines on the left - the name, its
        // serial number or unique identifier when it has one, its hardware and board - and its badges on the right, a
        // line each from the top - its status, how many workbooks it has (an outlined badge, as the bar's "Device"
        // drop-down counts them), and its condition in stars when it has any. With no serial the board moves up a
        // line ("it should not show a blank line - then instead move up the hardware/board"), and the stars, the third
        // badge, keep a line of their own. What the search found is marked on every line.
        //
        // Sorted by hand, a handle at the card's right edge drags it (ListRowDrag) - a press on the handle is the
        // drag's, never the card's pick, which draws the list again.
        // ###########################################################################################
        private DeviceListCard BuildDeviceCard(OwnedDeviceRecord device, int workbookCount, bool isSelected, bool byHand)
        {
            var lines = new List<Control>
            {
                this.BuildHighlightedTextBlock(device.Name, 11, TextWrapping.Wrap, block =>
                {
                    block.FontWeight = FontWeight.Bold;
                    block.VerticalAlignment = VerticalAlignment.Center;
                }),
            };

            string serial = device.SerialNumber.Trim();

            if (serial.Length > 0)
            {
                lines.Add(this.BuildHighlightedTextBlock(serial, 10, TextWrapping.Wrap, block =>
                {
                    block.VerticalAlignment = VerticalAlignment.Center;
                    block.Classes.Add("WorkbooksMuted");
                }));
            }

            lines.Add(this.BuildHighlightedTextBlock(BoardTextOf(device.BoardKey), 10, TextWrapping.Wrap, block =>
                block.Classes.Add("WorkbooksMuted")));

            var badges = new List<Control>
            {
                WorklogInfoPillBuilder.BuildDeviceStatusPill(device.Status, fontSize: 10),
                WorklogInfoPillBuilder.BuildCountPill(workbookCount, fontSize: 10),
            };

            if (OwnedDeviceCondition.Within(device.Condition) > 0)
            {
                StackPanel stars = ConditionStars.Build(device.Condition, size: 11);
                ToolTip.SetTip(stars, OwnedDeviceWording.ConditionLine(device.Condition));
                badges.Add(stars);
            }

            var body = new Grid { ColumnDefinitions = new ColumnDefinitions("*,Auto"), RowSpacing = 3 };

            for (int row = 0; row < Math.Max(lines.Count, badges.Count); row++)
            {
                body.RowDefinitions.Add(new RowDefinition(GridLength.Auto));

                if (row < lines.Count)
                {
                    Grid.SetRow(lines[row], row);
                    body.Children.Add(lines[row]);
                }

                if (row < badges.Count)
                {
                    Control badge = badges[row];
                    badge.HorizontalAlignment = HorizontalAlignment.Right;
                    badge.VerticalAlignment = VerticalAlignment.Center;
                    badge.Margin = new Thickness(4, 0, 0, 0);
                    Grid.SetRow(badge, row);
                    Grid.SetColumn(badge, 1);
                    body.Children.Add(badge);
                }
            }

            Control content = byHand ? this.WithDragHandle(body) : body;
            var card = new DeviceListCard(device, content);

            if (isSelected)
            {
                card.Classes.Add("Selected");
            }

            card.PointerPressed += (_, _) => this.SelectDevice(device.Id);
            return card;
        }

        // The card's lines with the drag handle at their right - Font Awesome's up-down, as every row CRT moves - under
        // the north/south cursor (owner request, 2026-11-07: "the icon for up/down should show the north/south icon when
        // hovering"), the shared one the board pane's previews take while they can be moved.
        private Grid WithDragHandle(Control body)
        {
            var handle = new Border
            {
                Width = 18,
                Margin = new Thickness(4, 0, -3, 0),
                Background = Brushes.Transparent,
                Cursor = TabWorkbooks.PreviewDragCursor,
                VerticalAlignment = VerticalAlignment.Stretch,
                Child = new TextBlock
                {
                    Text = "",
                    FontFamily = ThemeResources.ResolveFontAwesomeSolid(),
                    FontSize = 12,
                    Opacity = 0.6,
                    HorizontalAlignment = HorizontalAlignment.Center,
                    VerticalAlignment = VerticalAlignment.Center,
                },
            };

            handle.Classes.Add("DeviceDragHandle");
            ToolTip.SetTip(handle, OwnedDeviceWording.DragDeviceTip);

            handle.PointerPressed += (sender, e) =>
            {
                this.thisDeviceDrag?.Press(sender, e);
                e.Handled = true;
            };

            var grid = new Grid { ColumnDefinitions = new ColumnDefinitions("*,Auto") };
            Grid.SetColumn(handle, 1);
            grid.Children.Add(body);
            grid.Children.Add(handle);
            return grid;
        }

        // ###########################################################################################
        // The list's cards, the very collection its items are and its drag moves (ListRowDrag) - set up once, from the
        // constructor. The drag's host is the list's own area, an ancestor of the list wherever it stands - in Main's
        // left panel or, in a tab built on its own, above the workbooks.
        // ###########################################################################################
        private readonly ObservableCollection<DeviceListCard> thisDeviceCards = new();

        private ListRowDrag<DeviceListCard>? thisDeviceDrag;

        private void WireDeviceList()
        {
            if (this.DeviceListPanel == null || this.DeviceListArea == null)
            {
                return;
            }

            this.DeviceListPanel.ItemsSource = this.thisDeviceCards;
            this.thisDeviceDrag = new ListRowDrag<DeviceListCard>(
                this.DeviceListArea,
                this.DeviceListPanel,
                this.DeviceListScroller,
                this.thisDeviceCards,
                (_, _, _) => this.SaveDeviceOrderAsListed());
        }

        // The cards "My devices" shows - for the tests, which read them with the tab never shown.
        internal IReadOnlyList<DeviceListCard> DeviceCardsForTests => this.thisDeviceCards;

        internal ListRowDrag<DeviceListCard>? DeviceDragForTests => this.thisDeviceDrag;

        // ###########################################################################################
        // A card dropped: the order as the list now shows it, kept - the devices not listed (another board's under
        // the scope, or one the search left out) keeping their places (OwnedDeviceSorting.Reordered) - then every list
        // drawn again in it, the bar's "Device" drop-down too.
        // ###########################################################################################
        private void SaveDeviceOrderAsListed()
        {
            List<int> all = OwnedDeviceSorting.Sort(this.thisDevices, OwnedDeviceSorting.Manual, Main.FormatBoardKeyAsFullNames)
                .Select(device => device.Id)
                .ToList();

            List<int> order = OwnedDeviceSorting.Reordered(all, this.thisDeviceCards.Select(card => card.Device.Id).ToList());

            if (!OwnedDeviceStore.SetOrder(order))
            {
                Logger.Warning("The order of the devices could not be written in full - see above");
            }

            this.RefreshAfterDeviceChange();
        }

        // The devices in the order "Sorting of devices" asks for - the bar's drop-down's too (Main.FillBarPickers).
        internal static List<OwnedDeviceRecord> SortDevices(IEnumerable<OwnedDeviceRecord> devices) =>
            OwnedDeviceSorting.Sort(devices, UserSettings.DeviceSorting, Main.FormatBoardKeyAsFullNames);

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

            StackPanel stars = ConditionStars.Build(device.Condition, size: 17);
            ToolTip.SetTip(stars, OwnedDeviceWording.ConditionLine(device.Condition));
            this.DeviceConditionHost.Child = stars;

            this.DeviceDetailText.Text = OwnedDeviceWording.DetailLine(device, BoardTextOf(device.BoardKey));

            string description = OwnedDeviceWording.DescriptionText(device.Description);
            TextLinkRenderer.SetLinkText(this.DeviceDescriptionText, description);
            this.DeviceDescriptionText.IsVisible = description.Length > 0;

            this.FillDeviceRepairs(device);
            this.FillDeviceLinks(device);
            this.FillDeviceComments(device);
            this.FillDevicePhotos(device);
        }

        // ###########################################################################################
        // Its links of interest, as the worklog editor lists them - a row each, its headline, the row clicked opening
        // the address (ExternalTargetLauncher, the one way out of CRT). Added and changed in "Edit device".
        // ###########################################################################################
        private void FillDeviceLinks(OwnedDeviceRecord device)
        {
            List<WorklogLinkRecord> links = device.Links.Where(link => !string.IsNullOrWhiteSpace(link.Url)).ToList();

            // A headline only over something (owner request, 2026-11-07).
            this.DeviceLinksSection.IsVisible = links.Count > 0;
            this.DeviceLinksHeadingText.Text = OwnedDeviceWording.LinksHeading(links.Count);

            this.DeviceLinksPanel.Children.Clear();

            foreach (WorklogLinkRecord link in links)
            {
                string headline = string.IsNullOrWhiteSpace(link.Headline) ? link.Url.Trim() : link.Headline.Trim();
                var row = new Border
                {
                    Tag = link,
                    Background = ThemeResources.ResolveBrush("Form_Bg", Brushes.Transparent),
                    BorderBrush = ThemeResources.ResolveBrush("Form_Border", Brushes.Gray),
                    BorderThickness = new Thickness(1),
                    CornerRadius = new CornerRadius(3),
                    Padding = new Thickness(8, 6),
                    Cursor = new Cursor(StandardCursorType.Hand),
                    Child = new TextBlock
                    {
                        Text = headline,
                        FontSize = 12,
                        FontWeight = FontWeight.SemiBold,
                        TextTrimming = TextTrimming.CharacterEllipsis,
                        Foreground = ThemeResources.ResolveBrush("Link_Fg", Brushes.SteelBlue),
                        TextDecorations = TextDecorations.Underline,
                    },
                };

                ToolTip.SetTip(row, link.Url.Trim());
                string url = link.Url.Trim();
                row.PointerPressed += (_, e) =>
                {
                    if (e.GetCurrentPoint(row).Properties.IsLeftButtonPressed)
                    {
                        this.OpenDeviceLink(url);
                    }
                };

                this.DeviceLinksPanel.Children.Add(row);
            }
        }

        // A link clicked - handed the address that would be opened, a test stands in for opening it.
        internal Action<string>? OpenLinkOverrideForTests { get; set; }

        // As typed, opened by LinkTarget's rule - "www.c64-wiki.com" too (code review, 2026-11-08).
        private void OpenDeviceLink(string url) => ExternalTargetLauncher.OpenLinkOrSay(url, this.OpenLinkOverrideForTests);

        // ###########################################################################################
        // Its comments, newest first as the worklog editor lists them, each dated - read only here; "Edit
        // device" adds, changes and deletes them (owner request, 2026-11-06).
        // ###########################################################################################
        private void FillDeviceComments(OwnedDeviceRecord device)
        {
            List<WorklogCommentRecord> comments = device.Comments
                .Where(comment => !string.IsNullOrWhiteSpace(comment.Text))
                .OrderByDescending(comment => comment.Date)
                .ThenByDescending(comment => comment.Id)
                .ToList();

            this.DeviceCommentsSection.IsVisible = comments.Count > 0;
            this.DeviceCommentsHeadingText.Text = OwnedDeviceWording.CommentsHeading(comments.Count);

            this.DeviceCommentsPanel.Children.Clear();

            foreach (WorklogCommentRecord comment in comments)
            {
                var date = new TextBlock { Text = OwnedDeviceWording.CommentDate(comment.Date), FontSize = 10 };
                date.Classes.Add("WorkbooksMuted");

                var text = new TextBlock { FontSize = 11, TextWrapping = TextWrapping.Wrap };
                // Markup: trimmed so its first line keeps the indent that says what it is (code review, 2026-11-08).
                TextLinkRenderer.SetLinkText(text, RichNote.TrimMarkup(comment.Text));

                var body = new StackPanel { Spacing = 1 };
                body.Children.Add(date);
                body.Children.Add(text);

                this.DeviceCommentsPanel.Children.Add(new Border
                {
                    Tag = comment,
                    Background = ThemeResources.ResolveBrush("Form_Bg", Brushes.Transparent),
                    BorderBrush = ThemeResources.ResolveBrush("Form_Border", Brushes.Gray),
                    BorderThickness = new Thickness(1),
                    CornerRadius = new CornerRadius(3),
                    Padding = new Thickness(8, 6),
                    Child = body,
                });
            }
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

        // ###########################################################################################
        // Its photos, in their order, EACH IN A PANEL OF ITS OWN WITH ITS COMMENT UNDER IT (owner request, 2026-11-06: "I
        // want each photo to have its own panel and with its text displayed below") - a photo clicked opened whole, to
        // zoom in with the wheel. Added, removed and ordered in "Edit device" only (owner request, 2026-11-07: the page
        // is only read); the section shown only with photos - or with photos "Add device" could not copy in to say.
        // ###########################################################################################
        private void FillDevicePhotos(OwnedDeviceRecord device)
        {
            this.DevicePhotosHeadingText.Text = OwnedDeviceWording.PhotosHeading(device.Photos.Count);

            string? note = this.thisPhotosNote is { } said && said.DeviceId == device.Id ? said.Text : null;
            this.DevicePhotosNoteText.Text = note ?? string.Empty;
            this.DevicePhotosNoteText.IsVisible = note is not null;
            this.DevicePhotosSection.IsVisible = device.Photos.Count > 0 || note is not null;

            this.DevicePhotosPanel.Children.Clear();

            foreach (WorklogAttachmentRecord photo in device.Photos.OrderBy(photo => photo.DisplayOrder).ThenBy(photo => photo.Id))
            {
                var image = new Image { Width = 160, Height = 120, Stretch = Stretch.Uniform, Cursor = new Cursor(StandardCursorType.Hand) };
                ToolTip.SetTip(image, "Click to view");
                string? path = OwnedDeviceStore.PhotoPath(device.Id, photo.FileName);
                ShowDevicePhoto(image, path);
                image.PointerPressed += (_, e) =>
                {
                    if (e.GetCurrentPoint(image).Properties.IsLeftButtonPressed)
                    {
                        e.Handled = true;
                        this.ViewDevicePhoto(photo, path);
                    }
                };

                var comment = new TextBlock
                {
                    FontSize = 11,
                    TextWrapping = TextWrapping.Wrap,
                    Width = 160,
                    IsVisible = !string.IsNullOrWhiteSpace(photo.Comment),
                };
                TextLinkRenderer.SetLinkText(comment, RichNote.TrimMarkup(photo.Comment));

                var body = new StackPanel { Spacing = 5 };
                body.Children.Add(image);
                body.Children.Add(comment);

                this.DevicePhotosPanel.Children.Add(new Border
                {
                    Tag = photo,
                    Background = ThemeResources.ResolveBrush("Form_Bg", Brushes.Transparent),
                    BorderBrush = ThemeResources.ResolveBrush("Form_Border", Brushes.Gray),
                    BorderThickness = new Thickness(1),
                    CornerRadius = new CornerRadius(3),
                    Padding = new Thickness(6),
                    Child = body,
                });
            }
        }

        // ###########################################################################################
        // A photo decoded small OFF the UI thread and shown once read - ONCE per photo, the read the device
        // form shares (DevicePhotoThumbnails). A photo that cannot be read leaves its frame empty.
        // ###########################################################################################
        private static void ShowDevicePhoto(Image image, string? path)
        {
            Task<Bitmap?> reading = DevicePhotoThumbnails.Read(path);

            if (reading.IsCompletedSuccessfully)
            {
                image.Source = reading.Result;
                return;
            }

            _ = reading.ContinueWith(
                done => Dispatcher.UIThread.Post(() => image.Source = done.Result),
                TaskContinuationOptions.OnlyOnRanToCompletion);
        }

        // Stands in for showing the photo viewer in a test, which cannot show a dialog. Null in the app.
        internal Action<WorklogPhotoViewerWindow>? ShowPhotoViewerOverrideForTests { get; set; }

        // ###########################################################################################
        // A photo of the device opened whole - the worklog's own viewer, which zooms with the wheel (owner
        // request, 2026-11-06: "when I click a photo in here and in the worklog, that I then can zoom
        // in/out with scrollwheel").
        // ###########################################################################################
        private void ViewDevicePhoto(WorklogAttachmentRecord photo, string? path)
        {
            var viewer = new WorklogPhotoViewerWindow();
            viewer.Initialize(photo.FileName, photo.Comment ?? string.Empty, path);

            if (this.ShowPhotoViewerOverrideForTests is { } show)
            {
                show(viewer);
            }
            else if (this.DeviceFormOwner is { } owner)
            {
                _ = viewer.ShowDialog(owner);
            }
        }

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

        // Whether a board comes in PAL and NTSC, which the device form asks of the board picked - null:
        // the board as CRT views it (OwnedDeviceWindow.BoardHasRegionsAsViewedAsync). A test's stand-in.
        internal Func<string, Task<bool>>? BoardHasRegionsOverrideForTests { get; set; }

        // Stands in for showing the device form in a test, which cannot show a dialog: handed the form as
        // the tab filled it in, it closes it as a user would - with the record saved, or null. Null in
        // the app.
        internal Func<OwnedDeviceWindow, Task<OwnedDeviceRecord?>>? ShowDeviceFormOverrideForTests { get; set; }

        // ###########################################################################################
        // "Add device" - the worklog bar's (Main.AddDeviceFromBarAsync; the one above "My devices" went on
        // 2026-11-06, owner request: "no need for it"): the form on ANY hardware and board, the board shown
        // preselected - no workbook needed, as no workbook is made (owner request, 2026-11-06: "it is
        // not logical that I need to create a workbook to create a new device") - and the device
        // added picked, as a card clicked is, its page open. True when a device was added.
        // ###########################################################################################
        internal async Task<bool> AddDeviceAsync()
        {
            var dialog = new OwnedDeviceWindow();
            dialog.Initialize(
                this.BoardsForDevices,
                this.BoardKeyOverrideForTests ?? this.MainWindow?.GetCurrentBoardKey(),
                boardHasRegions: this.BoardHasRegionsOverrideForTests);

            if (await this.ShowDeviceFormAsync(dialog) is not { } added)
            {
                return false;
            }

            // A photo picked in the form that could not be copied in as it was added is said on its page - noted
            // BEFORE the page is drawn, which draws it (code review, 2026-11-08).
            if (dialog.PhotosNotAdded > 0)
            {
                this.ShowDevicePhotosNote(added.Id, OwnedDeviceWording.PhotosAdded(dialog.PhotosAdded, dialog.PhotosNotAdded));
            }

            // Picked, as a card clicked is - its page open, on its own board.
            if (this.MainWindow is { } main)
            {
                main.PickDevice(added, showPage: true);
            }
            else
            {
                UserSettings.WorklogBarDeviceId = added.Id;
                this.ShowDevicePage(added.Id);
                this.RefreshWorkbooks();
            }

            return true;
        }

        private async void OnEditDeviceClick(object? sender, RoutedEventArgs e)
        {
            try
            {
                await this.EditSelectedDeviceAsync();
            }
            catch (Exception ex)
            {
                Logger.Critical($"Editing a device failed - [{ex}]");
            }
        }

        // "Edit device": the form on the device whose page is shown, then the page as it was saved.
        internal async Task EditSelectedDeviceAsync()
        {
            if (this.SelectedDevice is not { } device)
            {
                return;
            }

            // With workbooks, its hardware and board are locked (owner decision, 2026-11-06: "Locked while it has workbooks").
            var dialog = new OwnedDeviceWindow();
            dialog.InitializeForEdit(
                device,
                this.BoardsForDevices,
                this.BoardHasRegionsOverrideForTests,
                OwnedDeviceRepairs.Of(device.Id, WorklogManager.GetAllWorkbooks()).Count);

            // Its comments and photos are saved as they change (owner decision, 2026-11-06: "At once, as a
            // worklog does"), so the form called off may have changed it too.
            if (await this.ShowDeviceFormAsync(dialog) is not null || dialog.ChangedOnDisk)
            {
                this.RefreshAfterDeviceChange();
            }
        }

        // The device form shown over CRT, closing with the record saved - null when cancelled, or with
        // no window to show it over.
        private async Task<OwnedDeviceRecord?> ShowDeviceFormAsync(OwnedDeviceWindow dialog)
        {
            if (this.ShowDeviceFormOverrideForTests is { } show)
            {
                return await show(dialog);
            }

            return this.DeviceFormOwner is { } owner
                ? await dialog.ShowDialog<OwnedDeviceRecord?>(owner)
                : null;
        }

        // ###########################################################################################
        // The window the device form opens over: the one this tab is in - or, while another tab is shown,
        // CRT's own. The worklog bar's "Add device" asks from any tab, and a tab not shown is out of the
        // window's visual tree, so it had no window to show the form over and nothing happened (owner
        // report, 2026-11-06: "clicking the 'Add device' in the top-bar result in nothing").
        // ###########################################################################################
        private Window? DeviceFormOwner => TopLevel.GetTopLevel(this) as Window ?? this.MainWindow;

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
        // The devices read again, then the page drawn - after a change made on the page itself. The
        // page is drawn from the list, so drawing it alone showed the device as it was BEFORE the
        // change (what was just added did not appear).
        // ###########################################################################################
        private void ShowDeviceAgain()
        {
            // A click, not a refresh pass: the workbooks read as they now are.
            this.thisAllWorkbooksThisPass = null;
            this.thisEntriesReadThisPass.Clear();

            this.RefreshDeviceList();
            this.ApplyDeviceView();
        }

        // ###########################################################################################
        // What "Add device" said of the photos it could not copy in - on THAT device's page, drawn with it (FillDevicePhotos),
        // until another device's page is shown (code review, 2026-11-08: set after the page was drawn, inside the photos'
        // section the page had hidden for a device with no photos, it was never seen).
        // ###########################################################################################
        private (int DeviceId, string Text)? thisPhotosNote;

        private void ShowDevicePhotosNote(int deviceId, string? text)
        {
            this.thisPhotosNote = string.IsNullOrEmpty(text) ? null : (deviceId, text);

            if (this.thisSelectedDeviceId == deviceId && this.SelectedDevice is { } device)
            {
                this.FillDevicePhotos(device);
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

        // ###########################################################################################
        // The workbook shown, put on a device (or on none) - then everything drawn again.
        //
        // *** THE DEVICE PICKED FOLLOWS IT (code review, 2026-11-08). *** CRT works only in the device picked's workbooks,
        // so the workbook moved to another device - or to none - dropped out of the list and the bar, and the device's
        // other workbook became the one worked in. It takes the pick along as a workbook activated off the device picked
        // does (OwnedDeviceRepairs.PickFollowing): its new device, or all devices for none.
        // ###########################################################################################
        internal void PutShownWorkbookOnDevice(int? deviceId)
        {
            if (this.thisHeaderWorkbook is not { } workbook || workbook.OwnedDeviceId == deviceId)
            {
                return;
            }

            if (WorklogManager.SetWorkbookDevice(workbook.Id, deviceId) is { } moved)
            {
                UserSettings.WorklogBarDeviceId = OwnedDeviceRepairs.PickFollowing(UserSettings.WorklogBarDeviceId, moved, this.thisDevices);
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
