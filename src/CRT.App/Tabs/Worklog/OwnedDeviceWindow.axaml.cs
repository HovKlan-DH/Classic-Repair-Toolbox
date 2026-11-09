using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Media;
using Avalonia.Media.Imaging;
using Avalonia.Threading;
using Avalonia.VisualTree;
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
    // "Add device" and "Edit device" (CRT 4.0 step 4) - one form for both, as "Add workbook" is one form
    // for creating and editing a workbook. Writes through OwnedDeviceStore and closes with the record
    // that reached disk (null when cancelled); a device that cannot be saved keeps the form open with
    // the reason.
    //
    // LAID OUT AS THE WORKLOG'S EDITOR IS (owner request, 2026-11-06: "I think this modal needs to be more
    // like the 'Worklog' modal, where it has these same possibilities: Name + Hardware + Board + Status +
    // Region + Serial + 'Description' as-of today - Comments - Photos"): what it is and its comments on
    // the left, its photos on the right. A photo's thumbnail clicked opens the worklog's own viewer, which
    // zooms with the wheel ("when I click a photo in here and in the worklog, that I then can zoom in/out").
    //
    // COMMENTS AND PHOTOS OF AN EXISTING DEVICE ARE SAVED AS THEY CHANGE (owner decision, 2026-11-06, asked:
    // "At once, as a worklog does") - Cancel takes back only the fields, and ChangedOnDisk tells the caller
    // the device changed all the same. A new device has no folder yet: the form holds its comments and
    // photos, and "Add device" writes them with it.
    //
    // PHOTOS ARE COPIED IN, AND THEIR THUMBNAILS READ, OFF THE UI THREAD (code review, 2026-11-08): a copy
    // under the form's own "please wait" (CopyPhotosInAsync), a thumbnail shown once it is read
    // (ShowThumbnail) - as the device's page reads its photos.
    //
    // Opened from the worklog bar's "Add device" - any hardware and board, no workbook needed (owner
    // request, 2026-11-06) - from "Add device" in "Add workbook" (the repair's board, fixed), and from
    // "Edit device" on a device's page.
    //
    // The hardware and board are picked from the boards CRT knows (the drop-down lists' own), the
    // board shown preselected for a new device. A device whose board CRT no longer lists keeps it.
    //
    // REGION FOLLOWS THE BOARD (owner request, 2026-11-06: "currently I always see the drop-down of a
    // region, even if that board does not have the region concept"): it shows only while the board
    // picked comes in PAL and NTSC - BoardRegions.Has, CRT's one rule, asked of the board as CRT views
    // it, read off the UI thread (BoardHasRegionsAsViewedAsync). Each board is asked once per form;
    // until a board answers, the form stays as it is - Region stands beside Status, so nothing below it
    // moves. Nothing is set where it cannot be seen: hidden, the region goes back to what the device has
    // stored - none for a new one - so a region an existing device has on a board without regions is
    // KEPT, never wiped by opening and saving the form. "Unknown" is none, and both it and the status
    // "Unknown" come first in their lists (owner request, 2026-11-06).
    //
    // The description is several lines: Enter there starts a new line, and Ctrl+Enter saves, as the note
    // of "Add workbook" does.
    // ###########################################################################################
    public partial class OwnedDeviceWindow : Window
    {
        // How a device with no region reads in the drop-down ("Not set" until 2026-11-06, owner request).
        internal const string NoRegion = "Unknown";

        private IReadOnlyList<HardwareBoardEntry> thisBoards = [];
        private OwnedDeviceRecord? thisEditing;

        // The board kept among the choices though CRT may no longer list it: the device's own when
        // editing, the fixed one when adding for a repair.
        private string? thisKeptBoardKey;

        // Whether a board comes in PAL and NTSC, by board key - the host's question (a test's stand-in),
        // else the board as CRT views it - and the answers given, kept for the form's life.
        private Func<string, Task<IReadOnlyList<string>>> thisBoardRegions = BoardRegionsAsViewedAsync;

        private readonly Dictionary<string, IReadOnlyList<string>> thisRegionsByBoard = new(StringComparer.OrdinalIgnoreCase);

        // Set while the code fills the drop-downs, so their own selection changes ask nothing.
        private bool thisFillingChoices;

        // Its condition, 0 to 5 stars, as the stars show it.
        private int thisCondition;

        // ###########################################################################################
        // THE COMMENTS AND PHOTOS LISTED. An existing device's are the device as last saved (thisListed); a
        // new one's are held here until "Add device" - a photo as the file picked, copied in only then.
        // ###########################################################################################
        private OwnedDeviceRecord? thisListed;

        private readonly List<WorklogCommentRecord> thisHeldComments = new();

        private readonly List<WorklogLinkRecord> thisHeldLinks = new();

        private readonly List<HeldPhoto> thisHeldPhotos = new();

        private sealed record HeldPhoto(int Id, string SourcePath, string Comment);

        // ###########################################################################################
        // The photos' rows, ONE collection changed in place - ListRowDrag moves rows in it as they are dragged, and
        // tells PhotoMoved where one landed (owner request, 2026-11-06: "I would like to be able to rearrange the
        // photos, just like it is possible to do in the worklogs"). Its drag is let go of before the rows are made again.
        // ###########################################################################################
        private readonly ObservableCollection<WorklogAttachmentRow> thisPhotoRows = new();

        private readonly ListRowDrag<WorklogAttachmentRow> thisPhotoDrag;

        // The record that reached disk - what the dialog closes with. For the tests, which show the
        // window without an owner and so cannot read a ShowDialog result.
        internal OwnedDeviceRecord? Saved { get; private set; }

        // Whether an existing device's comments or photos were saved while the form was open - so the
        // device changed even when the form is then called off.
        internal bool ChangedOnDisk { get; private set; }

        // How many of a new device's photos could not be copied in as it was added - said on its page.
        internal int PhotosNotAdded { get; private set; }

        // How many of a new device's photos were - said with those that were not.
        internal int PhotosAdded { get; private set; }

        // What the last board's regions, asked and not answered at once, are waited on by - for a test.
        internal Task RegionsFollowedForTests { get; private set; } = Task.CompletedTask;

        // Stand-ins for the small windows the form opens, which a test cannot show: a comment asked for
        // (handed the text it has, null for a new one; null back when called off), a photo and its comment
        // (handed the photo's row, null for a new one), and the photo viewer.
        internal Func<string?, Task<string?>>? AskCommentOverrideForTests { get; set; }

        internal Func<WorklogAttachmentRow?, Task<WorklogAddPhotoWindow.PhotoResult?>>? AskPhotoOverrideForTests { get; set; }

        // A link asked for (handed the headline and address it has, null for a new one; null back when called off),
        // and one opened.
        internal Func<(string Headline, string Url)?, Task<(string Headline, string Url)?>>? AskLinkOverrideForTests { get; set; }

        internal Action<string>? OpenLinkOverrideForTests { get; set; }

        internal Action<WorklogPhotoViewerWindow>? ShowPhotoViewerOverrideForTests { get; set; }

        public OwnedDeviceWindow()
        {
            this.InitializeComponent();

            this.PhotosList.ItemsSource = this.thisPhotoRows;
            this.thisPhotoDrag = new ListRowDrag<WorklogAttachmentRow>(
                this, this.PhotosList, this.PhotosScroller, this.thisPhotoRows, (_, _, _) => this.PhotoMoved());

            this.Opened += (_, _) =>
                Dispatcher.UIThread.Post(() => this.NameTextBox.Focus(), DispatcherPriority.Background);

            this.AddHandler(KeyDownEvent, this.OnWindowPreviewKeyDown, RoutingStrategies.Tunnel);

            this.BuildConditionStars();
            this.ShowCommentsAndPhotos();
        }

        // ###########################################################################################
        // A new device, on `boardKey` (the board shown) when CRT lists it. `boardFixed`: the device
        // is for a repair on that board ("Add device" in "Add workbook"), so its hardware and board are
        // shown, not to be changed - a device with another board is no device for it.
        // `boardRegions`: the regions a board comes in (BoardRegions.Of) - null asks the board as CRT views it.
        // ###########################################################################################
        public void Initialize(
            IReadOnlyList<HardwareBoardEntry> boards,
            string? boardKey,
            bool boardFixed = false,
            Func<string, Task<IReadOnlyList<string>>>? boardRegions = null)
        {
            this.thisBoards = boards ?? [];
            this.thisBoardRegions = boardRegions ?? BoardRegionsAsViewedAsync;

            // A fixed board is kept even when CRT no longer lists it (code review, 2026-10-29: the
            // locked box fell back to another board, and the device was saved there).
            this.thisKeptBoardKey = boardFixed ? boardKey : null;
            this.FillChoices(boardKey, region: string.Empty, OwnedDeviceStatus.Unknown);
            this.HardwareComboBox.IsEnabled = !boardFixed;
            this.BoardComboBox.IsEnabled = !boardFixed;
            this.UpdateSaveEnabled();
        }

        // ###########################################################################################
        // An existing device: its fields filled in, "Save device" in place of "Add device", its comments
        // and photos listed. A region it has shows at once - its board most likely comes in both - until
        // its board answers.
        // ###########################################################################################
        // `workbookCount`: how many workbooks the device has - with any, its hardware and board are locked (owner decision,
        // 2026-11-06, asked: "Locked while it has workbooks"): a workbook is associated only to a device of its own board.
        public void InitializeForEdit(
            OwnedDeviceRecord device,
            IReadOnlyList<HardwareBoardEntry> boards,
            Func<string, Task<IReadOnlyList<string>>>? boardRegions = null,
            int workbookCount = 0)
        {
            this.thisEditing = device;
            this.thisListed = device;
            this.thisKeptBoardKey = device.BoardKey;
            this.thisBoards = boards ?? [];
            this.thisBoardRegions = boardRegions ?? BoardRegionsAsViewedAsync;

            this.Title = "Edit device";
            this.HeaderText.Text = "Edit device";
            this.SaveButton.Content = "Save device";
            this.NameTextBox.Text = device.Name;
            this.SerialTextBox.Text = device.SerialNumber;
            this.DescriptionTextBox.Markup = device.Description;
            this.RegionPanel.IsVisible = device.Region.Trim().Length > 0;
            this.SetCondition(device.Condition);

            this.FillChoices(device.BoardKey, device.Region, device.Status);

            bool locked = workbookCount > 0;
            this.HardwareComboBox.IsEnabled = !locked;
            this.BoardComboBox.IsEnabled = !locked;
            this.BoardLockedText.Text = locked ? OwnedDeviceWording.BoardLocked(workbookCount) : string.Empty;
            this.BoardLockedText.IsVisible = locked;

            this.UpdateSaveEnabled();
            this.ShowCommentsAndPhotos();
        }

        // The region and status first, then the hardware and boards - whose board asks for its regions.
        private void FillChoices(string? boardKey, string region, string status)
        {
            this.ListRegions(BoardRegions.PalAndNtsc, region);

            // A status this CRT does not offer - "Not mine any more" an earlier pre-release stored, or a later
            // CRT's - is listed after them, so opening and saving the form keeps it.
            List<string> statuses = [.. OwnedDeviceStatus.All];
            string normalised = OwnedDeviceStatus.Normalise(status);
            if (!statuses.Contains(normalised))
            {
                statuses.Add(normalised);
            }

            this.StatusComboBox.ItemsSource = statuses;
            this.StatusComboBox.SelectedItem = normalised;

            (string hardware, string board) = BoardKeyParts.Split(boardKey);

            List<string> hardwareNames = this.thisBoards
                .Select(entry => entry.HardwareName)
                .Where(name => !string.IsNullOrWhiteSpace(name))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();

            // A board CRT no longer lists stays the device's.
            if (hardware.Length > 0 && !hardwareNames.Contains(hardware, StringComparer.OrdinalIgnoreCase))
            {
                hardwareNames.Add(hardware);
            }

            this.thisFillingChoices = true;

            try
            {
                this.HardwareComboBox.ItemsSource = hardwareNames;
                this.HardwareComboBox.SelectedItem = hardwareNames.FirstOrDefault(name => string.Equals(name, hardware, StringComparison.OrdinalIgnoreCase))
                    ?? hardwareNames.FirstOrDefault();
            }
            finally
            {
                this.thisFillingChoices = false;
            }

            this.FillBoards(board);
        }

        private void OnHardwareSelectionChanged(object? sender, SelectionChangedEventArgs e)
        {
            if (!this.thisFillingChoices)
            {
                this.FillBoards(this.BoardComboBox.SelectedItem as string);
            }
        }

        // A board picked: the region shown or not, as it comes.
        private void OnBoardSelectionChanged(object? sender, SelectionChangedEventArgs e)
        {
            if (!this.thisFillingChoices)
            {
                this.FollowBoardRegions();
            }
        }

        // The boards of the hardware picked, keeping `keep` when that hardware has it - then the
        // region as the board picked comes.
        private void FillBoards(string? keep)
        {
            string hardware = this.HardwareComboBox.SelectedItem as string ?? string.Empty;

            List<string> boards = this.thisBoards
                .Where(entry => string.Equals(entry.HardwareName, hardware, StringComparison.OrdinalIgnoreCase))
                .Select(entry => entry.BoardName)
                .Where(name => !string.IsNullOrWhiteSpace(name))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();

            // The device's own board - or the fixed one - when CRT no longer lists it.
            if (this.thisKeptBoardKey is { } kept && BoardKeyParts.Split(kept) is var (ownHardware, ownBoard) &&
                string.Equals(ownHardware, hardware, StringComparison.OrdinalIgnoreCase) &&
                ownBoard.Length > 0 && !boards.Contains(ownBoard, StringComparer.OrdinalIgnoreCase))
            {
                boards.Add(ownBoard);
            }

            this.thisFillingChoices = true;

            try
            {
                this.BoardComboBox.ItemsSource = boards;
                this.BoardComboBox.SelectedItem = boards.FirstOrDefault(name => string.Equals(name, keep, StringComparison.OrdinalIgnoreCase))
                    ?? boards.FirstOrDefault();
            }
            finally
            {
                this.thisFillingChoices = false;
            }

            this.FollowBoardRegions();
        }

        // The board picked, as "{hardware}|{board}" - empty while either is not picked.
        private string SelectedBoardKey()
        {
            string hardware = this.HardwareComboBox.SelectedItem as string ?? string.Empty;
            string board = this.BoardComboBox.SelectedItem as string ?? string.Empty;

            return hardware.Length == 0 || board.Length == 0 ? string.Empty : $"{hardware}|{board}";
        }

        // ###########################################################################################
        // Region shown as the board picked comes: an answer known shows at once, one given at once
        // (the board's data read already) too; any other is waited for with the form as it is, and
        // taken only while that board is still the one picked.
        // ###########################################################################################
        private void FollowBoardRegions()
        {
            string boardKey = this.SelectedBoardKey();

            if (boardKey.Length == 0)
            {
                this.ShowRegion([]);
                return;
            }

            if (this.thisRegionsByBoard.TryGetValue(boardKey, out IReadOnlyList<string>? known))
            {
                this.ShowRegion(known);
                return;
            }

            Task<IReadOnlyList<string>> asking;

            try
            {
                asking = this.thisBoardRegions(boardKey);
            }
            catch (Exception ex)
            {
                asking = Task.FromException<IReadOnlyList<string>>(ex);
            }

            if (asking.IsCompleted)
            {
                this.TakeRegionsAnswer(boardKey, AnswerOf(boardKey, asking));
                return;
            }

            this.RegionsFollowedForTests = this.TakeRegionsAnswerWhenGivenAsync(boardKey, asking);
        }

        // The answer taken on the UI thread once given - the await comes back to it.
        private async Task TakeRegionsAnswerWhenGivenAsync(string boardKey, Task<IReadOnlyList<string>> asking)
        {
            IReadOnlyList<string> regions;

            try
            {
                regions = await asking;
            }
            catch (Exception ex)
            {
                regions = NoAnswer(boardKey, ex);
            }

            this.TakeRegionsAnswer(boardKey, regions);
        }

        // An answer given at once - one that failed is no regions.
        private static IReadOnlyList<string> AnswerOf(string boardKey, Task<IReadOnlyList<string>> asking) =>
            asking.IsCompletedSuccessfully
                ? asking.Result
                : NoAnswer(boardKey, asking.Exception?.GetBaseException() ?? new OperationCanceledException());

        // A question that failed is no regions - said in the log, the form left without the box.
        private static IReadOnlyList<string> NoAnswer(string boardKey, Exception ex)
        {
            Logger.Warning($"Could not tell which regions board [{boardKey}] comes in - [{ex.Message}]");
            return [];
        }

        private void TakeRegionsAnswer(string boardKey, IReadOnlyList<string> regions)
        {
            this.thisRegionsByBoard[boardKey] = regions;

            if (string.Equals(this.SelectedBoardKey(), boardKey, StringComparison.OrdinalIgnoreCase))
            {
                this.ShowRegion(regions);
            }
        }

        // ###########################################################################################
        // Shown with the board's regions listed (2026-11-09: names of its own since), or hidden - hidden, the region back
        // to what the device has stored (none for a new one). What is picked stays picked only where the new list has it
        // (code review, 2026-11-09: "Japan" picked for one board stayed picked on a PAL/NTSC board, listed after its regions,
        // and was saved so); the device's own region comes back on its own board or one that has it, else nothing is picked.
        // ###########################################################################################
        private void ShowRegion(IReadOnlyList<string> regions)
        {
            this.RegionPanel.IsVisible = regions.Count > 0;

            string own = this.thisEditing?.Region?.Trim() ?? string.Empty;

            bool Listed(string region) => regions.Contains(region, StringComparer.OrdinalIgnoreCase);

            bool ownBoard = this.IsDevicesOwnBoardPicked();

            string picked =
                regions.Count == 0 ? own
                : this.RegionComboBox.SelectedItem is string shown && shown != NoRegion && Listed(shown) ? shown
                : Listed(own) || ownBoard ? own
                : string.Empty;

            // The device's own region is offered beside the board's only on its own board, or while the box is hidden (second
            // code review, 2026-11-09: offered on a board without it, it could still be picked and saved there).
            this.ListRegions(regions.Count > 0 ? regions : BoardRegions.PalAndNtsc, picked, keepOwn: regions.Count == 0 || ownBoard);
        }

        // Whether the board picked is the one the device being edited is on.
        private bool IsDevicesOwnBoardPicked() =>
            this.thisEditing is { } device && string.Equals(device.BoardKey, this.SelectedBoardKey(), StringComparison.OrdinalIgnoreCase);

        // The drop-down's list: "Unknown", the regions, and the device's own when they do not have it (`keepOwn`) - then
        // `region` picked.
        private void ListRegions(IReadOnlyList<string> regions, string? region, bool keepOwn = true)
        {
            List<string> choices = [NoRegion, .. regions];
            string own = keepOwn ? this.thisEditing?.Region?.Trim() ?? string.Empty : string.Empty;

            foreach (string kept in new[] { own, region?.Trim() ?? string.Empty })
            {
                if (kept.Length > 0 && !choices.Contains(kept, StringComparer.OrdinalIgnoreCase))
                {
                    choices.Add(kept);
                }
            }

            if (this.RegionComboBox.ItemsSource is not IReadOnlyList<string> listed || !listed.SequenceEqual(choices, StringComparer.Ordinal))
            {
                this.RegionComboBox.ItemsSource = choices;
            }

            this.SelectRegion(region);
        }

        // `region` picked in the drop-down - "Unknown" for none, or one the list does not have.
        private void SelectRegion(string? region)
        {
            string wanted = region?.Trim() ?? string.Empty;
            IEnumerable<string> choices = this.RegionComboBox.ItemsSource as IEnumerable<string> ?? [];

            this.RegionComboBox.SelectedItem = wanted.Length == 0
                ? NoRegion
                : choices.FirstOrDefault(choice => string.Equals(choice, wanted, StringComparison.OrdinalIgnoreCase)) ?? NoRegion;
        }

        // ###########################################################################################
        // The regions the board `boardKey` names comes in, by BoardRegions' rule - read from
        // its board file as CRT views it (the draft's, unless "view boards as officially published"),
        // the file chosen as DataManager chooses it, through BoardDataReader's cache (the board shown
        // is held there already). Off the UI thread. A board CRT does not list, has no file for, or
        // cannot read, has none.
        // ###########################################################################################
        internal static Task<IReadOnlyList<string>> BoardRegionsAsViewedAsync(string boardKey)
        {
            if (Main.FindEntryForBoardKey(boardKey) is not { } entry || string.IsNullOrWhiteSpace(entry.ExcelDataFile))
            {
                return Task.FromResult<IReadOnlyList<string>>([]);
            }

            // Read here, on the UI thread: the pool reads no setting.
            string dataRoot = DataManager.DataRoot;
            string draftsRoot = DraftManager.DraftsRoot;
            bool viewOfficialOnly = UserSettings.ViewOfficialPublishedOnly;

            return Task.Run(async () =>
            {
                DraftBoardSource.ViewedBoard viewed = DraftBoardSource.ResolveViewed(
                    dataRoot, draftsRoot, entry.ExcelDataFile, viewOfficialOnly, entry.IsPublished);

                return File.Exists(viewed.BoardFile)
                    ? BoardRegions.Of(await BoardDataReader.LoadAsync(viewed.BoardFile, viewed.BoardFile))
                    : [];
            });
        }

        // ---- Condition ------------------------------------------------------------------------------

        // The five stars, each taking a click (ConditionStars draws them).
        private void BuildConditionStars()
        {
            StackPanel stars = ConditionStars.Build(this.thisCondition, size: 24);
            ToolTip.SetTip(stars, OwnedDeviceWording.ConditionHint);

            // The stars light up to the one under the pointer, as a rating does, and back to the condition set when the
            // pointer leaves them.
            stars.Background = Brushes.Transparent;
            stars.PointerExited += (_, _) => ConditionStars.Show(stars, this.thisCondition);

            foreach (Border star in stars.Children.OfType<Border>())
            {
                // A background, or the gaps between the points take no click.
                star.Background = Brushes.Transparent;
                star.Cursor = new Cursor(StandardCursorType.Hand);
                star.PointerEntered += (_, _) =>
                {
                    if (star.Tag is int number)
                    {
                        ConditionStars.Show(stars, number);
                    }
                };
                star.PointerPressed += (_, e) =>
                {
                    if (e.GetCurrentPoint(star).Properties.IsLeftButtonPressed && star.Tag is int number)
                    {
                        e.Handled = true;
                        this.ClickStar(number);
                    }
                };
            }

            this.ConditionStarsHost.Child = stars;
        }

        // Star `number` clicked: that many - the one set clicked again, none (OwnedDeviceCondition.AfterClick).
        internal void ClickStar(int number) => this.SetCondition(OwnedDeviceCondition.AfterClick(this.thisCondition, number));

        internal int Condition => this.thisCondition;

        private void SetCondition(int stars)
        {
            this.thisCondition = OwnedDeviceCondition.Within(stars);

            if (this.ConditionStarsHost.Child is Panel row)
            {
                ConditionStars.Show(row, this.thisCondition);
            }
        }

        // ---- Comments and photos --------------------------------------------------------------------

        // The comments and photos as they now are - the device's as saved, or the ones held for a new one.
        private void ShowCommentsAndPhotos()
        {
            IEnumerable<WorklogCommentRecord> comments = this.thisListed?.Comments ?? (IEnumerable<WorklogCommentRecord>)this.thisHeldComments;

            List<WorklogCommentRow> commentRows = comments
                .Where(comment => !string.IsNullOrWhiteSpace(comment.Text))
                .OrderByDescending(comment => comment.Date)
                .ThenByDescending(comment => comment.Id)
                .Select(comment => new WorklogCommentRow
                {
                    Id = comment.Id,
                    // Markup: trimmed so its first line keeps the indent that says what it is (code review, 2026-11-08).
                    Text = RichNote.TrimMarkup(comment.Text),
                    DateText = OwnedDeviceWording.CommentDate(comment.Date),
                })
                .ToList();

            this.CommentsList.ItemsSource = commentRows;
            this.CommentsHeadingText.Text = OwnedDeviceWording.CommentsHeading(commentRows.Count);
            this.NoCommentsText.IsVisible = commentRows.Count == 0;

            IEnumerable<WorklogLinkRecord> links = this.thisListed?.Links ?? (IEnumerable<WorklogLinkRecord>)this.thisHeldLinks;

            List<WorklogLinkRow> linkRows = links
                .Where(link => !string.IsNullOrWhiteSpace(link.Url))
                .Select(link => new WorklogLinkRow
                {
                    Id = link.Id,
                    Headline = string.IsNullOrWhiteSpace(link.Headline) ? link.Url.Trim() : link.Headline.Trim(),
                    Url = link.Url.Trim(),
                })
                .ToList();

            this.LinksList.ItemsSource = linkRows;
            this.LinksHeadingText.Text = OwnedDeviceWording.LinksHeading(linkRows.Count);
            this.NoLinksText.IsVisible = linkRows.Count == 0;

            List<WorklogAttachmentRow> photoRows = this.thisListed is { } device
                ? device.Photos
                    .OrderBy(photo => photo.DisplayOrder)
                    .ThenBy(photo => photo.Id)
                    .Select(photo => this.PhotoRow(photo.Id, photo.FileName, photo.Comment, OwnedDeviceStore.PhotoPath(device.Id, photo.FileName)))
                    .ToList()
                : this.thisHeldPhotos
                    .Select(photo => this.PhotoRow(photo.Id, Path.GetFileName(photo.SourcePath), photo.Comment, photo.SourcePath))
                    .ToList();

            // Changed in place - the drag holds this collection - and only once no drag is under way in it.
            this.thisPhotoDrag.Reset();
            this.thisPhotoRows.Clear();

            foreach (WorklogAttachmentRow row in photoRows)
            {
                this.thisPhotoRows.Add(row);
            }

            this.PhotosHeadingText.Text = OwnedDeviceWording.PhotosHeading(photoRows.Count);
            this.NoPhotosText.IsVisible = photoRows.Count == 0;
        }

        // A photo's row: its thumbnail decoded small - none for a file that is not there or no picture.
        private WorklogAttachmentRow PhotoRow(int id, string fileName, string? comment, string? path)
        {
            var row = new WorklogAttachmentRow
            {
                Id = id,
                FileName = fileName,
                DisplayFileName = fileName,
                Comment = comment ?? string.Empty,
            };

            this.ShowThumbnail(row, path);
            return row;
        }

        // Run on the pool before each thumbnail is read - for a test to hold it there.
        internal Action? WhileThumbnailIsReadForTests { get; set; }

        // ###########################################################################################
        // *** EACH PHOTO DECODED SMALL OFF THE UI THREAD, ONCE (code review, 2026-11-08: six camera photos decoded on the
        // UI thread froze CRT's window as the form opened). *** The device's page's own read (DevicePhotoThumbnails) - the
        // rows are made again after every change, a drag's drop too - set on the row once in, the row saying meanwhile that
        // it is being read (IsThumbnailPending).
        // ###########################################################################################
        private void ShowThumbnail(WorklogAttachmentRow row, string? path)
        {
            Task<Bitmap?> reading = DevicePhotoThumbnails.Read(path, this.WhileThumbnailIsReadForTests);

            if (reading.IsCompleted)
            {
                row.Thumbnail = reading.IsCompletedSuccessfully ? reading.Result : null;
                return;
            }

            row.IsThumbnailPending = true;

            _ = reading.ContinueWith(
                read => Dispatcher.UIThread.Post(() =>
                {
                    row.Thumbnail = read.IsCompletedSuccessfully ? read.Result : null;
                    row.IsThumbnailPending = false;
                }),
                TaskScheduler.Default);
        }

        // ###########################################################################################
        // A photo dragged to another place: the rows' new order saved at once for an existing device (OwnedDeviceStore.
        // SetPhotoOrder), the held photos put in it for a new one.
        // ###########################################################################################
        private void PhotoMoved()
        {
            List<int> order = this.thisPhotoRows.Select(row => row.Id).ToList();

            if (this.thisListed is { } device)
            {
                this.TakeSaved(OwnedDeviceStore.SetPhotoOrder(device.Id, order), "The photos' order could not be saved - see the log for details.");
                return;
            }

            List<HeldPhoto> reordered = order
                .Select(id => this.thisHeldPhotos.FirstOrDefault(photo => photo.Id == id))
                .OfType<HeldPhoto>()
                .ToList();
            this.thisHeldPhotos.Clear();
            this.thisHeldPhotos.AddRange(reordered);
            this.ShowCommentsAndPhotos();
        }

        // A photo row's handle pressed - the drag armed, started only once the pointer moves.
        private void OnPhotoHandlePressed(object? sender, PointerPressedEventArgs e) => this.thisPhotoDrag.Press(sender, e);

        // Whether a photo is being dragged - for a test.
        internal bool IsDraggingPhotoForTests => this.thisPhotoDrag.IsDragging;

        // ---- Links of interest ----------------------------------------------------------------------

        private async void OnAddLinkClick(object? sender, RoutedEventArgs e)
        {
            try
            {
                await this.AddLinkAsync();
            }
            catch (Exception ex)
            {
                Logger.Critical($"Adding a link to a device failed - [{ex}]");
            }
        }

        // "Add link": the worklog's own small window - a headline and an address - then the link saved, or held for a
        // new device until it is added.
        internal async Task AddLinkAsync()
        {
            if (await this.AskLinkAsync(null) is not { } link)
            {
                return;
            }

            if (this.thisListed is { } device)
            {
                this.TakeSaved(OwnedDeviceStore.AddLink(device.Id, link.Headline, link.Url), "The link could not be saved - see the log for details.");
                return;
            }

            this.thisHeldLinks.Add(new WorklogLinkRecord
            {
                Id = RecordIds.Next(this.thisHeldLinks.Select(each => each.Id)),
                Headline = link.Headline.Trim(),
                Url = link.Url.Trim(),
            });

            this.ShowCommentsAndPhotos();
        }

        private async void OnEditLinkClick(object? sender, RoutedEventArgs e)
        {
            try
            {
                if (sender is Control { Tag: int id })
                {
                    await this.EditLinkAsync(id);
                }
            }
            catch (Exception ex)
            {
                Logger.Critical($"Changing a device's link failed - [{ex}]");
            }
        }

        internal async Task EditLinkAsync(int linkId)
        {
            IEnumerable<WorklogLinkRecord> links = this.thisListed?.Links ?? (IEnumerable<WorklogLinkRecord>)this.thisHeldLinks;

            if (links.FirstOrDefault(each => each.Id == linkId) is not { } link ||
                await this.AskLinkAsync((link.Headline, link.Url)) is not { } changed)
            {
                return;
            }

            if (this.thisListed is { } device)
            {
                this.TakeSaved(OwnedDeviceStore.UpdateLink(device.Id, linkId, changed.Headline, changed.Url), "The link could not be saved - see the log for details.");
                return;
            }

            link.Headline = changed.Headline.Trim();
            link.Url = changed.Url.Trim();
            this.ShowCommentsAndPhotos();
        }

        private void OnDeleteLinkClick(object? sender, RoutedEventArgs e)
        {
            if (sender is Control { Tag: int id })
            {
                this.DeleteLink(id);
            }
        }

        internal void DeleteLink(int linkId)
        {
            if (this.thisListed is { } device)
            {
                this.TakeSaved(OwnedDeviceStore.RemoveLink(device.Id, linkId), "The link could not be deleted - see the log for details.");
                return;
            }

            this.thisHeldLinks.RemoveAll(link => link.Id == linkId);
            this.ShowCommentsAndPhotos();
        }

        // A link's row clicked: its address opened - ExternalTargetLauncher, the one way out of CRT, as a worklog's.
        private void OnLinkRowPointerPressed(object? sender, PointerPressedEventArgs e)
        {
            if (sender is Control { Tag: string url } row && !string.IsNullOrWhiteSpace(url) &&
                e.GetCurrentPoint(row).Properties.IsLeftButtonPressed)
            {
                // As typed, opened by LinkTarget's rule - "www.c64-wiki.com" too (code review, 2026-11-08).
                ExternalTargetLauncher.OpenLinkOrSay(url, this.OpenLinkOverrideForTests);
            }
        }

        // The worklog's "Add link" window - on a link's headline and address to change it.
        private async Task<(string Headline, string Url)?> AskLinkAsync((string Headline, string Url)? link)
        {
            if (this.AskLinkOverrideForTests is { } ask)
            {
                return await ask(link);
            }

            var dialog = new WorklogAddLinkWindow();

            if (link is { } existing)
            {
                dialog.InitializeForEdit(existing.Headline, existing.Url);
            }

            return await dialog.ShowDialog<(string Headline, string Url)?>(this);
        }

        // An existing device's change saved: listed as it now is, and said to the caller.
        private void TakeSaved(OwnedDeviceRecord? saved, string failure)
        {
            if (saved is null)
            {
                this.ShowProblem(failure);
                return;
            }

            this.thisListed = saved;
            this.ChangedOnDisk = true;
            this.HideProblem();
            this.ShowCommentsAndPhotos();
        }

        private async void OnAddCommentClick(object? sender, RoutedEventArgs e)
        {
            try
            {
                await this.AddCommentAsync();
            }
            catch (Exception ex)
            {
                Logger.Critical($"Adding a comment to a device failed - [{ex}]");
            }
        }

        // "Add comment": the worklog's own small window, then the comment saved - or held, for a new device.
        internal async Task AddCommentAsync()
        {
            if (await this.AskCommentAsync(null) is not { Length: > 0 } text)
            {
                return;
            }

            if (this.thisListed is { } device)
            {
                this.TakeSaved(OwnedDeviceStore.AddComment(device.Id, text, DateTime.Now), "The comment could not be saved - see the log for details.");
                return;
            }

            this.thisHeldComments.Add(new WorklogCommentRecord
            {
                Id = RecordIds.Next(this.thisHeldComments.Select(comment => comment.Id)),
                Text = text,
                Date = DateTime.Now,
            });

            this.ShowCommentsAndPhotos();
        }

        // A comment's row clicked: changed in the same small window.
        private async void OnCommentRowPointerPressed(object? sender, PointerPressedEventArgs e)
        {
            try
            {
                if (sender is Control { Tag: int id } row && e.GetCurrentPoint(row).Properties.IsLeftButtonPressed)
                {
                    await this.EditCommentAsync(id);
                }
            }
            catch (Exception ex)
            {
                Logger.Critical($"Changing a device's comment failed - [{ex}]");
            }
        }

        internal async Task EditCommentAsync(int commentId)
        {
            IEnumerable<WorklogCommentRecord> comments = this.thisListed?.Comments ?? (IEnumerable<WorklogCommentRecord>)this.thisHeldComments;

            if (comments.FirstOrDefault(comment => comment.Id == commentId) is not { } comment ||
                await this.AskCommentAsync(comment.Text) is not { Length: > 0 } text ||
                string.Equals(text, RichNote.TrimMarkup(comment.Text), StringComparison.Ordinal))
            {
                return;
            }

            if (this.thisListed is { } device)
            {
                this.TakeSaved(OwnedDeviceStore.UpdateComment(device.Id, commentId, text), "The comment could not be saved - see the log for details.");
                return;
            }

            comment.Text = text;
            this.ShowCommentsAndPhotos();
        }

        private void OnDeleteCommentClick(object? sender, RoutedEventArgs e)
        {
            if (sender is Control { Tag: int id })
            {
                this.DeleteComment(id);
            }
        }

        internal void DeleteComment(int commentId)
        {
            if (this.thisListed is { } device)
            {
                this.TakeSaved(OwnedDeviceStore.RemoveComment(device.Id, commentId), "The comment could not be deleted - see the log for details.");
                return;
            }

            this.thisHeldComments.RemoveAll(comment => comment.Id == commentId);
            this.ShowCommentsAndPhotos();
        }

        // The worklog's "Add comment" window - on `text` to change it - the comment typed, or null. Its markup trimmed
        // as every note is (RichNote.TrimMarkup).
        private async Task<string?> AskCommentAsync(string? text)
        {
            if (this.AskCommentOverrideForTests is { } ask)
            {
                return await ask(text) is { } answered ? RichNote.TrimMarkup(answered) : null;
            }

            var dialog = new WorklogAddCommentWindow();

            if (text is not null)
            {
                dialog.InitializeForEdit(text);
            }

            return await dialog.ShowDialog<string?>(this) is { } typed ? RichNote.TrimMarkup(typed) : null;
        }

        private async void OnAddPhotoClick(object? sender, RoutedEventArgs e)
        {
            try
            {
                await this.AddPhotoAsync();
            }
            catch (Exception ex)
            {
                Logger.Critical($"Adding a photo to a device failed - [{ex}]");
            }
        }

        // "Add photo": the worklog's own small window - a picture and its comment - then the photo copied
        // into the device's folder, or held for a new device until it is added.
        internal async Task AddPhotoAsync()
        {
            if (await this.AskPhotoAsync(null) is not { SourcePath: { Length: > 0 } path } result)
            {
                return;
            }

            if (this.thisListed is { } device)
            {
                bool added = await this.CopyPhotosInAsync(() => OwnedDeviceStore.AddPhotos(device.Id, [new OwnedDevicePhotoSource(path, result.Comment)]).Count > 0);
                this.TakeSaved(
                    added ? OwnedDeviceStore.Get(device.Id) : null,
                    "The photo could not be added - it is no picture CRT can show, open in another program, or no longer where it was picked.");
                return;
            }

            this.thisHeldPhotos.Add(new HeldPhoto(
                RecordIds.Next(this.thisHeldPhotos.Select(photo => photo.Id)),
                path,
                RichNote.TrimMarkup(result.Comment)));

            this.ShowCommentsAndPhotos();
        }

        private async void OnEditPhotoClick(object? sender, RoutedEventArgs e)
        {
            try
            {
                if (sender is Control { Tag: int id })
                {
                    await this.EditPhotoAsync(id);
                }
            }
            catch (Exception ex)
            {
                Logger.Critical($"Changing a device's photo failed - [{ex}]");
            }
        }

        // A photo's comment changed - and its picture, when another is picked.
        internal async Task EditPhotoAsync(int photoId)
        {
            if ((this.PhotosList.ItemsSource as IEnumerable<WorklogAttachmentRow>)?.FirstOrDefault(row => row.Id == photoId) is not { } row ||
                await this.AskPhotoAsync(row) is not { } result)
            {
                return;
            }

            if (this.thisListed is { } device)
            {
                // Another picture picked is copied in; a comment alone changed is no copy, and no wait.
                bool saved = string.IsNullOrWhiteSpace(result.SourcePath)
                    ? OwnedDeviceStore.UpdatePhoto(device.Id, photoId, result.SourcePath, result.Comment)
                    : await this.CopyPhotosInAsync(() => OwnedDeviceStore.UpdatePhoto(device.Id, photoId, result.SourcePath, result.Comment));
                this.TakeSaved(saved ? OwnedDeviceStore.Get(device.Id) : null, "The photo could not be changed - see the log for details.");
                return;
            }

            int index = this.thisHeldPhotos.FindIndex(photo => photo.Id == photoId);

            if (index >= 0)
            {
                HeldPhoto held = this.thisHeldPhotos[index];
                this.thisHeldPhotos[index] = held with
                {
                    SourcePath = string.IsNullOrWhiteSpace(result.SourcePath) ? held.SourcePath : result.SourcePath,
                    Comment = RichNote.TrimMarkup(result.Comment),
                };
            }

            this.ShowCommentsAndPhotos();
        }

        private void OnDeletePhotoClick(object? sender, RoutedEventArgs e)
        {
            if (sender is Control { Tag: int id })
            {
                this.DeletePhoto(id);
            }
        }

        internal void DeletePhoto(int photoId)
        {
            if (this.thisListed is { } device)
            {
                bool removed = OwnedDeviceStore.RemovePhoto(device.Id, photoId);
                this.TakeSaved(removed ? OwnedDeviceStore.Get(device.Id) : null, "The photo could not be deleted - see the log for details.");
                return;
            }

            this.thisHeldPhotos.RemoveAll(photo => photo.Id == photoId);
            this.ShowCommentsAndPhotos();
        }

        // ###########################################################################################
        // *** PHOTOS ARE COPIED IN OFF THE UI THREAD, UNDER THE FORM'S "PLEASE WAIT" (code review, 2026-11-08; CLAUDE.md:
        // "Photos are copied in, and a device deleted, off the UI thread under the BusyOverlay"). *** A folder of camera
        // photos copied on the UI thread froze CRT's window with nothing said. The form cannot be closed meanwhile
        // (OnClosing): a new device's photos would land on a device its caller was told was never added.
        // ###########################################################################################
        private bool thisCopyingPhotos;

        // Run on the pool before photos are copied in - for a test to hold the copy there.
        internal Action? WhilePhotosAreCopiedForTests { get; set; }

        private async Task<T> CopyPhotosInAsync<T>(Func<T> copy)
        {
            Action? whileCopied = this.WhilePhotosAreCopiedForTests;
            this.thisCopyingPhotos = true;

            try
            {
                return await BusyOverlay.RunLocalAsync(this, CrtWaitWording.AddingDevicePhotos, () => Task.Run(() =>
                {
                    whileCopied?.Invoke();
                    return copy();
                }));
            }
            finally
            {
                this.thisCopyingPhotos = false;
            }
        }

        protected override void OnClosing(WindowClosingEventArgs e)
        {
            if (this.thisCopyingPhotos)
            {
                e.Cancel = true;
            }

            base.OnClosing(e);
        }

        // The worklog's "Add photo" window - on a photo's row to change it.
        private async Task<WorklogAddPhotoWindow.PhotoResult?> AskPhotoAsync(WorklogAttachmentRow? row)
        {
            if (this.AskPhotoOverrideForTests is { } ask)
            {
                return await ask(row);
            }

            var dialog = new WorklogAddPhotoWindow();

            if (row is not null)
            {
                dialog.InitializeForEdit(row.DisplayFileName, row.Comment, this.PhotoPathOf(row.Id));
            }

            return await dialog.ShowDialog<WorklogAddPhotoWindow.PhotoResult?>(this);
        }

        // Where a listed photo lies: in the device's folder, or - held for a new device - where it was picked.
        private string? PhotoPathOf(int photoId)
        {
            if (this.thisListed is { } device)
            {
                return device.Photos.FirstOrDefault(photo => photo.Id == photoId) is { } photo
                    ? OwnedDeviceStore.PhotoPath(device.Id, photo.FileName)
                    : null;
            }

            return this.thisHeldPhotos.FirstOrDefault(photo => photo.Id == photoId)?.SourcePath;
        }

        // ###########################################################################################
        // A photo's thumbnail clicked: the photo whole, in the worklog's viewer, which zooms with the wheel
        // (owner request, 2026-11-06). The click goes no further.
        // ###########################################################################################
        private void OnPhotoThumbnailPointerPressed(object? sender, PointerPressedEventArgs e)
        {
            if (sender is Control { Tag: int id } thumbnail && e.GetCurrentPoint(thumbnail).Properties.IsLeftButtonPressed)
            {
                e.Handled = true;
                this.ViewPhoto(id);
            }
        }

        internal void ViewPhoto(int photoId)
        {
            if ((this.PhotosList.ItemsSource as IEnumerable<WorklogAttachmentRow>)?.FirstOrDefault(row => row.Id == photoId) is not { } row)
            {
                return;
            }

            var viewer = new WorklogPhotoViewerWindow();
            viewer.Initialize(row.DisplayFileName, row.Comment, this.PhotoPathOf(photoId));

            if (this.ShowPhotoViewerOverrideForTests is { } show)
            {
                show(viewer);
                return;
            }

            _ = viewer.ShowDialog(this);
        }

        // ---- Save and cancel ------------------------------------------------------------------------

        // What the form holds, as the store takes it.
        internal OwnedDeviceDetails Details()
        {
            string region = this.RegionComboBox.SelectedItem as string ?? NoRegion;

            return new OwnedDeviceDetails(
                this.NameTextBox.Text?.Trim() ?? string.Empty,
                this.SelectedBoardKey(),
                region == NoRegion ? string.Empty : region,
                this.SerialTextBox.Text?.Trim() ?? string.Empty,
                RichNote.TrimMarkup(this.DescriptionTextBox.Markup),
                this.StatusComboBox.SelectedItem as string ?? OwnedDeviceStatus.Unknown,
                this.thisCondition);
        }

        private void OnNameTextChanged(object? sender, TextChangedEventArgs e) => this.UpdateSaveEnabled();

        private void UpdateSaveEnabled() =>
            this.SaveButton.IsEnabled = !string.IsNullOrWhiteSpace(this.NameTextBox.Text);

        // ###########################################################################################
        // Escape cancels and Enter saves - but Enter in a box that takes lines of its own (the
        // description) is a new line there, so Ctrl+Enter saves from it. Enter on a BUTTON is that
        // button's (code review, 2026-10-28: Enter on a focused Cancel saved the form). Neither key
        // is the form's while a drop-down list is open: Enter picks in it, Escape closes it (code
        // review, 2026-11-08: Escape there cancelled the form, everything typed and held for a new
        // device thrown away, nothing asked).
        // ###########################################################################################
        private void OnWindowPreviewKeyDown(object? sender, KeyEventArgs e)
        {
            if (this.IsADropDownOpen)
            {
                return;
            }

            if (e.Key == Key.Escape)
            {
                this.Close(null);
                e.Handled = true;
            }
            else if (e.Key == Key.Enter && e.Source is not Button &&
                     (e.KeyModifiers.HasFlag(KeyModifiers.Control) || !IsInBoxTakingLines(e.Source)))
            {
                this.StartSave();
                e.Handled = true;
            }
        }

        private bool IsADropDownOpen =>
            this.HardwareComboBox.IsDropDownOpen || this.BoardComboBox.IsDropDownOpen ||
            this.RegionComboBox.IsDropDownOpen || this.StatusComboBox.IsDropDownOpen;

        private static bool IsInBoxTakingLines(object? source) =>
            source is Visual visual && visual.FindAncestorOfType<TextBox>(includeSelf: true) is { AcceptsReturn: true };

        private void OnCancelClick(object? sender, RoutedEventArgs e) => this.Close(null);

        private void OnSaveClick(object? sender, RoutedEventArgs e) => this.StartSave();

        // ###########################################################################################
        // The save "Add device" or Enter started - done at once unless a new device's photos are copied in, which waits
        // under the form's "please wait" (CopyPhotosInAsync). One at a time: Enter pressed again meanwhile starts none.
        // ###########################################################################################
        private Task thisSaving = Task.CompletedTask;

        // The save last started - for a test, to wait for a new device's photos to be copied in.
        internal Task SavingForTests => this.thisSaving;

        private void StartSave()
        {
            if (!this.thisSaving.IsCompleted)
            {
                return;
            }

            this.thisSaving = this.SaveAsync();
        }

        // ###########################################################################################
        // The fields written - an existing device's over it (its comments and photos are on disk already),
        // a new one created with the comments held for it, then its photos copied in, off the UI thread. A
        // photo that could not be copied leaves the device added, and is counted for the caller to say.
        // ###########################################################################################
        internal async Task SaveAsync()
        {
            try
            {
                await this.SaveCoreAsync();
            }
            catch (Exception ex)
            {
                // Said on the form, beside "Add device" - not on CRT's error card too (code review, 2026-11-08).
                Logger.Critical($"Saving a device failed - [{ex}]", alreadyShown: true);
                this.ShowProblem("Could not save the device - see the log for details.");
            }
        }

        private async Task SaveCoreAsync()
        {
            OwnedDeviceDetails details = this.Details();

            if (details.Name.Length == 0)
            {
                this.ValidationText.Text = OwnedDeviceWording.NameNeeded;
                this.ValidationText.IsVisible = true;
                this.NameTextBox.Focus();
                return;
            }

            OwnedDeviceRecord? saved = this.thisEditing is { } device
                ? OwnedDeviceStore.Update(device.Id, details)
                : OwnedDeviceStore.Create(details, this.thisHeldComments, this.thisHeldLinks);

            if (saved is null)
            {
                // Kept open with what was typed - closing would read as cancelled.
                this.ShowProblem("Could not save the device - see the log for details.");
                return;
            }

            if (this.thisEditing is null && this.thisHeldPhotos.Count > 0)
            {
                int deviceId = saved.Id;
                List<OwnedDevicePhotoSource> photos = this.thisHeldPhotos.Select(photo => new OwnedDevicePhotoSource(photo.SourcePath, photo.Comment)).ToList();

                this.PhotosAdded = await this.CopyPhotosInAsync(() => OwnedDeviceStore.AddPhotos(deviceId, photos).Count);
                this.PhotosNotAdded = this.thisHeldPhotos.Count - this.PhotosAdded;
                saved = OwnedDeviceStore.Get(deviceId) ?? saved;
            }

            this.Saved = saved;
            this.Close(saved);
        }

        private void ShowProblem(string text)
        {
            this.ProblemText.Text = text;
            this.ProblemText.IsVisible = true;
        }

        private void HideProblem() => this.ProblemText.IsVisible = false;
    }
}
