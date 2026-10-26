using Avalonia;
using Avalonia.Controls;
using Avalonia.Media.Imaging;
using Avalonia.Threading;
using Handlers.DataHandling;
using Handlers.Geometry;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Tabs.TabSchematics;

namespace CRT
{
    // ###########################################################################################
    // Hardware/board selection: the drop-downs themselves, loading a selected board's data,
    // board-key parsing/formatting, and the Configuration tab's catalogue-visibility tree
    // (hardware/board/schematic checkboxes) as it feeds back into those drop-downs. See
    // Main.axaml.cs for the file map of the whole partial class.
    // ###########################################################################################
    public partial class Main
    {
        // How many boards' schematic thumbnails are kept (SchematicThumbnailSet): the board shown
        // and the two before it, so going back and forth decodes nothing. A set is a few megabytes
        // per image at 800 pixels wide.
        internal const int ThumbnailSetsKept = 3;

        // The thumbnail sets by board (and viewed draft folder), each at the signature of the
        // images it was made from. Read on the pool, stored and evicted on the UI thread only.
        private readonly StampedLruCache<SchematicThumbnailSet> _thumbnailSets =
            new(ThumbnailSetsKept, StringComparer.OrdinalIgnoreCase);

        // The running board load's cancellation - cancelled by the next one (LoadSelectedBoardAsync).
        private CancellationTokenSource? _boardLoadCts;
        // ###########################################################################################
        // Populates the hardware drop-down with distinct hardware names from loaded data.
        // ###########################################################################################
        private void PopulateHardwareDropDown()
        {
            var hardwareNames = BuildVisibleHardwareNames(UserSettings.CatalogueUncheckedKeysSnapshot);

            this.HardwareComboBox.ItemsSource = hardwareNames;

            if (hardwareNames.Count == 0)
            {
                this.HardwareComboBox.SelectedIndex = -1;
                return;
            }

            var lastHardware = UserSettings.GetLastHardware();
            var savedIndex = hardwareNames.FindIndex(h =>
                string.Equals(h, lastHardware, StringComparison.OrdinalIgnoreCase));

            this.HardwareComboBox.SelectedIndex = savedIndex >= 0 ? savedIndex : 0;
        }

        // ###########################################################################################
        // Whether a hardware belongs in the Hardware drop-down: not itself unchecked, AND has at
        // least one board that is not itself unchecked. Schematic-level visibility is deliberately
        // NOT considered here - that would mean loading every board's Excel file just to populate a
        // drop-down, which the board dropdown/thumbnail filters already handle lazily per selection.
        // ###########################################################################################
        private static bool HasAnyVisibleBoard(IReadOnlySet<string> uncheckedKeys, string hardwareName)
        {
            if (!CatalogueVisibility.IsHardwareVisible(uncheckedKeys, hardwareName))
                return false;

            return DataManager.HardwareBoards.Any(entry =>
                string.Equals(entry.HardwareName, hardwareName, StringComparison.OrdinalIgnoreCase) &&
                CatalogueVisibility.IsBoardVisible(uncheckedKeys, hardwareName, entry.BoardName));
        }

        // ###########################################################################################
        // Rebuilds the hardware and board selectors after the main Excel data changed, while trying
        // to preserve the current selection when those entries still exist.
        // ###########################################################################################
        private void RefreshHardwareAndBoardSelectionsAfterMainExcelSync()
        {
            string previousHardware = this.HardwareComboBox.SelectedItem as string ?? string.Empty;
            string previousBoard = this.BoardComboBox.SelectedItem as string ?? string.Empty;

            this.PopulateHardwareDropDown();

            var hardwareNames = this.HardwareComboBox.ItemsSource?
                .Cast<string>()
                .ToList() ?? new List<string>();

            if (hardwareNames.Count == 0)
            {
                return;
            }

            int hardwareIndex = hardwareNames.FindIndex(h =>
                string.Equals(h, previousHardware, StringComparison.OrdinalIgnoreCase));

            this.HardwareComboBox.SelectedIndex = hardwareIndex >= 0 ? hardwareIndex : 0;

            var boardNames = this.BoardComboBox.ItemsSource?
                .Cast<string>()
                .ToList() ?? new List<string>();

            if (boardNames.Count == 0)
            {
                return;
            }

            int boardIndex = boardNames.FindIndex(b =>
                string.Equals(b, previousBoard, StringComparison.OrdinalIgnoreCase));

            this.BoardComboBox.SelectedIndex = boardIndex >= 0 ? boardIndex : 0;
        }

        // ###########################################################################################
        // Filters the board drop-down to only show boards belonging to the selected hardware.
        // ###########################################################################################
        private void OnHardwareSelectionChanged(object? sender, SelectionChangedEventArgs e)
        {
            // ApplyCatalogueVisibility is repopulating ItemsSource around a selection it restores to
            // the SAME value, so there is nothing here to do: the board list it would rebuild is
            // assigned by that caller itself, and clearing the component search or re-persisting
            // LastHardware would both be acting on a hardware change that is not happening.
            //
            // The guard sits at the top rather than only around the board rebuild BECAUSE of that -
            // every statement below is a response to the selection having changed. Anything added to
            // this method that must run even for an unchanged selection belongs ABOVE this line.
            if (this._suppressBoardHardwareSelectionReload)
                return;

            this._suppressComponentSearchRefresh = true;
            this.ComponentSearchTextBox.Text = string.Empty;
            this._suppressComponentSearchRefresh = false;

            var selectedHardware = this.HardwareComboBox.SelectedItem as string;

            var boards = this.BuildVisibleBoardNamesForSelectedHardware(
                UserSettings.CatalogueUncheckedKeysSnapshot);

            this.BoardComboBox.ItemsSource = boards;

            if (string.IsNullOrWhiteSpace(selectedHardware) || boards.Count == 0)
            {
                this.BoardComboBox.SelectedIndex = -1;
                return;
            }

            UserSettings.SetLastHardware(selectedHardware);

            // Honoured once and cleared, so the caller that set it gets the board it asked for
            // rather than this hardware's saved last board - see the field's own comment for why
            // going through the saved board costs an entire extra board load.
            var pendingBoard = this._pendingBoardSelectionOverride;
            this._pendingBoardSelectionOverride = null;

            var targetBoard = pendingBoard ?? UserSettings.GetLastBoardForHardware(selectedHardware);
            var targetIndex = boards.FindIndex(b =>
                string.Equals(b, targetBoard, StringComparison.OrdinalIgnoreCase));

            this.BoardComboBox.SelectedIndex = targetIndex >= 0 ? targetIndex : 0;
        }

        // ###########################################################################################
        // Handles board selection changes and loads the visible board UI first, then starts heavier
        // schematic/KiCad work in the background so the window can remain responsive immediately.
        // ###########################################################################################
        private async void OnBoardSelectionChanged(object? sender, SelectionChangedEventArgs e)
        {
            if (this._suppressBoardHardwareSelectionReload)
                return;

            // Wrapped for the reason TabWorkbooks.OpenEntryEditor documents: this is an "async void"
            // handler, so anything thrown - including by the long SYNCHRONOUS prologue in
            // LoadSelectedBoardAsync, which disposes every thumbnail bitmap and reaches into
            // TabSchematics' controls before the first await - is rethrown on the sync context with
            // no caller to catch it, and reaches App's global handler as a PROCESS-FATAL crash.
            //
            // The board load's own I/O is already guarded (BoardDataReader.LoadAsync catches and
            // returns null), so this is about everything else: a disposed bitmap, a renamed control,
            // a board whose Excel is mid-sync. Losing the whole application on a board change - the
            // single most-used action in the app - is the outcome being prevented.
            try
            {
                await this.LoadSelectedBoardAsync(clearComponentSearch: ReferenceEquals(sender, this.BoardComboBox));
            }
            catch (Exception ex)
            {
                Logger.Critical($"Loading the selected board failed: [{ex.Message}]");
            }
        }

        // ###########################################################################################
        // The board load itself, separated from the SelectionChanged handler so it can also be
        // invoked where no selection actually changed - see ApplyCatalogueVisibility, which needs a
        // reload when a schematic under the CURRENT board is hidden or shown and re-selecting the
        // same index raises no event.
        //
        // clearComponentSearch carries what the handler used to read off its sender: the component
        // search box is emptied for a board change the user made in the drop-down, but not for a
        // load triggered some other way.
        // ###########################################################################################
        private async Task LoadSelectedBoardAsync(bool clearComponentSearch)
        {
            this.TabSchematicsControl.CancelWorklogEntryMode();

            // The Workbooks tab's search box is deliberately NOT cleared here. It used to be, on
            // the reasoning that a board change is a change of subject - but in "Show all workbooks"
            // scope, clicking a search result on another board IS how the user follows the search,
            // and that click switches the board, so clearing here wiped the query at the exact
            // moment it had done its job. See TabWorkbooks.ClearSearch, which is now the only thing
            // that clears it. (OnHardwareSelectionChanged still clears ComponentSearchTextBox, which
            // is a different box with no cross-board results to follow.)

            this._suppressCategoryFilterSave = true;
            int loadVersion = unchecked(++this._boardSelectionLoadVersion);

            // The load this one supersedes stops decoding thumbnails (code review, 2026-10-13); it
            // used to run to its end and only then find it had been overtaken.
            this._boardLoadCts?.Cancel();
            var loadCts = new CancellationTokenSource();
            this._boardLoadCts = loadCts;
            CancellationToken loadToken = loadCts.Token;

            if (clearComponentSearch)
            {
                this._suppressComponentSearchRefresh = true;
                this.ComponentSearchTextBox.Text = string.Empty;
                this._suppressComponentSearchRefresh = false;
            }

            // No thumbnail drag or held press may outlive the tiles it is on, in any gallery.
            this.TabSchematicsControl.CancelThumbnailDrags();

            // Only the highlight layers are the thumbnails' own; the base thumbnails belong to the
            // thumbnail cache (_thumbnailSets), which disposes a set when it lets it go.
            foreach (var thumb in this.TabSchematicsControl.currentThumbnails)
            {
                if (!ReferenceEquals(thumb.ImageSource, thumb.BaseThumbnail))
                {
                    (thumb.ImageSource as IDisposable)?.Dispose();
                }

                (thumb.HighlightImage as IDisposable)?.Dispose();
                thumb.HighlightImage = null;
            }

            // Emptying the thumbnails leaves no schematic selected (TabSchematics.OnCurrentThumbnailsChanged).
            this.TabSchematicsControl.currentThumbnails.Clear();
            this.CategoryFilterListBox.ItemsSource = null;
            this.ComponentFilterListBox.ItemsSource = null;

            this.TabSchematicsControl.highlightIndexBySchematic = new(StringComparer.OrdinalIgnoreCase);
            this.TabSchematicsControl.schematicByName = new(StringComparer.OrdinalIgnoreCase);
            this.SetComponentHighlightRects(new(StringComparer.OrdinalIgnoreCase));

            this._currentBoardData = null;
            this._currentBoardDraftSummary = BoardDraftSummary.Empty;
            this.UpdateRegionButtonsState();
            this.PopulateBoardInfoSection(null, null);
            this.TabSchematicsControl.ResetSchematicsViewer();

            var selectedHardware = this.HardwareComboBox.SelectedItem as string;
            var selectedBoard = this.BoardComboBox.SelectedItem as string;

            if (string.IsNullOrEmpty(selectedHardware) || string.IsNullOrEmpty(selectedBoard))
            {
                // No board selected at all. _currentBoardData was cleared above, so refresh to the
                // empty state rather than leaving the bar and the Workbooks tab showing the board
                // that WAS selected a moment ago - stale worklog surfaces above a blank schematic
                // view read as the previous board still being loaded.
                this.NoteBoardOnScreen(null);
                this.RefreshWorklogBar();
                this.ShowBoardInContributeTab(null);
                return;
            }

            UserSettings.SetLastHardware(selectedHardware);
            UserSettings.SetLastBoardForHardware(selectedHardware, selectedBoard);

            string boardKey = this.GetCurrentBoardKey();

            this.ApplySchematicsSplitterRatioForCurrentBoard(boardKey);
            this.ApplyThumbnailsDetachedStateForBoardChange();

            var entry = DataManager.HardwareBoards.FirstOrDefault(ent =>
                string.Equals(ent.HardwareName, selectedHardware, StringComparison.OrdinalIgnoreCase) &&
                string.Equals(ent.BoardName, selectedBoard, StringComparison.OrdinalIgnoreCase));

            if (entry == null || string.IsNullOrWhiteSpace(entry.ExcelDataFile))
            {
                // Same reasoning as the no-selection case above: this board has no data file, so
                // there is nothing to load and _currentBoardData stays null. Refresh so the worklog
                // surfaces show THIS board (its workbook list does not need board data) with an
                // empty board pane, rather than the previous board's previews.
                this.NoteBoardOnScreen(null);
                this.RefreshWorklogBar();
                this.ShowBoardInContributeTab(null);
                return;
            }

            var boardData = await DataManager.LoadBoardDataAsync(entry);

            // A SUPERSEDED load returns without touching anything: the user has selected another
            // board since, and that newer load owns every surface now - refreshing here would put
            // this board's worklog state under the newer board's header.
            if (loadVersion != this._boardSelectionLoadVersion)
            {
                return;
            }

            if (boardData == null)
            {
                // The board's Excel file is missing or unreadable. Still the currently selected
                // board, so its worklog surfaces must show IT (empty board pane, real workbook
                // list) rather than whatever the previously selected board left on screen.
                this.NoteBoardOnScreen(null);
                this.RefreshWorklogBar();
                this.ShowBoardInContributeTab(null);
                return;
            }

            this._currentBoardData = boardData;

            // A schematic image renamed - in the board's published data or in this user's draft:
            // their own worklogs and settings that name it follow, before anything below reads them
            // (Main.SchematicRenames.cs).
            this.CatchUpOnSchematicRenames(entry, boardData);

            // A view of this board starts now, and counts after ten seconds (Main.BoardViews.cs) - a
            // reload of the board already on screen is not a new one.
            this.NoteBoardOnScreen(entry);

            // Read right after the await above, which is the call that set it - see this field's
            // own header comment for why this is snapshotted rather than read fresh at each of the
            // several places the component list gets rebuilt. Forced to Empty while "view as
            // officially published" is on: boardData above already has no draft merged into it in
            // that case (see DataManager.LoadBoardDataAsync), and a chip or tint marking rows as
            // drafted while the toggle claims "this is what everyone else sees" would contradict
            // the very thing the toggle promises.
            this._currentBoardDraftSummary = UserSettings.ViewOfficialPublishedOnly
                ? BoardDraftSummary.Empty
                : DataManager.LastLoadedDraftSummary;

            this.UpdateRegionButtonsState();
            this.PopulateBoardInfoSection(boardData.RevisionDate, boardData.Credits);

            // Costs nothing here: the draft's base revision and the official RevisionDate were both
            // published by the load that just finished (session 2d).
            this.EvaluateDraftDriftForCurrentBoard();

            // AFTER _currentBoardData is assigned, deliberately - this call used to sit above the
            // await, before the board data existed. Everything board-data-dependent in
            // RefreshWorklogBar was wrong in that first pass, not just the board pane:
            // RefreshSelectedSchematicEntries reset the selected schematic and drew the placeholder,
            // and SetShowWorklogEntriesList re-seeded the Schematics overlay against a
            // just-blanked highlight cache. It was patched with a second, narrow board-pane refresh
            // here; running the whole refresh once, in the right place, fixes the class of bug
            // rather than the one instance, and removes a full board-pane rebuild per board switch.
            //
            // Nothing above the await needs it back: the splitter-ratio restore reads only boardKey.
            this.RefreshWorklogBar();

            var categories = ComponentListBuilder.BuildDistinctCategories(boardData);
            this.CategoryFilterListBox.ItemsSource = categories;

            var savedCategories = UserSettings.GetSelectedCategories(boardKey);
            if (savedCategories == null)
            {
                try
                {
                    this.CategoryFilterListBox.SelectAll();
                }
                catch (OutOfMemoryException ex)
                {
                    Logger.Debug(ex, "Failed to apply default category selection - group was too large to select");
                }
            }
            else
            {
                for (int i = 0; i < categories.Count; i++)
                {
                    if (savedCategories.Contains(categories[i], StringComparer.OrdinalIgnoreCase))
                    {
                        this.CategoryFilterListBox.Selection.Select(i);
                    }
                }
            }

            this._suppressCategoryFilterSave = false;

            var activeCategories = new HashSet<string>(
                this.CategoryFilterListBox.SelectedItems?.Cast<string>() ?? Enumerable.Empty<string>(),
                StringComparer.OrdinalIgnoreCase);

            string searchTerm = this.ComponentSearchTextBox?.Text ?? string.Empty;
            var componentItems = ComponentListBuilder.BuildComponentItems(
                boardData, UserSettings.Region, activeCategories, searchTerm, this._currentBoardDraftSummary.DraftedComponentBoardLabels);

            this._suppressComponentHighlightUpdate = true;
            this.ComponentFilterListBox.ItemsSource = componentItems;

            if (!string.IsNullOrWhiteSpace(searchTerm))
            {
                // ItemsSource was just reassigned above; SelectAll can throw if Avalonia's
                // selection model has not yet caught up with the new item count.
                try
                {
                    this.ComponentFilterListBox.SelectAll();
                }
                catch
                {
                }
            }

            this._suppressComponentHighlightUpdate = false;

            // The VIEWED draft folder: empty under "view boards as officially published", so the
            // schematic images follow the workbook the toggle chose rather than the draft's.
            string draftBoardFolder = DraftManager.GetViewedBoardFolder(entry.ExcelDataFile);

            // ###########################################################################################
            // *** EVERYTHING THE POOL NEEDS IS READ HERE, ON THE UI THREAD (code review, 2026-10-13). ***
            // The background half read the hardware and board drop-downs (through the KiCad paths'
            // lookup) and the live hidden-images set the Configuration tab changes on the UI thread -
            // controls and a HashSet read off the UI thread while they could change. They are snapshots
            // now; the pool touches no control and builds no UI object (the thumbnails are decoded, not
            // rendered through an Image control).
            //
            // *** AND THE KiCad FILES ARE COPIED THERE (code review, 2026-10-25). *** Copied on the UI
            // thread, the first showing of a board froze the window for the seconds tens of megabytes
            // take; now they copy on the pool while the thumbnails decode (KiCadRawPathsFor).
            // ###########################################################################################
            var uncheckedKeys = UserSettings.CatalogueUncheckedKeysSnapshot;
            string region = UserSettings.Region;
            string dataRoot = DataManager.DataRoot;
            string excelDataFile = entry.ExcelDataFile;
            string thumbnailSetKey = $"{boardKey}\u001F{draftBoardFolder}";

            _ = Task.Run(async () =>
            {
                try
                {
                    Task<List<string>> rawPathsCopied = Task.Run(() => Main.KiCadRawPathsFor(dataRoot, draftBoardFolder, excelDataFile));

                    var highlightRects = HighlightRectBuilder.BuildHighlightRects(boardData, region);

                    bool IsSchematicVisible(string schematicName) =>
                        CatalogueVisibility.IsSchematicVisible(uncheckedKeys, selectedHardware!, selectedBoard!, schematicName);

                    var schematicByName = boardData.Schematics
                        .Where(schematic => !string.IsNullOrWhiteSpace(schematic.SchematicName))
                        .Where(schematic => IsSchematicVisible(schematic.SchematicName))
                        .ToDictionary(
                            schematic => schematic.SchematicName,
                            schematic => schematic,
                            StringComparer.OrdinalIgnoreCase);

                    // Each visible schematic's image file. One attached only in a local draft (not
                    // yet synced under Data/) is resolved by the same official-first/draft-fallback
                    // rule every other BoardData File value uses - see DraftFileResolver.
                    var images = boardData.Schematics
                        .Where(schematic => !string.IsNullOrWhiteSpace(schematic.SchematicImageFile))
                        .Where(schematic => IsSchematicVisible(schematic.SchematicName))
                        .Select(schematic => (
                            Name: schematic.SchematicName,
                            FullPath: DraftFileResolver.Resolve(dataRoot, draftBoardFolder, schematic.SchematicImageFile, entry.ExcelDataFile) ??
                                Path.Combine(dataRoot, schematic.SchematicImageFile.Replace('/', Path.DirectorySeparatorChar))))
                        .ToList();

                    // The board shown before, unchanged on disk: its thumbnails as they were. Otherwise
                    // decoded now, straight to thumbnail size, stopping if another board is chosen.
                    string signature = SchematicThumbnailSet.SignatureOf(images);
                    bool isNewSet = !this._thumbnailSets.TryGet(thumbnailSetKey, signature, out SchematicThumbnailSet? thumbnailSet);

                    if (isNewSet)
                    {
                        thumbnailSet = SchematicThumbnailSet.Decode(images, signature, AppConfig.ThumbnailMaxWidth, loadToken);
                    }

                    var thumbnails = thumbnailSet!.Images
                        .Select(image => new SchematicThumbnail
                        {
                            Name = image.Name,
                            ImageFilePath = image.FullPath,
                            BaseThumbnail = image.Thumbnail,
                            OriginalPixelSize = image.OriginalPixelSize,
                            ImageSource = image.Thumbnail,
                            VisualOpacity = 1.0,
                            IsMatchForSelection = false
                        })
                        .ToList();

                    List<string> rawPaths = await rawPathsCopied;

                    await Dispatcher.UIThread.InvokeAsync(() =>
                    {
                        if (loadVersion != this._boardSelectionLoadVersion)
                        {
                            // Overtaken: a set made for this load is nobody's - one from the cache
                            // stays there.
                            if (isNewSet)
                            {
                                thumbnailSet.Dispose();
                            }

                            return;
                        }

                        // The schematic image is about to appear while the KiCad project is still
                        // loading behind it, so flag that wait as soon as the board has KiCad data.
                        this.TabSchematicsControl.SetKiCadInitializingIndicatorVisible(rawPaths.Count > 0);

                        this.SetComponentHighlightRects(highlightRects);
                        this.TabSchematicsControl.schematicByName = schematicByName;
                        this.TabSchematicsControl.highlightIndexBySchematic = new(StringComparer.OrdinalIgnoreCase);

                        this.TabSchematicsControl.LoadSortedThumbnails(boardKey, thumbnails);

                        // The cache takes the new set only now, with the old board's thumbnails off
                        // screen (cleared when this load began): a set it lets go can be disposed.
                        if (isNewSet)
                        {
                            foreach (var released in this._thumbnailSets.Store(thumbnailSetKey, signature, thumbnailSet))
                            {
                                if (!ReferenceEquals(released, thumbnailSet))
                                {
                                    released.Dispose();
                                }
                            }
                        }

                        if (this.TabSchematicsControl.currentThumbnails.Count > 0)
                        {
                            string? savedSchematic = UserSettings.GetLastSchematicForBoard(boardKey);
                            var orderedThumbnails = this.TabSchematicsControl.currentThumbnails.ToList();

                            int savedIndex = string.IsNullOrEmpty(savedSchematic)
                                ? -1
                                : orderedThumbnails.FindIndex(thumbnail =>
                                    string.Equals(thumbnail.Name, savedSchematic, StringComparison.OrdinalIgnoreCase));

                            this.TabSchematicsControl.SelectedThumbnail =
                                orderedThumbnails[savedIndex >= 0 ? savedIndex : 0];
                        }

                        // Found by name in the board's file list, with the root it was found under -
                        // the launcher opens nothing outside it.
                        var localFiles = boardData.BoardLocalFiles.Select(file =>
                        {
                            DraftFileResolution? resolved = string.IsNullOrWhiteSpace(file.File)
                                ? null
                                : DraftFileResolver.ResolveWithSource(dataRoot, draftBoardFolder, file.File, entry.ExcelDataFile);

                            return new ResourceItem(
                                file.Category,
                                file.Name,
                                resolved?.FullPath ??
                                    (string.IsNullOrWhiteSpace(file.File)
                                        ? string.Empty
                                        : Path.Combine(dataRoot, file.File.Replace('/', Path.DirectorySeparatorChar))),
                                resolved?.Root);
                        });

                        var webLinks = boardData.BoardLinks.Select(link => new ResourceItem(
                            link.Category,
                            link.Name,
                            link.Url));

                        this.TabResources.LoadData(localFiles, webLinks);
                        this.TabOverview.LoadData(boardData);
                        this.ShowBoardInContributeTab(entry);
                        this.TabOverview.ApplyFilter(this.ComponentSearchTextBox?.Text ?? string.Empty);
                    }, DispatcherPriority.Background);

                    if (rawPaths.Count > 0)
                    {
                        await Dispatcher.UIThread.InvokeAsync(async () =>
                        {
                            if (loadVersion != this._boardSelectionLoadVersion)
                            {
                                return;
                            }

                            await this.TabSchematicsControl.LoadKiCadProjectForCurrentBoardAsync(rawPaths);
                        }, DispatcherPriority.Background);
                    }
                }
                catch (OperationCanceledException) when (loadToken.IsCancellationRequested)
                {
                    // Another board was chosen while this one's thumbnails were decoding.
                }
                catch (Exception ex)
                {
                    Logger.Warning($"Background board load failed - [{ex}]");
                }
            });
        }

        // ###########################################################################################
        // Handles component selection changes and drives highlight updates in both the main viewer
        // and all thumbnails.
        // ###########################################################################################
        private void OnComponentFilterSelectionChanged(object? sender, SelectionChangedEventArgs e)
        {
            if (this._suppressComponentHighlightUpdate)
                return;

            var boardLabels = this.ComponentFilterListBox.SelectedItems?
                .Cast<ComponentListItem>()
                .Select(item => item.BoardLabel)
                .Where(l => !string.IsNullOrEmpty(l))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList() ?? new List<string>();

            this.TabSchematicsControl.UpdateHighlightsForComponents(boardLabels);
            this.TabOverview.ApplyFilter(this.ComponentSearchTextBox?.Text ?? string.Empty);
        }

        // ###########################################################################################
        // Saves the selected category list for the current board whenever the user changes it.
        // ###########################################################################################
        private void OnCategoryFilterSelectionChanged(object? sender, SelectionChangedEventArgs e)
        {
            if (this._suppressCategoryFilterSave)
                return;

            var boardKey = this.GetCurrentBoardKey();
            if (string.IsNullOrEmpty(boardKey))
                return;

            var selected = this.CategoryFilterListBox.SelectedItems?
                .Cast<string>()
                .ToList() ?? new List<string>();

            UserSettings.SetSelectedCategories(boardKey, selected);

            if (this._currentBoardData != null)
            {
                var previouslySelectedKeys = new HashSet<string>(
                    this.ComponentFilterListBox.SelectedItems?.Cast<ComponentListItem>()
                        .Select(i => i.SelectionKey) ?? Enumerable.Empty<string>(),
                    StringComparer.OrdinalIgnoreCase);

                var categoryFilter = new HashSet<string>(selected, StringComparer.OrdinalIgnoreCase);
                var searchTerm = this.ComponentSearchTextBox?.Text ?? string.Empty;
                var componentItems = ComponentListBuilder.BuildComponentItems(
                    this._currentBoardData, this._localRegion, categoryFilter, searchTerm, this._currentBoardDraftSummary.DraftedComponentBoardLabels);

                this._suppressComponentHighlightUpdate = true;
                this.ComponentFilterListBox.ItemsSource = componentItems;

                if (!string.IsNullOrWhiteSpace(searchTerm))
                {
                    // ItemsSource was just reassigned above; SelectAll can throw if Avalonia's
                    // selection model has not yet caught up with the new item count.
                    try { this.ComponentFilterListBox.SelectAll(); } catch { }
                }
                else
                {
                    for (int i = 0; i < componentItems.Count; i++)
                    {
                        if (previouslySelectedKeys.Contains(componentItems[i].SelectionKey))
                            this.ComponentFilterListBox.Selection.Select(i);
                    }
                }

                this._suppressComponentHighlightUpdate = false;

                var survivingLabels = componentItems
                    .Where(item => previouslySelectedKeys.Contains(item.SelectionKey))
                    .Select(item => item.BoardLabel)
                    .Where(l => !string.IsNullOrEmpty(l))
                    .ToList();

                if (!string.IsNullOrWhiteSpace(searchTerm))
                {
                    survivingLabels = componentItems
                        .Select(item => item.BoardLabel)
                        .Where(l => !string.IsNullOrEmpty(l))
                        .ToList();
                }

                this.TabSchematicsControl.UpdateHighlightsForComponents(survivingLabels);
                this.TabOverview.ApplyFilter(searchTerm);
            }
        }

        // ###########################################################################################
        // Returns a composite key uniquely identifying the current hardware and board selection.
        // ###########################################################################################
        internal string GetCurrentBoardKey()
        {
            if (this.BoardKeyOverrideForTests != null)
                return this.BoardKeyOverrideForTests;

            var hw = this.HardwareComboBox.SelectedItem as string;
            var board = this.BoardComboBox.SelectedItem as string;
            if (string.IsNullOrEmpty(hw) || string.IsNullOrEmpty(board))
            {
                return string.Empty;
            }
            return $"{hw}|{board}";
        }

        // ###########################################################################################
        // Returns the currently selected hardware/board entry, or null if the selection is invalid.
        // ###########################################################################################
        internal HardwareBoardEntry? GetCurrentBoardEntry()
        {
            var selectedHardware = this.HardwareComboBox.SelectedItem as string;
            var selectedBoard = this.BoardComboBox.SelectedItem as string;

            if (string.IsNullOrWhiteSpace(selectedHardware) || string.IsNullOrWhiteSpace(selectedBoard))
            {
                return null;
            }

            return DataManager.HardwareBoards.FirstOrDefault(entry =>
                string.Equals(entry.HardwareName, selectedHardware, StringComparison.OrdinalIgnoreCase) &&
                string.Equals(entry.BoardName, selectedBoard, StringComparison.OrdinalIgnoreCase));
        }

        // ###########################################################################################
        // Finds the hardware/board entry a "{hardware}|{board}" composite key names, by matching
        // against DataManager.HardwareBoards rather than splitting the string - a hardware or board
        // name containing "|" would otherwise split wrong, and this way GetCurrentBoardKey stays the
        // one place that knows the composite's format.
        //
        // THE one reader of that format besides GetCurrentBoardKey itself: both
        // TryResolveHardwareAndBoardForBoardKey and FormatBoardKeyForDisplay go through here rather
        // than each running their own FirstOrDefault, so a change to the key's shape is a change in
        // exactly two places rather than three.
        // ###########################################################################################
        private static HardwareBoardEntry? FindEntryForBoardKey(string boardKey) =>
            DataManager.HardwareBoards.FirstOrDefault(e =>
                string.Equals($"{e.HardwareName}|{e.BoardName}", boardKey, StringComparison.OrdinalIgnoreCase));

        // ###########################################################################################
        // Reverses a board key into the two drop-down names, for the worklog picker jumping to a
        // workbook's board when it differs from the one on screen.
        // ###########################################################################################
        private static bool TryResolveHardwareAndBoardForBoardKey(string boardKey, out string hardwareName, out string boardName)
        {
            var entry = FindEntryForBoardKey(boardKey);

            hardwareName = entry?.HardwareName ?? string.Empty;
            boardName = entry?.BoardName ?? string.Empty;
            return entry != null;
        }

        // ###########################################################################################
        // Short "hardware/board" label for a workbook's BoardKey, for the worklog picker's item
        // labels - see WorklogJobBox.ItemTemplate in the constructor. Deliberately not
        // TryResolveHardwareAndBoardForBoardKey's HardwareName/BoardName (the full names from the
        // main Excel sheet, e.g. "Commodore 64" / "250469 (short board)") - those are too long once
        // several workbooks from different boards sit in one dropdown. HardwareBoardEntry's own
        // ShortHardwareBoardLabel reads the same short names straight from ExcelDataFile's folder
        // structure instead (e.g. "C64/250469") - see its own comment for why.
        //
        // Falls back to the raw key if the board no longer exists in the synced data, e.g. content
        // that was later removed from classic-repair-toolbox.dk.
        // ###########################################################################################
        internal static string FormatBoardKeyForDisplay(string boardKey) =>
            FindEntryForBoardKey(boardKey)?.ShortHardwareBoardLabel is { Length: > 0 } shortLabel
                ? shortLabel
                : boardKey;

        // ###########################################################################################
        // The FULL hardware and board names for a workbook's BoardKey, as the two drop-downs at the
        // top-left show them - "Commodore 128" and "310378 (C128 & C128D)".
        //
        // For the Workbooks tab's cards, which have a line of their own to spend on this and are
        // read while deciding which repair job to open. They used to show
        // FormatBoardKeyForDisplay's short folder-derived label ("C128/310378"), which is a
        // storage-layout detail rather than a name the user chose the board by - the drop-downs
        // never show it, so it had to be mentally mapped back. The short form is still right for
        // the worklog PICKER, where several boards sit in one dropdown row and the full names run
        // too long, so both formatters exist deliberately.
        //
        // The names come from the KEY ITSELF ("HardwareName|BoardName") rather than from the synced
        // data, so a workbook whose board was later removed from classic-repair-toolbox.dk still
        // names its board properly instead of falling back to a raw key with a pipe in it. The
        // resolved entry is preferred when there is one, since that is the authority on the current
        // names; the split is the fallback.
        //
        // Returns the two parts rather than one joined string: the card renders them on separate
        // lines, and joining here would make the caller split them again.
        // ###########################################################################################
        internal static (string Hardware, string Board) FormatBoardKeyAsFullNames(string boardKey)
        {
            var entry = FindEntryForBoardKey(boardKey);

            if (entry != null &&
                !string.IsNullOrWhiteSpace(entry.HardwareName) &&
                !string.IsNullOrWhiteSpace(entry.BoardName))
            {
                return (entry.HardwareName, entry.BoardName);
            }

            // "Hardware|Board" is the key's own shape - see FindEntryForBoardKey, which builds it.
            int separator = boardKey?.IndexOf('|') ?? -1;

            if (separator > 0 && separator < boardKey!.Length - 1)
            {
                return (boardKey[..separator], boardKey[(separator + 1)..]);
            }

            // Not a key of the expected shape at all (blank, or hand-edited): show whatever it is on
            // one line rather than an empty card row.
            return (boardKey ?? string.Empty, string.Empty);
        }

        // ###########################################################################################
        // Resolves the full path to the currently selected board Excel file.
        // ###########################################################################################
        internal string GetCurrentBoardExcelPath()
        {
            var entry = this.GetCurrentBoardEntry();
            if (entry == null || string.IsNullOrWhiteSpace(entry.ExcelDataFile))
            {
                return string.Empty;
            }

            return Path.Combine(DataManager.DataRoot, entry.ExcelDataFile.Replace('/', Path.DirectorySeparatorChar));
        }

        // ###########################################################################################
        // Resolves full paths to the raw KiCad files of a board - `draftFolder` the VIEWED draft
        // folder (DraftManager.GetViewedBoardFolder), so "view boards as officially published" shows
        // only the published KiCad data: a drafted import is not something everyone else sees yet.
        //
        // *** IN CRT 4.0 ONE BOARD FILE NAMES THEM ALL (2026-10-24). *** A draft is the whole board -
        // its file lists the published KiCad files as well as the ones imported into it - so with a
        // draft only the draft's are read, else the published board's. Their bytes are copied under
        // their names into CRT's cache, where the loaders find a project's sheets beside each other
        // (BoardKiCadFolder). The "Schematic images" window's CAD names read the same files.
        //
        // *** NEVER ON THE UI THREAD (code review, 2026-10-25). *** The first copy of a large board is
        // tens of megabytes; the caller reads the three paths on the UI thread and asks this on the
        // pool. It touches no control, so it is safe there.
        // ###########################################################################################
        internal static List<string> KiCadRawPathsFor(string dataRoot, string draftFolder, string excelDataFile) =>
            KiCadRawFileScanner.Scan(BoardKiCadFolder.ForBoard(dataRoot, draftFolder, excelDataFile, AppConfig.KiCadCacheFolder));

        // ###########################################################################################
        // Reloads the currently selected board from disk and restores the given schematic selection.
        //
        // *** IT ALSO RE-EVALUATES THE DRAFTS TAB, because every LOCAL EDIT arrives here
        // (owner report, 2026-09-23). *** Both paths that write a draft call this - the
        // Contribute window's component save and the label editor's - and a save is very often the
        // moment a board gains its FIRST draft. Without this the tab simply did not appear:
        // "I need to restart the app for it to show".
        //
        // The Drafts tab is not a user preference, it is a fact about the Drafts folder's contents,
        // so the place to refresh it is wherever that folder changes. This method is the one funnel
        // both editors already pass through - the same reasoning that makes Main.RefreshWorklogBar
        // the single place every worklog change is redrawn from, rather than each editor
        // remembering to poke the tab itself.
        // ###########################################################################################
        internal void ReloadCurrentBoardFromDisk(string schematicNameToRestore)
        {
            var boardKey = this.GetCurrentBoardKey();
            if (!string.IsNullOrWhiteSpace(boardKey) && !string.IsNullOrWhiteSpace(schematicNameToRestore))
            {
                UserSettings.SetLastSchematicForBoard(boardKey, schematicNameToRestore);
            }

            var entry = this.GetCurrentBoardEntry();
            if (entry != null && !string.IsNullOrWhiteSpace(entry.ExcelDataFile))
            {
                DataManager.ClearBoardCache(entry.ExcelDataFile);
            }

            // BEFORE the reload rather than after: the reload is fire-and-forget, so ordering this
            // behind it would race the board load for no benefit. Nothing here depends on the new
            // board data - it reads the Drafts folder.
            this.ApplyContributeTabState();

            // The Drafts tab's table, open on that draft, reads its file again if something wrote it
            // past the draft held in memory (CRT 4.0 step 1 - TabDrafts.ReadTableDraftAgain).
            this.TabDrafts.ReadTableDraftAgain();

            // Not a board CHANGE, so the component search box is deliberately left alone - the same
            // thing passing a null sender used to express before this became a typed argument.
            _ = this.LoadSelectedBoardAsync(clearComponentSearch: false);
        }

        // ###########################################################################################
        // Re-filters the Hardware/Board drop-downs against the Configuration tab's catalogue
        // visibility tree after a checkbox there changes, so an unchecked hardware/board disappears
        // immediately rather than only after the next board load.
        //
        // Reuses RefreshHardwareAndBoardSelectionsAfterMainExcelSync, which already preserves the
        // current selection where it is still valid and falls back to index 0 otherwise - exactly
        // what is needed here too. If the currently selected board's own visibility changed (now
        // hidden), the selection-index change that follows re-fires OnHardwareSelectionChanged/
        // OnBoardSelectionChanged through the normal SelectionChanged event, which rebuilds the
        // schematic thumbnails using the same filter.
        //
        // WHICH of those two happens is decided here, rather than by always doing both. That
        // refresh reassigns HardwareComboBox.ItemsSource, and reassigning it re-fires the whole
        // selection chain even when the list is identical - which means a full board reload: the
        // board Excel re-read, every schematic bitmap re-decoded, and (see
        // TabSchematics.ThumbnailsDetach.cs / ApplyThumbnailsDetachedStateForBoardChange) the detached
        // thumbnails window closed and reopened on whatever OTHER monitor it is sitting on. Running
        // that on EVERY checkbox toggle made ticking through a tree of hardware/boards, none of which
        // belong to the board on screen, cost a full reload - including that window flicker - each
        // time. Reported exactly that way: unchecking hardware unrelated to the active board still
        // refreshed the detached thumbnails window.
        //
        // So: a hardware/board key rebuilds the drop-downs' CONTENTS unconditionally (so a hidden
        // entry disappears from the list next time it is opened, and the board list picks up an
        // added/removed sibling board), but only RESELECTS - and so only reloads - when the currently
        // selected hardware or board no longer survives in the new lists. Repopulating the ItemsSource
        // while the current selection is still valid goes through _suppressBoardHardwareSelectionReload
        // so the momentary SelectedItem loss that ItemsSource reassignment causes cannot cascade into
        // OnHardwareSelectionChanged/OnBoardSelectionChanged. A SCHEMATIC key never touches a
        // drop-down at all - it only changes which thumbnails the current board shows, and then only
        // when the key belongs to the board on screen, which is the one case that still needs a
        // reload.
        // ###########################################################################################
        public void ApplyCatalogueVisibility(string changedKey)
        {
            var uncheckedKeys = UserSettings.CatalogueUncheckedKeysSnapshot;

            if (CatalogueVisibility.IsSchematicKey(changedKey))
            {
                // A schematic toggle only matters to the board currently on screen; the filter is
                // applied while a board loads, so any other board picks it up on its next load.
                if (CatalogueVisibility.KeyNamesBoard(changedKey, this.GetCurrentBoardKeyParts()))
                {
                    this.ReloadCurrentBoardForCatalogueVisibility();
                }

                return;
            }

            var previousHardwareNames = this.HardwareComboBox.ItemsSource?
                .Cast<string>()
                .ToList() ?? new List<string>();

            var previousBoardNames = this.BoardComboBox.ItemsSource?
                .Cast<string>()
                .ToList() ?? new List<string>();

            var hardwareNames = BuildVisibleHardwareNames(uncheckedKeys);
            var boardNames = this.BuildVisibleBoardNamesForSelectedHardware(uncheckedKeys);

            if (hardwareNames.SequenceEqual(previousHardwareNames, StringComparer.OrdinalIgnoreCase) &&
                boardNames.SequenceEqual(previousBoardNames, StringComparer.OrdinalIgnoreCase))
            {
                return;
            }

            var (currentHardware, currentBoard) = this.GetCurrentBoardKeyParts();

            bool currentSelectionStillValid =
                CatalogueVisibility.CurrentSelectionSurvives(hardwareNames, boardNames, currentHardware, currentBoard);

            if (!currentSelectionStillValid)
            {
                // The board or hardware actually on screen was hidden (or nothing was selected to
                // begin with) - the existing reselect-and-reload path is exactly what is needed here.
                this.RefreshHardwareAndBoardSelectionsAfterMainExcelSync();
                return;
            }

            // The current selection survives untouched: some OTHER hardware/board's checkbox changed.
            // Repopulate the drop-downs' items so the change is reflected next time either is opened,
            // but suppress the SelectionChanged cascade that reassigning ItemsSource would otherwise
            // raise for a selection that is not actually changing.
            //
            // Both drop-downs are assigned the SAME way - ItemsSource, then the live selected value
            // back. Deliberately NOT PopulateHardwareDropDown() for the hardware half, even though it
            // assigns the identical list: that method exists to apply the SAVED last hardware
            // (UserSettings.GetLastHardware, else index 0), which is a different value from the one on
            // screen whenever the two have diverged. The end result is the same either way, because
            // the SelectedItem line below immediately corrects it and the suppression flag stops
            // anything reacting in between - but only by luck of the ordering of those two lines.
            // Restoring what IS selected is the whole intent here, so it is what the code says,
            // rather than setting the wrong value and relying on the next statement to undo it.
            this._suppressBoardHardwareSelectionReload = true;
            try
            {
                this.HardwareComboBox.ItemsSource = hardwareNames;
                this.HardwareComboBox.SelectedItem = currentHardware;
                this.BoardComboBox.ItemsSource = boardNames;
                this.BoardComboBox.SelectedItem = currentBoard;
            }
            finally
            {
                this._suppressBoardHardwareSelectionReload = false;
            }
        }

        // ###########################################################################################
        // Re-runs the current board's load so a setting that changes what that load PRODUCES -
        // catalogue visibility filtering, or "view as officially published" (session 2b) - is
        // applied afresh. Re-selecting the same index raises no SelectionChanged, so the load is
        // invoked directly. The component search box is left alone: the board has not changed, and
        // the user was ticking a checkbox on another tab, not choosing a different board.
        //
        // Named for catalogue visibility historically; kept as the one shared reload rather than
        // adding a near-identical second private method for the toggle.
        // ###########################################################################################
        private void ReloadCurrentBoardForCatalogueVisibility()
        {
            if (this.BoardComboBox.SelectedItem is not string)
            {
                return;
            }

            _ = this.LoadSelectedBoardAsync(clearComponentSearch: false);
        }

        // ###########################################################################################
        // The selected hardware and board as the pair CatalogueVisibility compares a key against -
        // empty strings when nothing is selected, which no key can match.
        // ###########################################################################################
        private (string HardwareName, string BoardName) GetCurrentBoardKeyParts() =>
            (this.HardwareComboBox.SelectedItem as string ?? string.Empty,
             this.BoardComboBox.SelectedItem as string ?? string.Empty);

        private static List<string> BuildVisibleHardwareNames(IReadOnlySet<string> uncheckedKeys) =>
            DataManager.HardwareBoards
                .Select(e => e.HardwareName)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .Where(hardwareName => HasAnyVisibleBoard(uncheckedKeys, hardwareName))
                .ToList();

        private List<string> BuildVisibleBoardNamesForSelectedHardware(IReadOnlySet<string> uncheckedKeys)
        {
            if (this.HardwareComboBox.SelectedItem is not string selectedHardware)
            {
                return new List<string>();
            }

            return DataManager.HardwareBoards
                .Where(entry => string.Equals(entry.HardwareName, selectedHardware, StringComparison.OrdinalIgnoreCase))
                .Select(entry => entry.BoardName)
                .Where(b => !string.IsNullOrWhiteSpace(b))
                .Where(b => CatalogueVisibility.IsBoardVisible(uncheckedKeys, selectedHardware, b))
                .ToList();
        }
    }
}
