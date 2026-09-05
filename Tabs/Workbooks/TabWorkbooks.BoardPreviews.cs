using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Layout;
using Avalonia.Media;
using Avalonia.Media.Imaging;
using Handlers.DataHandling;
using Handlers.Geometry;
using Handlers.Theming;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Tabs.TabSchematics;

namespace CRT
{
    // ###########################################################################################
    // THE BOARD PANE (marker 3 in the mockup): every schematic image that has one or more worklog
    // entries belonging to the SELECTED workbook, each rendered with its own entries - a coloured,
    // dashed bounds rectangle for an entry with "show marked area" on, a plain "#N" pill (no
    // bounds) for one without. Real images, real entries, no sample data. Clicking a pill opens
    // that entry's editor, exactly as it does on the Schematics tab's own "Show worklogs" view -
    // see OnPreviewBadgePointerPressed. Clicking anywhere else on a preview SELECTS that schematic
    // (a highlighted border) and drives the entry list on the right - see SelectSchematic and
    // RefreshSelectedSchematicEntries.
    //
    // Reuses the exact same pieces TabSchematics.Worklog.cs uses for the "Show worklogs" list view
    // on the main Schematics tab: WorklogEntriesOverlay draws the bounds, and the badge is built to
    // the identical recipe (padlock glyph, category-coloured pill, white state disc) so an entry
    // looks the same wherever it appears in the app. What is NOT reused is the zoom/pan machinery -
    // this pane never zooms, so ViewMatrix is always the identity and the badge canvas needs no
    // inverse-scale transform or viewport-edge nudging.
    //
    // Rebuilt by RefreshBoardPreviews, called from SelectWorkbook, from the end of RefreshWorkbooks,
    // from RefreshBoardPreviewsForCurrentSelection (Main's public entry point - see that method's
    // own header for why board data needs a second, later refresh pass), and from
    // OnPreviewBadgePointerPressed after a save - so the board can never show a workbook other than
    // the one the top-line and the left panel agree on, nor a stale entry after it was edited.
    // ###########################################################################################
    public partial class TabWorkbooks
    {
        // Shown by BOTH the board pane and the entry list when the board has no worklogs at all,
        // which is why it is one constant rather than the same sentence typed in two places: the two
        // messages are on screen together, and WorkbooksBoardPreviewTests asserts they are identical,
        // so a divergence showed up as a failing test rather than as the UI telling anyone.
        //
        // Punctuated like every other empty state in this file - the two it sits beside ("...match
        // your search." and "...for this schematic yet.") both end in a full stop, and this one
        // briefly did not.
        private const string NoWorklogsForBoardMessage =
            "No worklogs recorded yet for any schematics in this board.";

        // The schematic currently selected in the board pane - drives the entry list on the right.
        // Null means none selected (before the first RefreshBoardPreviews call, or the workbook has
        // no entries at all). RefreshBoardPreviews defaults this to the first schematic it builds a
        // preview for whenever the current value no longer names one that is actually shown, the
        // same "stay selected if still valid, else fall back" rule RefreshWorkbooks applies to the
        // selected workbook - see SelectSchematic.
        private string? thisSelectedSchematicName;

        // Which board thisSelectedSchematicName was chosen against.
        //
        // Without this, the selection was only reset when its name was absent from the NEW board's
        // grouped entries - so a schematic name shared between two boards (common across Commodore
        // revisions: "Motherboard", "Sheet 1") carried a selection made against a different board
        // straight over. thisSelectedWorkbookId does not have this problem because it is re-derived
        // per board from the saved active id; this field had no such anchor.
        private string? thisSelectedSchematicBoardKey;

        // ###########################################################################################
        // One Hand cursor for every preview and every badge this pane builds.
        //
        // Cursor is IDisposable and holds an HCURSOR on Win32, and these were constructed per preview
        // AND per badge on every rebuild, against controls thrown away by the next one, with nothing
        // disposing them. Elsewhere in the app a per-call new Cursor(...) is assigned to a single
        // long-lived control, which bounds it; here the count grew with rebuilds. One static instance
        // for the process removes the question rather than adding disposal to two more places.
        // ###########################################################################################
        private static readonly Cursor HandCursor = new(StandardCursorType.Hand);

        // ###########################################################################################
        // Every schematic bitmap this pane has decoded, keyed by the absolute path it came from.
        //
        // WHO OWNS THESE: this tab, for as long as it is loaded. Nothing else may dispose one - and
        // that is the point of caching them rather than decoding per rebuild and disposing on clear.
        //
        // RefreshBoardPreviews clears BoardPreviewPanel and rebuilds every preview from scratch, and
        // it is reached from RefreshWorkbooks, SelectWorkbook, RefreshBoardPreviewsForCurrentSelection
        // and OnPreviewBadgePointerPressed - so Main.RefreshWorklogBar drives it on every board
        // change, entry save and workbook create/close. A fresh `new Bitmap(fullPath)` per preview
        // per pass stranded a full-resolution decode each time (a 4220x2941 schematic is ~47 MB of
        // BGRA), and the badge/SizeChanged closures below capture the bitmap, so the discarded
        // subtree kept itself alive through its own event tables rather than merely waiting for a
        // collection.
        //
        // Disposing on clear instead would be WORSE than the leak: ShowDialog does not block the
        // dispatcher, so RefreshWorklogBar can re-enter while a pill's editor is up, and that editor
        // documents that its schematic bitmap belongs to the caller. Disposing under it leaves
        // EditorLocationPreviewImage rendering a dead Skia surface - an ObjectDisposedException on
        // the render thread, which is fatal in Avalonia. Sharing one instance per path removes the
        // question: the same bitmap is handed to every preview and every editor, and nothing is
        // disposed until the tab itself goes away.
        //
        // Bounded by the number of distinct schematic images on the boards visited this session, not
        // by how often the pane is rebuilt, which is the difference that matters.
        // ###########################################################################################
        private readonly Dictionary<string, Bitmap> thisSchematicBitmapsByPath = new(StringComparer.OrdinalIgnoreCase);

        // ###########################################################################################
        // Returns the shared decoded bitmap for a schematic image path, decoding it on first use.
        // Null when the file is missing or cannot be decoded - one bad path drops that schematic
        // rather than the whole pane, the same failure mode DataManager already tolerates elsewhere.
        // ###########################################################################################
        private Bitmap? GetOrDecodeSchematicBitmap(string fullPath)
        {
            if (this.thisSchematicBitmapsByPath.TryGetValue(fullPath, out var cached))
                return cached;

            try
            {
                if (!File.Exists(fullPath))
                    return null;

                var decoded = new Bitmap(fullPath);
                this.thisSchematicBitmapsByPath[fullPath] = decoded;
                return decoded;
            }
            catch (Exception ex)
            {
                Logger.Warning($"Workbooks tab: could not load schematic image [{fullPath}] - [{ex.Message}]");
                return null;
            }
        }

        // ###########################################################################################
        // Drops every control that holds one of the cached bitmaps, so DisposeSchematicBitmaps can
        // run without leaving anything on screen pointing at a freed Skia surface.
        //
        // MUST be called before DisposeSchematicBitmaps, and is the whole reason that pairing
        // exists: an Image keeps its Source across a detach, so disposing first and clearing second
        // (or not clearing at all) leaves a window in which the renderer can touch a disposed
        // bitmap - an ObjectDisposedException on the render thread, which is fatal. See
        // OnDetachedFromVisualTree's own comment for the tab-switch case that hit this.
        //
        // Only the board pane holds them. The entry list beside it draws text and pills, and the
        // workbook list is text-only, so neither can strand a bitmap reference.
        // ###########################################################################################
        private void ClearBoardPreviewsBeforeDisposingBitmaps()
        {
            this.BoardPreviewPanel?.Children.Clear();

            // The selected schematic is a name, not a control, but it names a preview that no longer
            // exists - clearing it keeps the "keep the selection if it is still shown, else fall
            // back" rule in RefreshBoardPreviews reading against reality on the way back in.
            this.thisSelectedSchematicName = null;
        }

        // ###########################################################################################
        // Releases every decoded schematic bitmap. Called from DetachedFromVisualTree, immediately
        // after ClearBoardPreviewsBeforeDisposingBitmaps has removed everything that references them.
        //
        // Not on a board change or a workbook switch: an editor opened from a pill outlives the
        // refresh that a save triggers, and it renders the bitmap this tab handed it. See the cache's
        // own header.
        // ###########################################################################################
        private void DisposeSchematicBitmaps()
        {
            foreach (var bitmap in this.thisSchematicBitmapsByPath.Values)
            {
                bitmap.Dispose();
            }

            this.thisSchematicBitmapsByPath.Clear();
        }

        // ###########################################################################################
        // Public entry point for Main: rebuilds JUST the board pane for whatever workbook and
        // schematic are currently selected, without touching the workbook list or the top-line.
        //
        // Called from Main.SetComponentHighlightRects, i.e. whenever the component highlight-rect
        // cache is replaced - a board load finishing, or a region switch. The pane's pills go on
        // screen as soon as the board data is assigned, but that cache is populated later, by the
        // board load's fire-and-forget task; a pill clicked in between found no rects for its
        // schematic, so BuildWorklogEntryComponentScope returned null and the editor opened WITHOUT
        // "Mark components in scope" - the same "the two modals are not identical" bug this feature
        // was written to fix, back as an intermittent one. A region switch had the mirror problem:
        // HighlightRectBuilder skips highlights failing IsVisibleByRegion, so the cache legitimately
        // changes shape and the pane has to be rebuilt against it.
        //
        // The list and the top-line are deliberately NOT rebuilt here: neither depends on the
        // highlight cache, and rebuilding them would reset nothing usefully.
        // ###########################################################################################
        public void RefreshBoardPreviewsForCurrentSelection()
        {
            this.StartFreshBoardPass();
        }

        // ###########################################################################################
        // One refresh pass over the board pane AND the summary strip above it, from freshly-read
        // entries.
        //
        // WHY THE THREE THINGS GO TOGETHER: thisEntriesReadThisPass caches an entries.json read for
        // the duration of one pass, so a pass that does not clear it redraws from whatever the
        // previous pass read - after a save, that is the pre-save record. And the summary strip is
        // computed from those same entries, so rebuilding the pane without recomputing it leaves
        // the totals above the pane disagreeing with the pills inside it.
        //
        // Every path that rebuilds the pane WITHOUT coming through RefreshWorkbooks (which does all
        // three itself) must call this rather than RefreshBoardPreviews directly. Three such paths
        // existed and each had got a different subset right: the highlight-cache refresh and
        // SelectSchematic cleared the cache but never touched the summary, and the no-MainWindow
        // save path did neither.
        // ###########################################################################################
        private void StartFreshBoardPass()
        {
            this.thisEntriesReadThisPass.Clear();
            this.RefreshBoardPreviews();
            this.RefreshSummaryForShownWorkbook();
        }

        // The summary strip for whichever workbook the tab is currently showing, or nothing when
        // none is selected - ApplySummaryForWorkbook needs the record, not just the id.
        private void RefreshSummaryForShownWorkbook()
        {
            if (this.thisSelectedWorkbookId <= 0)
                return;

            string boardKey = this.BoardKeyOverrideForTests ?? this.MainWindow?.GetCurrentBoardKey() ?? string.Empty;

            var workbook = WorklogManager.GetWorkbooksForBoard(boardKey)
                .FirstOrDefault(w => w.Id == this.thisSelectedWorkbookId);

            if (workbook != null)
            {
                this.ApplySummaryForWorkbook(workbook);
            }
        }

        // ###########################################################################################
        // Rebuilds the board pane for the currently selected workbook.
        //
        // Groups the workbook's entries by schematic name, keeps only the ones that match a real
        // schematic image on the current board (an entry can reference a schematic that was since
        // renamed or removed from the board data - it is skipped rather than shown with no image),
        // and builds one preview per remaining schematic.
        // ###########################################################################################
        private void RefreshBoardPreviews()
        {
            if (this.BoardPreviewPanel == null)
                return;

            this.BoardPreviewPanel.Children.Clear();

            var boardData = this.CurrentBoardDataForPreviews;

            // GroupBy before ToDictionary, NOT a bare ToDictionary: board Excel files arrive from
            // classic-repair-toolbox.dk independently of app releases, and BoardDataReader.MapSchematics
            // does no dedup, no uniqueness validation and no case normalisation - whatever the
            // Schematics sheet holds arrives verbatim. A bare ToDictionary throws ArgumentException on
            // the first duplicate, and OrdinalIgnoreCase makes that MORE likely than the default
            // comparer would ("Sheet 1" and "sheet 1" collide here while being two distinct schematics
            // everywhere else). Nothing in this tab catches, so it propagated out through
            // RefreshWorkbooks and Main.RefreshWorklogBar and took down board selection entirely.
            // First wins, matching the GroupBy on entries just below, which already tolerated this.
            var schematicsByName = (boardData?.Schematics ?? new List<BoardSchematicEntry>())
                .Where(s => !string.IsNullOrWhiteSpace(s.SchematicName))
                .GroupBy(s => s.SchematicName, StringComparer.OrdinalIgnoreCase)
                .ToDictionary(g => g.Key, g => g.First(), StringComparer.OrdinalIgnoreCase);

            // > 0, not >= 0: workbook ids start at 1 and the rest of the app uses 0 as "no workbook"
            // (Main.RefreshWorklogBar's own "showByDefault ? activeWorkbook.Id : 0"), so admitting 0
            // here would have this half of the feature treat a sentinel as a real workbook.
            // Through the pass cache, so a rebuild that already read this workbook's entries for the
            // search filter does not read them again - see GetEntriesForThisPass.
            var allWorkbookEntries = this.thisSelectedWorkbookId > 0
                ? this.GetEntriesForThisPass(this.thisSelectedWorkbookId)
                : new List<WorklogEntryRecord>();

            // Narrowed to the search's matches (null = no search active), using the set
            // RefreshWorkbooks already computed - so a filtered board pane shows exactly the pills
            // whose entries matched, and the entry list below it agrees, both from one decision.
            var matchedEntryIds = this.MatchedEntryIdsForWorkbook(this.thisSelectedWorkbookId);
            var entries = matchedEntryIds == null
                ? allWorkbookEntries
                : allWorkbookEntries.Where(e => matchedEntryIds.Contains(e.Id)).ToList();

            var entriesBySchematic = entries
                .Where(e => schematicsByName.ContainsKey(e.SchematicName))
                .GroupBy(e => e.SchematicName, StringComparer.OrdinalIgnoreCase)
                .OrderBy(g => g.Key, StringComparer.OrdinalIgnoreCase)
                .ToList();

            // A BOARD switch drops the selection outright, before the "still valid?" test below could
            // accidentally keep it: schematic names are not unique across boards, so a name that
            // happens to exist on both would otherwise carry a selection made against the previous
            // board - see thisSelectedSchematicBoardKey.
            string currentBoardKey = this.BoardKeyOverrideForTests ?? this.MainWindow?.GetCurrentBoardKey() ?? string.Empty;
            if (!string.Equals(currentBoardKey, this.thisSelectedSchematicBoardKey, StringComparison.OrdinalIgnoreCase))
            {
                this.thisSelectedSchematicName = null;
                this.thisSelectedSchematicBoardKey = currentBoardKey;
            }

            // Otherwise it stays selected across a rebuild if it is still one of the schematics
            // actually shown (a save via the pill's editor re-enters this method, and must not
            // silently reset which schematic's entries the right panel is showing); failing that it
            // falls back to the first one - the same "keep if still valid, else fall back" rule
            // RefreshWorkbooks applies to the selected workbook.
            if (entriesBySchematic.Count == 0)
            {
                this.thisSelectedSchematicName = null;
            }
            else if (!entriesBySchematic.Any(g => string.Equals(g.Key, this.thisSelectedSchematicName, StringComparison.OrdinalIgnoreCase)))
            {
                this.thisSelectedSchematicName = entriesBySchematic[0].Key;
            }

            foreach (var group in entriesBySchematic)
            {
                var schematic = schematicsByName[group.Key];
                bool isSelected = string.Equals(group.Key, this.thisSelectedSchematicName, StringComparison.OrdinalIgnoreCase);
                var preview = this.BuildSchematicPreview(schematic, group.ToList(), this.thisSelectedWorkbookId, isSelected);
                if (preview == null)
                    continue;

                this.BoardPreviewPanel.Children.Add(preview);
            }

            this.NoBoardPreviewsText.IsVisible = entriesBySchematic.Count == 0;

            // Three different reasons this pane can be empty, and they must not be described with
            // one another's wording.
            //
            // The first is specific to "Show all workbooks" scope: the list can be showing cards for
            // OTHER boards while this pane can only ever render the CURRENTLY LOADED board's
            // schematics, so a search matching only another board's workbooks leaves a populated
            // list and a non-zero count above a pane with nothing in it and no top-line at all.
            // Without this the user is looking at three result cards and a blank rectangle, with
            // nothing saying the results are simply not on this board.
            if (this.thisSelectedWorkbookId <= 0 && this.thisHasWorkbooksOnOtherBoardsOnly)
            {
                this.NoBoardPreviewsText.Text =
                    "The matching workbooks are on other boards. Click one to switch to its board.";
            }
            else
            {
                // As on the two lists: "none matched" is not the same as "none recorded", and saying
                // "yet" for a search result reads as the entries having gone missing.
                this.NoBoardPreviewsText.Text = matchedEntryIds != null && allWorkbookEntries.Count > 0
                    ? "No worklogs in this workbook match your search."
                    : NoWorklogsForBoardMessage;
            }

            // The entries this pass already read are handed on rather than re-read: GetEntries has
            // no cache (File.ReadAllText + Deserialize + a per-entry normalise loop, every call), and
            // this method and the entry list were reading the same workbook's file twice per rebuild.
            this.RefreshSelectedSchematicEntries(entries);
        }

        // ###########################################################################################
        // Selects a schematic by name, called when the user clicks its preview (anywhere except a
        // pill - see BuildSchematicPreview's own click handler). Restyles the previews' border
        // (exactly one gets the selected colour) and rebuilds the entry list on the right for it.
        //
        // Re-clicking the already-selected schematic is a no-op, matching SelectWorkbook's own rule
        // for the same reason: there is nothing to select that is not already selected.
        // ###########################################################################################
        // Exposed to the test project so WorkbooksBoardPreviewTests can select a schematic without
        // fighting pointer-event routing against a preview Border - the same idea as
        // SelectWorkbookForTests. The running app always reaches this through a preview's own
        // PointerPressed handler in BuildSchematicPreview, never directly.
        internal void SelectSchematicForTests(string schematicName) => this.SelectSchematic(schematicName);

        private void SelectSchematic(string schematicName)
        {
            if (string.Equals(schematicName, this.thisSelectedSchematicName, StringComparison.OrdinalIgnoreCase))
                return;

            this.thisSelectedSchematicName = schematicName;

            // This is its own refresh pass - it does not come through RefreshWorkbooks, which is
            // what normally clears the cache - so clear it here too. Without this a click could be
            // served entries read before an intervening save. The summary is recomputed with it,
            // for the reason StartFreshBoardPass gives.
            this.thisEntriesReadThisPass.Clear();
            this.RefreshSummaryForShownWorkbook();

            // Makes the newly clicked schematic the one waiting on the Schematics tab too - see
            // PropagateSelectedSchematicToSchematicsTab's own header.
            this.PropagateSelectedSchematicToSchematicsTab();

            foreach (var child in this.BoardPreviewPanel.Children)
            {
                if (child is not Border preview || preview.Tag is not string previewSchematicName)
                    continue;

                ApplySchematicPreviewSelectedBorder(preview, string.Equals(previewSchematicName, schematicName, StringComparison.OrdinalIgnoreCase));
            }

            this.RefreshSelectedSchematicEntries();
        }

        // ###########################################################################################
        // Makes thisSelectedSchematicName the one showing on the Schematics tab too - asked for
        // explicitly: selecting a workbook (or a schematic within it) on THIS tab should have the
        // matching schematic image waiting on the Schematics tab, without actually switching there.
        // The user stays on Workbooks; only what the Schematics tab would show if they clicked over
        // to it changes.
        //
        // Goes through TabSchematics.SelectSchematicByName, which sets the thumbnail list's
        // SelectedItem - the exact same path a click on that thumbnail takes, so the full-res image,
        // overlays and "last schematic for this board" all follow exactly as they would from a real
        // click. That method already no-ops when the name is not found or already selected, so no
        // guard is needed here beyond the null checks for a headless test or a board with nothing
        // selected yet.
        //
        // CALL THIS ONLY FROM A USER-INITIATED SELECTION - SelectSchematic (a click on a preview)
        // and SelectWorkbook (a click on a workbook card). It must NOT be called from
        // RefreshBoardPreviews, where it sat at first: that method runs on EVERY refresh pass -
        // Main.RefreshWorklogBar drives it on every entry save, workbook create/delete, board load
        // and search-debounce tick - and it re-derives thisSelectedSchematicName to
        // entriesBySchematic[0] (alphabetically first) whenever the previous choice is not in the
        // rebuilt set. Hooked there, saving a worklog entry FROM the Schematics tab yanked that tab
        // off the schematic the user was working on and onto an unrelated one, discarding its
        // full-resolution bitmap; a refresh arriving mid "Add worklog" also cancelled the
        // area-marking mode outright, since OnSchematicsThumbnailSelectionChanged opens with
        // CancelWorklogEntryMode. Propagation is a response to the user choosing something here,
        // never a side effect of this tab refreshing itself.
        // ###########################################################################################
        private void PropagateSelectedSchematicToSchematicsTab()
        {
            if (this.thisSelectedSchematicName == null)
                return;

            this.MainWindow?.TabSchematicsControl?.SelectSchematicByName(this.thisSelectedSchematicName);
        }

        // ###########################################################################################
        // Rebuilds the entry list on the right (marker 4) for whichever schematic is currently
        // selected in the board pane. Called from RefreshBoardPreviews (a workbook switch, or the
        // selection defaulting to the first schematic) and from SelectSchematic (a click).
        //
        // Takes the workbook's entries when the caller has just read them (RefreshBoardPreviews does,
        // for its own grouping) and reads them itself otherwise - SelectSchematic runs this on its
        // own, with no board-pane rebuild alongside it and so nothing already in hand. GetEntries is
        // uncached and re-parses the whole file per call, so the shared read is worth threading
        // through.
        // ###########################################################################################
        private void RefreshSelectedSchematicEntries(List<WorklogEntryRecord>? workbookEntries = null)
        {
            if (this.SelectedSchematicEntriesPanel == null)
                return;

            this.SelectedSchematicEntriesPanel.Children.Clear();

            if (this.thisSelectedSchematicName == null)
            {
                this.SelectedSchematicEntriesHeaderText.Text = "Select a schematic";
                this.NoSelectedSchematicEntriesText.Text = NoWorklogsForBoardMessage;
                this.NoSelectedSchematicEntriesText.IsVisible = true;
                return;
            }

            var allEntries = workbookEntries
                ?? (this.thisSelectedWorkbookId > 0
                    ? this.GetEntriesForThisPass(this.thisSelectedWorkbookId)
                    : new List<WorklogEntryRecord>());

            // Narrowed to what the search matched, using the SAME set RefreshWorkbooks computed for
            // this workbook rather than re-running the query here - so the list cannot disagree with
            // the workbook card and the board pane about what matched. Null means no search is
            // active and every entry is shown.
            var matchedEntryIds = this.MatchedEntryIdsForWorkbook(this.thisSelectedWorkbookId);

            var entries = allEntries
                .Where(e => string.Equals(e.SchematicName, this.thisSelectedSchematicName, StringComparison.OrdinalIgnoreCase))
                .Where(e => matchedEntryIds == null || matchedEntryIds.Contains(e.Id))
                .OrderBy(e => e.Id)
                .ToList();

            this.SelectedSchematicEntriesHeaderText.Text =
                $"{this.thisSelectedSchematicName} · {WorklogEntryScope.FormatCount(entries.Count, "worklog", "worklogs")}";

            if (entries.Count == 0)
            {
                // A search that hid them all is a different thing from a schematic that never had
                // any, and saying "yet" for it reads as the entries having been lost.
                this.NoSelectedSchematicEntriesText.Text = matchedEntryIds != null
                    ? "No worklogs on this schematic match your search."
                    : "No worklogs for this schematic yet.";
                this.NoSelectedSchematicEntriesText.IsVisible = true;
                return;
            }

            this.NoSelectedSchematicEntriesText.IsVisible = false;

            foreach (var entry in entries)
            {
                this.SelectedSchematicEntriesPanel.Children.Add(this.BuildEntryDetailCard(entry));
            }
        }

        // ###########################################################################################
        // Builds one entry's detail card for the selected-schematic list: ONE 1px-bordered panel
        // holding four stacked rows - "#{N} {Title}", the description, a category chip + status
        // pill, and the stats row - per the layout this list was specifically asked for. Deliberately NOT the
        // anchor tag / timestamp / photo-thumbnail layout the entry list used to show (that markup
        // is gone) - this is a different, narrower view: what an entry IS, not where it sits or
        // what evidence backs it. The full picture is still one click away, via the entry's pill on
        // the board pane above (OnPreviewBadgePointerPressed).
        //
        // The "#N" badge is drawn in the same filled/white-text visual WorklogEntryEditorWindow uses
        // for its own EditorIdBadge - filled is right there, since it names WHICH workbook entry this
        // is rather than a selection state. The category chip and status pill beside it are
        // deliberately the INFORMATIONAL outlined variant instead, from the one shared
        // WorklogInfoPillBuilder - see that class for why they are not drawn here.
        // ###########################################################################################
        // An instance method rather than static: the title and description are drawn through
        // BuildHighlightedTextBlock, which needs THIS tab's current search query to know what to
        // mark. Everything else it builds is still static.
        private Border BuildEntryDetailCard(WorklogEntryRecord entry)
        {
            string title = string.IsNullOrWhiteSpace(entry.Title) ? "(untitled)" : entry.Title;
            Color categoryColor = ResolveWorklogCategoryColor(entry.Category);

            var idBadge = new Border
            {
                Background = new SolidColorBrush(categoryColor),
                CornerRadius = new CornerRadius(12),
                Padding = new Thickness(10, 3),
                VerticalAlignment = VerticalAlignment.Center,
                Child = new TextBlock
                {
                    Text = $"#{entry.Id}",
                    FontSize = 13,
                    FontWeight = FontWeight.Bold,
                    Foreground = Brushes.White
                }
            };

            var titleText = this.BuildHighlightedTextBlock(title, 13, TextWrapping.Wrap, block =>
            {
                block.FontWeight = FontWeight.Bold;
                block.VerticalAlignment = VerticalAlignment.Center;
            });

            var titleRow = new WrapPanel { ItemSpacing = 8, LineSpacing = 4 };
            titleRow.Children.Add(titleText);
            titleRow.Children.Add(idBadge);

            // "Delete worklog" pinned to the card's TOP-RIGHT corner, mirroring where "Delete
            // workbook" sits relative to the workbook it acts on. A Grid rather than another
            // WrapPanel item: the title wraps to as many lines as it needs and the button must stay
            // level with the FIRST of them (VerticalAlignment.Top) rather than drifting down beside
            // a long title - the same reason WorkbookHeaderActionsPanel is top-aligned.
            var titleRowGrid = new Grid
            {
                ColumnDefinitions = new ColumnDefinitions("*,Auto")
            };

            var deleteButton = BuildDeleteWorklogButton();
            deleteButton.Click += (_, _) => this.OnDeleteWorklogClick(entry);

            Grid.SetColumn(titleRow, 0);
            Grid.SetColumn(deleteButton, 1);
            titleRowGrid.Children.Add(titleRow);
            titleRowGrid.Children.Add(deleteButton);

            bool hasDescription = !string.IsNullOrWhiteSpace(entry.Description);
            var descriptionText = this.BuildHighlightedTextBlock(
                hasDescription ? entry.Description : "(no description)",
                11,
                TextWrapping.Wrap,
                block =>
                {
                    if (!hasDescription)
                        block.Foreground = ResolveThemeBrushStatic("Workbooks_Faint_Fg");
                },
                // The description is prose the user typed, so a URL in it is clickable here. The
                // title above deliberately is not - see ApplyHighlightedText's linkify note.
                linkify: hasDescription);

            // Both from the ONE shared informational builder - see WorklogInfoPillBuilder for why
            // these may not be drawn by hand here (they were, in a grey outline that did not match
            // the coloured one every other non-selectable pill in the app uses).
            var categoryChip = WorklogInfoPillBuilder.BuildCategoryChip(entry.Category);
            var statusPill = WorklogInfoPillBuilder.BuildStatePill(entry.State);

            var categoryStatusRow = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 6 };
            categoryStatusRow.Children.Add(categoryChip);
            categoryStatusRow.Children.Add(statusPill);

            var statsRow = BuildEntryStatsRow(entry);

            var stack = new StackPanel { Spacing = 6 };
            stack.Children.Add(titleRowGrid);
            stack.Children.Add(descriptionText);
            stack.Children.Add(categoryStatusRow);
            stack.Children.Add(statsRow);

            var card = new Border
            {
                Background = ResolveThemeBrushStatic("Workbooks_Panel_Bg"),
                BorderBrush = ResolveThemeBrushStatic("Workbooks_RowSeparator"),

                // 2px at rest as well as on hover, matching the schematic previews for the same
                // reason - see ApplySchematicPreviewSelectedBorder. Growing 1px to 2px on hover
                // would reflow the card's contents by a pixel as the pointer crossed it.
                BorderThickness = new Thickness(2),
                CornerRadius = new CornerRadius(4),
                Padding = new Thickness(10),

                // The whole card opens this entry's editor, exactly as its pill on the board pane
                // does - see OpenEntryEditor, which both go through. The Hand cursor is the only
                // thing that says so, since the card carries no button of its own; it is the same
                // shared instance the previews and pills use (see HandCursor).
                Cursor = HandCursor,
                Child = stack
            };

            // ###########################################################################################
            // HOVER shows the same IndianRed accent a SELECTED schematic preview uses, so the whole
            // tab speaks one colour language: red outline = "this is the one you are acting on".
            //
            // Hover rather than selection because this list HAS no selection - a card is a button,
            // not a choice that persists after the click. Without it the only affordance was the
            // Hand cursor, which does not say WHICH card a click will land on when several are
            // stacked; the schematic previews beside them already made that clear with an outline.
            //
            // PointerEntered/Exited rather than a Style with a :pointerover selector: these cards
            // are built in code (their brushes need the two-step theme lookup - see the class
            // header), so there is no template for a selector to attach to.
            // ###########################################################################################
            card.PointerEntered += (_, _) => ApplyEntryCardHoverBorder(card, isHovered: true);
            card.PointerExited += (_, _) => ApplyEntryCardHoverBorder(card, isHovered: false);

            card.PointerPressed += (_, e) => this.OnEntryDetailCardPointerPressed(entry, e);

            return card;
        }

        // ###########################################################################################
        // A click anywhere on an entry's detail card opens that entry in the full editor - the same
        // modal, through the same OpenEntryEditor, that the entry's pill on the board pane opens.
        //
        // The schematic bitmap is looked up by the entry's OWN schematic name rather than taken
        // from the selected preview: the two agree today (the list only ever shows the selected
        // schematic's entries), but resolving it from the entry means a future list showing more
        // than one schematic's entries cannot hand the editor the wrong board image. A missing or
        // unreadable image yields null, which the editor renders as no location preview rather
        // than refusing to open - an entry's text is still worth reading without its picture.
        // ###########################################################################################
        private void OnEntryDetailCardPointerPressed(WorklogEntryRecord entry, PointerPressedEventArgs e)
        {
            if (!e.GetCurrentPoint(this).Properties.IsLeftButtonPressed)
                return;

            e.Handled = true;

            this.OpenEntryEditor(this.thisSelectedWorkbookId, entry.Id, this.ResolveSchematicBitmapForEntry(entry));
        }

        // ###########################################################################################
        // The "Delete worklog" button in an entry detail card's top-right corner.
        //
        // Deliberately the SAME destructive styling "Delete workbook" carries in the header above
        // (the Button_Cancel_* brushes), at the same FontSize 11 - the two are the same kind of
        // permanent delete, one level apart, and a differently-coloured one here would read as a
        // different kind of action. No fixed Width, unlike the header's four buttons: those share
        // one because they line up in a grid of two rows, whereas this is a lone button and a fixed
        // width would either clip its label or leave it padded out at random.
        //
        // Its OWN Cursor is the default arrow rather than the card's Hand: the card behind it is
        // clickable as a whole (it opens the editor), and inheriting the Hand would say this button
        // does the same benign thing the rest of the card does. A shared static instance, like
        // HandCursor beside it - a Cursor is a disposable native handle, and one per card per
        // rebuild is a handle leaked on every board change and every entry save.
        // ###########################################################################################
        private static readonly Cursor ArrowCursor = new(StandardCursorType.Arrow);

        private static Button BuildDeleteWorklogButton() => new()
        {
            Content = "Delete worklog",
            FontSize = 11,
            Padding = new Thickness(8, 4),
            VerticalAlignment = VerticalAlignment.Top,
            HorizontalAlignment = HorizontalAlignment.Right,
            Foreground = ResolveThemeBrushStatic("Button_Cancel_Fg"),
            Background = ResolveThemeBrushStatic("Button_Cancel_Bg"),
            BorderBrush = ResolveThemeBrushStatic("Button_Cancel_Border"),
            Cursor = ArrowCursor
        };

        // ###########################################################################################
        // Deletes one worklog from the selected workbook, after the user confirms in
        // DeleteWorklogWindow - the same confirm-then-act shape "Delete workbook" uses, and for the
        // same reason: WorklogManager.DeleteEntry removes the entry's row AND its whole attachment
        // folder, photos and files included, with nothing to undo it.
        //
        // A Button rather than a bare click region, so the press never reaches the card's own
        // PointerPressed underneath it (a Button handles the press itself) - a single click must
        // not both open the editor and raise a delete confirmation over it.
        //
        // The refresh goes through Main.RefreshWorklogBar, the one funnel every worklog change
        // passes through, so the schematic overlay's rectangles and thumbnail pills lose the deleted
        // entry too rather than drawing a worklog that no longer exists. With no MainWindow
        // (headless tests) there is nothing to funnel through, so a full fresh pass rebuilds this
        // tab directly - a bare RefreshBoardPreviews would redraw from this pass's entry read cache,
        // which still holds the entry that was just deleted.
        // ###########################################################################################
        private async void OnDeleteWorklogClick(WorklogEntryRecord entry)
        {
            if (this.thisSelectedWorkbookId <= 0)
                return;

            if (TopLevel.GetTopLevel(this) is not Window ownerWindow)
                return;

            var dialog = new DeleteWorklogWindow();
            dialog.Initialize(entry);

            bool? confirmed = await dialog.ShowDialog<bool?>(ownerWindow);
            if (confirmed != true)
                return;

            if (!WorklogManager.DeleteEntry(this.thisSelectedWorkbookId, entry.Id))
            {
                // The same treatment a failed workbook delete gets: the user confirmed a
                // destructive action, and a list that simply does not change reads as "the click
                // did not register" and invites them to try again.
                await ShowWorkbookActionFailedAsync(
                    ownerWindow,
                    "Delete worklog",
                    $"Could not delete worklog #{entry.Id} - see the log for details.\n\n" +
                    "It may be open in another program, for example a photo or file from the worklog.");
                return;
            }

            if (this.MainWindow != null)
            {
                this.MainWindow.RefreshWorklogBar();
            }
            else
            {
                this.StartFreshBoardPass();
            }
        }

        // ###########################################################################################
        // The shared decoded bitmap for the schematic an entry sits on, or null when the board data
        // does not name that schematic or its image file is missing from the data root.
        //
        // Same lookup BuildSchematicPreview does for the pane itself, and it goes through the same
        // GetOrDecodeSchematicBitmap cache - so opening an entry from its card hands the editor the
        // very bitmap instance the pane is already drawing, not a second decode of the same file.
        // ###########################################################################################
        private Bitmap? ResolveSchematicBitmapForEntry(WorklogEntryRecord entry)
        {
            if (string.IsNullOrWhiteSpace(entry.SchematicName))
                return null;

            var schematic = this.CurrentBoardDataForPreviews?.Schematics
                ?.FirstOrDefault(s => string.Equals(s.SchematicName, entry.SchematicName, StringComparison.OrdinalIgnoreCase));

            if (string.IsNullOrWhiteSpace(schematic?.SchematicImageFile))
                return null;

            return this.GetOrDecodeSchematicBitmap(Path.Combine(
                DataManager.DataRoot,
                schematic.SchematicImageFile.Replace('/', Path.DirectorySeparatorChar)));
        }

        // ###########################################################################################
        // The fourth row on an entry's detail card: total hours spent and total cost (both summed
        // across the entry's Work done rows, the same sums WorklogEntryEditorWindow's own
        // RefreshWorkDoneRows shows beside that section's heading - "{hours} h" / the bare cost
        // number, no currency symbol, matching SummaryText's own "{HoursSpent:0.##} h · {Cost:0.##}"
        // formatting exactly), then how many comments, links, photos and files the entry carries -
        // one number each, since an entry can have any number of each and this card is meant to say
        // how much is behind the pill without opening it. A WrapPanel, matching every other
        // multi-item row on this card and in the full editor's own list headers: six items can
        // outrun the entry list's narrower widths, and wrapping is preferable to clipping or a
        // horizontal scrollbar inside a vertically-scrolling list.
        // ###########################################################################################
        private static WrapPanel BuildEntryStatsRow(WorklogEntryRecord entry)
        {
            var (totalHours, totalCost) = WorklogEntryScope.GetWorkDoneTotals(entry);

            var row = new WrapPanel { ItemSpacing = 10, LineSpacing = 4 };
            row.Children.Add(BuildEntryStatText($"{totalHours.ToString("0.##", CultureInfo.InvariantCulture)} h"));
            row.Children.Add(BuildEntryStatText(totalCost.ToString("0.##", CultureInfo.InvariantCulture)));
            row.Children.Add(BuildEntryStatText(WorklogEntryScope.FormatCount(entry.Comments.Count, "comment", "comments")));
            row.Children.Add(BuildEntryStatText(WorklogEntryScope.FormatCount(entry.Links.Count, "link", "links")));
            row.Children.Add(BuildEntryStatText(WorklogEntryScope.FormatCount(entry.Photos.Count, "photo", "photos")));
            row.Children.Add(BuildEntryStatText(WorklogEntryScope.FormatCount(entry.Files.Count, "file", "files")));
            return row;
        }

        private static TextBlock BuildEntryStatText(string text) => new()
        {
            Text = text,
            FontSize = 10,
            Foreground = ResolveThemeBrushStatic("Workbooks_Muted_Fg"),
            VerticalAlignment = VerticalAlignment.Center
        };


        // ###########################################################################################
        // Builds one schematic's preview: its image (natural aspect ratio, capped in height so a
        // workbook with several schematics does not force endless scrolling for one board-sized
        // image) with the entries overlay and badges on top, and a caption naming the schematic.
        //
        // Returns null when the image file cannot be loaded - happens if the data root is out of
        // sync with a board's Excel file, the same failure mode DataManager already tolerates
        // elsewhere - so one bad path drops that schematic rather than the whole pane.
        //
        // Clicking anywhere on the returned Border except a pill selects this schematic (see the
        // PointerPressed handler near the end and OnPreviewBadgePointerPressed's own e.Handled,
        // which stops a pill click from also reaching this one). The Border is tagged with the
        // schematic's name so SelectSchematic can find it again to restyle its border without a
        // second dictionary mapping controls back to names.
        //
        // An instance method, not static: each badge's click handler needs to reach
        // OnPreviewBadgePointerPressed on THIS tab.
        // ###########################################################################################
        private Border? BuildSchematicPreview(BoardSchematicEntry schematic, List<WorklogEntryRecord> entries, int workbookId, bool isSelected)
        {
            if (string.IsNullOrWhiteSpace(schematic.SchematicImageFile))
                return null;

            string fullPath = Path.Combine(
                DataManager.DataRoot,
                schematic.SchematicImageFile.Replace('/', Path.DirectorySeparatorChar));

            var bitmap = this.GetOrDecodeSchematicBitmap(fullPath);
            if (bitmap == null)
                return null;

            var image = new Image
            {
                Source = bitmap,
                Stretch = Stretch.Uniform,
                HorizontalAlignment = HorizontalAlignment.Left,
                VerticalAlignment = VerticalAlignment.Top,
                MaxHeight = 420
            };

            var overlayEntries = new List<WorklogEntriesOverlay.Entry>();

            // Anchored badges (Tag = the entry's pixel rect) sit on the marked area, positioned in
            // PositionPreviewOverlayAndBadges once the image has its real size. Parked badges (Tag =
            // null) have no area to anchor to - they stack in the image's own top-right corner
            // instead, exactly what TabSchematics.Worklog.cs's LayOutWorklogParkedBadges does for
            // the main Schematics view. Same Left/Top alignment as the overlay below, for the same
            // reason: both sets of positions are in the image's own content coordinates, which only
            // line up with the canvas's (0,0) if the canvas starts where the image does.
            // Hit-test visible, unlike the overlay and border outline below - the pills are the
            // only clickable thing on this preview (see OnPreviewBadgePointerPressed), matching the
            // real Schematics tab's own split between SchematicsWorklogEntriesOverlay
            // (IsHitTestVisible="False", the bounds) and SchematicsWorklogEntriesBadgeCanvas
            // (clickable, the pills).
            var badgeCanvas = new Canvas
            {
                HorizontalAlignment = HorizontalAlignment.Left,
                VerticalAlignment = VerticalAlignment.Top
            };

            foreach (var entry in entries)
            {
                Color color = ResolveWorklogCategoryColor(entry.Category);
                var badge = BuildPreviewBadge(entry, color);
                badgeCanvas.Children.Add(badge);

                // Captured entry id, not read back off Tag: Tag is already spoken for below (it
                // carries the anchor Rect for PositionPreviewOverlayAndBadges, or stays null for a
                // parked badge), and overloading it a second way for the click target would make
                // one of the two meanings silently clobber the other. workbookId is the same for
                // every badge in this preview, so it is captured once rather than per-entry.
                //
                // Not unsubscribed on rebuild, and does not need to be: this closure and the
                // SizeChanged one below reference only controls INSIDE this preview's own subtree
                // (plus `this`, which points outward and so retains nothing), so once
                // BoardPreviewPanel.Children.Clear() drops the preview the whole cycle is
                // unreachable together and collects. What did NOT collect was the decoded bitmap's
                // unmanaged Skia surface - see thisSchematicBitmapsByPath, which is what actually
                // fixed the leak.
                int entryId = entry.Id;
                badge.PointerPressed += (_, e) => this.OnPreviewBadgePointerPressed(workbookId, entryId, bitmap, e);

                // "Show marked area" ticked: a coloured bounds rectangle on the overlay, badge
                // anchored to it. Unticked: no rectangle at all, badge parked in the corner instead
                // (Tag stays null) - matching TabSchematics.Worklog.cs's own ShowMarkedArea branch
                // exactly. Getting this backwards was the original bug report: every badge was
                // anchored to its marker regardless of ShowMarkedArea, so an entry meant to show
                // only a parked pill still appeared pinned to the spot it was drawn at.
                if (entry.ShowMarkedArea)
                {
                    var pixelRect = new Rect(entry.AreaX, entry.AreaY, entry.AreaWidth, entry.AreaHeight);
                    overlayEntries.Add(new WorklogEntriesOverlay.Entry(pixelRect, color, entry.Id));
                    badge.Tag = pixelRect;
                }
            }

            // Left/Top-aligned to match the Image exactly - WorklogEntriesOverlay.Render assumes
            // its OWN Bounds starts at the image's content origin (see GetImageContentRect's own
            // comment: SchematicsImage is Left/Top-aligned, so bitmap content starts at (0,0)).
            // Stretch (the Grid default) would give it the whole cell instead, which is a
            // different rect than the image occupies whenever the image's aspect ratio does not
            // match the cell's - i.e. almost always.
            var overlay = new WorklogEntriesOverlay
            {
                IsHitTestVisible = false,
                HorizontalAlignment = HorizontalAlignment.Left,
                VerticalAlignment = VerticalAlignment.Top,
                BitmapPixelSize = bitmap.PixelSize,
                ViewMatrix = Matrix.Identity,
                Entries = overlayEntries
            };

            // The 1px border marking the image's boundary, so it reads clearly against the pane's
            // own background - Image has no border property of its own, and a schematic with a lot
            // of white/transparent area was otherwise hard to tell apart from the pane behind it.
            // A separate sibling Grid child, NOT a wrapper around the image: wrapping it would
            // inset the image by the border's thickness, throwing off every pixel-rect calculation
            // below that assumes image.Bounds starts at the image control's own (0,0) - exactly the
            // assumption GetImageContentRect documents for the real Schematics tab's own image.
            // Sized and positioned to match the image exactly in PositionPreviewOverlayAndBadges,
            // same as the overlay and the badge canvas.
            var imageBorderOutline = new Border
            {
                BorderBrush = Brushes.Black,
                BorderThickness = new Thickness(1),
                IsHitTestVisible = false,
                HorizontalAlignment = HorizontalAlignment.Left,
                VerticalAlignment = VerticalAlignment.Top
            };

            // The overlay, the border outline and the badge canvas all need the image's ARRANGED
            // size to place themselves correctly (GetImageContentRect depends on it), which is not
            // known until this preview has been through a layout pass. SizeChanged catches every
            // pass, not just the first, so a window resize that changes how large this
            // Uniform-stretched image renders keeps all three in step with it.
            image.SizeChanged += (_, _) => PositionPreviewOverlayAndBadges(image, overlay, imageBorderOutline, badgeCanvas, bitmap.PixelSize);

            var imageLayer = new Grid();
            imageLayer.Children.Add(image);
            imageLayer.Children.Add(imageBorderOutline);
            imageLayer.Children.Add(overlay);
            imageLayer.Children.Add(badgeCanvas);

            string displayName = string.IsNullOrWhiteSpace(schematic.CadName) ? schematic.SchematicName : schematic.CadName;
            var caption = new TextBlock
            {
                Text = displayName,
                FontSize = 11,
                FontWeight = FontWeight.Bold
            };

            var body = new StackPanel { Spacing = 6 };
            body.Children.Add(caption);
            body.Children.Add(imageLayer);

            var preview = new Border
            {
                Background = ThemeResources.ResolveBrush("Form_Bg", Brushes.Transparent),
                BorderThickness = new Thickness(2),
                CornerRadius = new CornerRadius(4),
                Padding = new Thickness(10),
                Cursor = HandCursor,
                Tag = schematic.SchematicName,
                Child = body
            };

            ApplySchematicPreviewSelectedBorder(preview, isSelected);

            string schematicName = schematic.SchematicName;
            preview.PointerPressed += (_, e) =>
            {
                if (e.GetCurrentPoint(preview).Properties.IsLeftButtonPressed)
                    this.SelectSchematic(schematicName);
            };

            return preview;
        }

        // ###########################################################################################
        // Colours an entry detail card's border for the hovered state - the SAME
        // Main_TabUnderline_Selected accent a selected schematic preview takes, deliberately: both
        // mean "the thing you are about to act on", and two different reds for that would read as
        // two different meanings.
        // ###########################################################################################
        private static void ApplyEntryCardHoverBorder(Border card, bool isHovered)
        {
            card.BorderBrush = isHovered
                ? ResolveThemeBrushStatic("Main_TabUnderline_Selected")
                : ResolveThemeBrushStatic("Workbooks_RowSeparator");
        }

        // ###########################################################################################
        // Colours a schematic preview's own border to show whether it is the selected one - the
        // same IndianRed accent BuildWorkbookCard's selected-card left edge uses, so "selected"
        // reads as one consistent colour language across this tab. A plain 1px neutral outline
        // (Workbooks_RowSeparator, matching every other panel border here) when not selected.
        //
        // 2px thick either way, not 1px growing to 2px on selection: changing thickness on select
        // would shift the whole preview by a pixel and visibly nudge the image and its pills,
        // which a colour change alone does not.
        // ###########################################################################################
        private static void ApplySchematicPreviewSelectedBorder(Border preview, bool isSelected)
        {
            preview.BorderBrush = isSelected
                ? ResolveThemeBrushStatic("Main_TabUnderline_Selected")
                : ResolveThemeBrushStatic("Workbooks_RowSeparator");
        }

        // ###########################################################################################
        // Positions the overlay and every badge for one schematic preview against the image's
        // current arranged size. Re-run on every SizeChanged, not only the first layout pass, so a
        // window resize keeps the bounds and pill positions in step with the Uniform-stretched
        // image rather than freezing them at whatever size the image happened to have on first
        // measure.
        // ###########################################################################################
        // A parked badge's stack sits inset from the image's own edges, matching
        // TabSchematics.Worklog.cs's WorklogParkedBadgeMargin/WorklogParkedBadgeSpacing constants -
        // the same visual gap the real Schematics tab uses for its own parked pills.
        private const double ParkedBadgeMargin = 10.0;

        private const double ParkedBadgeSpacing = 6.0;

        private static void PositionPreviewOverlayAndBadges(Image image, WorklogEntriesOverlay overlay, Border imageBorderOutline, Canvas badgeCanvas, PixelSize bitmapPixelSize)
        {
            var imageBounds = image.Bounds.Size;
            overlay.Width = imageBounds.Width;
            overlay.Height = imageBounds.Height;
            overlay.InvalidateVisual();

            imageBorderOutline.Width = imageBounds.Width;
            imageBorderOutline.Height = imageBounds.Height;

            if (imageBounds.Width <= 0 || imageBounds.Height <= 0)
                return;

            var contentRect = RectGeometry.GetImageContentRect(imageBounds, bitmapPixelSize);
            if (contentRect.Width <= 0 || contentRect.Height <= 0)
                return;

            // Measure first, then hand the whole set to WorklogBadgeLayout, which decides where each
            // one goes. The anchored-vs-parked rule lives there rather than here because getting it
            // backwards was a reported bug and it is worth pinning with a fast unit test rather than
            // only through a full headless layout pass - see that class.
            var badges = new List<Border>();
            var requests = new List<WorklogBadgeLayout.BadgePlacementRequest>();

            foreach (var child in badgeCanvas.Children)
            {
                if (child is not Border badge)
                    continue;

                badge.Measure(new Size(double.PositiveInfinity, double.PositiveInfinity));

                // Tag carries the anchor rect for a "show marked area" entry, and stays null for one
                // without - see BuildSchematicPreview, which sets it.
                badges.Add(badge);
                requests.Add(new WorklogBadgeLayout.BadgePlacementRequest(
                    badge.DesiredSize,
                    badge.Tag is Rect pixelRect ? pixelRect : null));
            }

            var positions = WorklogBadgeLayout.ArrangeBadges(
                requests, contentRect, bitmapPixelSize, ParkedBadgeMargin, ParkedBadgeSpacing);

            for (int i = 0; i < badges.Count && i < positions.Count; i++)
            {
                Canvas.SetLeft(badges[i], positions[i].X);
                Canvas.SetTop(badges[i], positions[i].Y);
            }
        }

        // ###########################################################################################
        // Resolves the "Mark components in scope"/"Mark components completed" checklist for an entry.
        //
        // The computation is WorklogEntryScope.BuildComponentsInScope, the SAME call
        // TabSchematics.Worklog.cs makes - not a near-identical copy of it, which is what the two
        // were. This tab and that one are required to open the SAME modal, and a fix applied to one
        // copy and missed on the other is exactly the divergence already reported once.
        //
        // What differs is only where the two inputs come from. TabSchematics reads its own
        // highlightRectsBySchematicAndLabel field; this tab borrows that same cache off
        // MainWindow.TabSchematicsControl (via HighlightRectsBySchematicAndLabelForPreviews, the same
        // override-then-real-MainWindow seam CurrentBoardDataForPreviews uses), which is really
        // Main's - it writes it from four places and TabSchematics never assigns it.
        //
        // The cache being CURRENT here is not an accident of timing: Main routes every write to it
        // through SetComponentHighlightRects, which refreshes this pane as it writes, so a pill only
        // exists on screen against the cache it will be looked up in. Before that, the pane's pills
        // appeared while the board load's fire-and-forget task had yet to populate it, and a click in
        // that window silently dropped the checklist.
        // ###########################################################################################
        private List<(string BoardLabel, string DisplayName)>? BuildWorklogEntryComponentScope(WorklogEntryRecord entry) =>
            WorklogEntryScope.BuildComponentsInScope(
                this.CurrentBoardDataForPreviews,
                this.HighlightRectsBySchematicAndLabelForPreviews,
                entry);

        // Exposed to the test project so a component-scope test can call this directly without
        // driving a real pointer press through to WorklogEntryEditorWindow.ShowDialog, which blocks
        // headlessly with nothing to dismiss it (see A_pill_has_a_hand_cursor_and_the_canvas_it_sits_on_is_clickable's
        // own comment) - the same idea as SelectWorkbookForTests and SelectSchematicForTests.
        internal List<(string BoardLabel, string DisplayName)>? BuildWorklogEntryComponentScopeForTests(WorklogEntryRecord entry) =>
            this.BuildWorklogEntryComponentScope(entry);

        // ###########################################################################################
        // Opens a pill's worklog entry in the same editor window the Schematics tab's own "Show
        // worklogs" badges open - see TabSchematics.Worklog.cs.OnWorklogEntryPillPointerPressed,
        // which this mirrors, INCLUDING the component-scope checklist (BuildWorklogEntryComponentScope
        // above) - the modal opened from here is now the exact same modal the Schematics tab opens,
        // not a lookalike missing that section, which was a reported gap.
        //
        // The WHOLE body is inside the try, not just the post-await refresh. This is an async void
        // handler, so an exception anywhere in it - including before the first await - is rethrown on
        // the sync context with no caller to catch it and reaches App's global handler as a
        // PROCESS-FATAL crash, losing unsaved work in every other tab from one click on a pill. The
        // prologue is not exception-free: GetEntries reads and deserializes entries.json, Initialize
        // loads XAML and decodes photo thumbnails off disk, and ShowDialog itself can throw
        // synchronously. entries.json locked by AV, cloud sync or a second CRT instance, or one
        // corrupt photo thumbnail, is enough.
        //
        // thisIsOpeningEntryEditor guards re-entrancy: e.Handled stops this press reaching the
        // preview's own handler, but does nothing about a SECOND press while the first dialog is
        // still being constructed - ShowDialog is awaited, not blocking, so the dispatcher keeps
        // pumping input. A double-click, or pill #1 then pill #2 in quick succession, opened two
        // editors over the same entries.json, and UpdateEntry rewrites that file wholesale, so the
        // second save silently discarded the first. TabSchematics is partly shielded from this by
        // its own thisIsWorklogEntryMode flag; this pane had no equivalent.
        // ###########################################################################################
        // True while a pill's editor is being opened or is on screen. See
        // OnPreviewBadgePointerPressed for what it guards against.
        private bool thisIsOpeningEntryEditor;

        private void OnPreviewBadgePointerPressed(int workbookId, int entryId, Bitmap schematicBitmap, PointerPressedEventArgs e)
        {
            if (!e.GetCurrentPoint(this).Properties.IsLeftButtonPressed)
                return;

            // Without this, the press bubbles up to the preview's own click handler and selects
            // the schematic underneath the pill in addition to opening the entry's editor - both
            // firing off one click is not what either affordance is meant to do.
            e.Handled = true;

            this.OpenEntryEditor(workbookId, entryId, schematicBitmap);
        }

        // ###########################################################################################
        // Opens the full worklog entry editor for one entry, and refreshes afterwards if it saved.
        //
        // TWO callers, and they must stay one implementation: a pill on the board pane
        // (OnPreviewBadgePointerPressed) and an entry's detail card in the right-hand list
        // (OnEntryDetailCardPointerPressed). Clicking either was asked for explicitly to do the
        // same thing - the card is simply the same entry rendered larger, so a click on it landing
        // anywhere other than this exact modal would be the "the two modals are not identical"
        // complaint over again, one level out.
        //
        // schematicBitmap is the entry's own schematic image, which the editor draws its location
        // preview from. It belongs to this tab's shared cache and must NOT be disposed here - see
        // thisSchematicBitmapsByPath. Null is tolerated by the editor (no preview drawn), which is
        // what a card whose schematic image is missing from the data root passes.
        // ###########################################################################################
        private async void OpenEntryEditor(int workbookId, int entryId, Bitmap? schematicBitmap)
        {
            if (this.thisIsOpeningEntryEditor)
                return;

            this.thisIsOpeningEntryEditor = true;

            try
            {
                var entry = WorklogManager.GetEntries(workbookId).FirstOrDefault(x => x.Id == entryId);
                if (entry == null)
                    return;

                if (TopLevel.GetTopLevel(this) is not Window ownerWindow)
                    return;

                var editor = new WorklogEntryEditorWindow();
                editor.Initialize(workbookId, entry, schematicBitmap);

                var componentsInScope = this.BuildWorklogEntryComponentScope(entry);
                if (componentsInScope != null)
                {
                    editor.InitializeComponentScope(componentsInScope);
                }

                await editor.ShowDialog(ownerWindow);

                if (!editor.WasSaved)
                    return;

                // ONE refresh, through Main: RefreshWorklogBar calls RefreshWorkbooks, which ends in
                // RefreshBoardPreviews - so a direct RefreshBoardPreviews() here as well simply
                // rebuilt the pane twice per save. With no MainWindow (headless tests) there is
                // nothing to funnel through, so the pane is rebuilt directly instead.
                if (this.MainWindow != null)
                {
                    this.MainWindow.RefreshWorklogBar();
                }
                else
                {
                    // A full pass, not a bare RefreshBoardPreviews: the entry that was just saved
                    // is still in this pass's read cache, so the pane would redraw the PRE-save
                    // record - wrong category, wrong state, and stale totals in the strip above it.
                    this.StartFreshBoardPass();
                }
            }
            catch (Exception ex)
            {
                Logger.Warning($"Workbooks tab: failed to open or refresh after editing entry [#{entryId}]: [{ex.Message}]");
            }
            finally
            {
                this.thisIsOpeningEntryEditor = false;
            }
        }

        // ###########################################################################################
        // One "#N" badge + state pill, from the SAME WorklogBadgeBuilder the Schematics tab's own
        // "Show worklogs" view uses - so an entry looks identical wherever it is drawn, rather than
        // by two copies of the same forty-five lines each asserting that it must.
        //
        // No ScaleTransform here: unlike the Schematics tab's badges, this pane never zooms, so the
        // badge is drawn at its natural size with no inverse-scale compensation. The click handler
        // that opens the entry's editor is wired up by the caller (BuildSchematicPreview).
        // ###########################################################################################
        private static Border BuildPreviewBadge(WorklogEntryRecord entry, Color categoryColor) =>
            WorklogBadgeBuilder.Build(entry, categoryColor, ResolveWorklogStateColor(entry.State));

        // ###########################################################################################
        // A saved entry's category and state colours.
        //
        // Both DELEGATE to WorklogInfoPillBuilder rather than resolving the theme keys again. Each
        // used to be its own copy of the same two-line lookup, which is exactly the duplication
        // that builder was introduced to end: the pills on this pane and the badges beside them
        // read their colours from different code, so the two could drift while every comment
        // claimed they matched.
        //
        // Kept as named methods rather than inlining the calls - they are used from several places
        // here, and the names say which of the two colours a call site means.
        // ###########################################################################################
        private static Color ResolveWorklogCategoryColor(string category) =>
            WorklogInfoPillBuilder.ResolveCategoryColor(category);

        private static Color ResolveWorklogStateColor(string state) =>
            WorklogInfoPillBuilder.ResolveStateColor(state);

        // See ThemeResources for why a ThemeDictionaries-scoped key needs the two-step
        // Application.Current + ActualThemeVariant lookup. Falls back to IndianRed, matching every
        // other worklog colour resolver in this app when a key is somehow missing.
        private static IBrush ResolveThemeBrushStatic(string resourceKey) =>
            ThemeResources.ResolveBrush(resourceKey);
    }
}
