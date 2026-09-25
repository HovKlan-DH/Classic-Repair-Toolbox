using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Shapes;
using Avalonia.Threading;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Media;
using Avalonia.Media.Imaging;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Handlers.DataHandling;
using Handlers.Theming;
using Handlers.Geometry;

namespace CRT
{
    // ###########################################################################################
    // The full worklog entry editor - the ONE place a worklog entry is written, whether it is being
    // created or edited. It edits the entry's title/description/category/state, its Links/Comments/
    // WorkDoneItems/Photos/Files sub-lists and its component checklists, and shows a read-only
    // preview of where the entry's marked area sits on its schematic.
    //
    // Two ways in, both of which must be called before ShowDialog:
    //   Initialize            - a SAVED entry, from clicking its pill on the Schematics tab (see
    //                           TabSchematics.Worklog.cs's OnWorklogEntryPillPointerPressed) or on
    //                           the Workbooks tab's board pane.
    //   InitializeForNewEntry - a NEW entry, from "Add worklog" after an area is drawn. It used to
    //                           open a small "New fault" quick card asking for a subset of these
    //                           same fields; that card is gone and this window is the whole step.
    //
    // Works on a private working copy of the WorklogEntryRecord (thisEntry) built from the caller's
    // record. For a saved entry, sub-list changes write through immediately (PersistEntrySilently)
    // and Save commits the direct fields via WorklogManager.UpdateEntry. For a NEW one nothing
    // reaches disk until Save, which writes the whole record through WorklogManager.AddEntryRecord -
    // see thisIsDraftEntry. Cancel/closing discards the working copy either way.
    //
    // The Photos and Files lists are ONE implementation parameterised by an AttachmentSection -
    // see WorklogEntryEditorWindow.Attachments.cs for which axes they differ on and why the ordering
    // rules inside those paths are not safe to hold two copies of.
    // ###########################################################################################
    public partial class WorklogEntryEditorWindow : Window
    {
        // ###########################################################################################
        // WorklogEntryEditorWindow is split across partial-class files by area; this file owns
        // construction, window placement, Initialize/InitializeForNewEntry, key handling, dirty
        // tracking and the direct-field <-> record sync. The others:
        //
        //   WorklogEntryEditorWindow.Components.cs  - component scope population/completion tracking
        //   WorklogEntryEditorWindow.ListSections.cs- shared collapsible list-section infrastructure
        //   WorklogEntryEditorWindow.Visuals.cs     - category chip / state pill visuals, location preview
        //   WorklogEntryEditorWindow.SubLists.cs    - Links/Comments/WorkDoneItems rows (add/edit/delete/sort)
        //   WorklogEntryEditorWindow.Attachments.cs - Photos/Files rows, add/edit/delete/drag-reorder
        //   WorklogEntryEditorWindow.SaveCancel.cs  - discard/cancel/save flow
        //   WorklogEntryEditorWindow.Models.cs      - the top-level row/record types the .axaml binds to
        // ###########################################################################################

        private int thisWorkbookId;
        private WorklogEntryRecord thisEntry = new();
        private Bitmap? thisSchematicBitmap;

        private string thisSelectedCategory = "Note";
        private string thisSelectedState = "Open";

        private readonly ObservableCollection<WorklogLinkRow> thisLinkRows = new();
        private readonly ObservableCollection<WorklogCommentRow> thisCommentRows = new();
        private readonly ObservableCollection<WorklogWorkDoneRow> thisWorkDoneRows = new();
        private readonly ObservableCollection<WorklogAttachmentRow> thisPhotoRows = new();
        private readonly ObservableCollection<WorklogAttachmentRow> thisFileRows = new();

        // ###########################################################################################
        // True while this window is editing an entry that does NOT yet exist on disk - the "Add
        // worklog" flow, which draws an area on the schematic and comes straight here (there is no
        // longer a small quick card in between; see InitializeForNewEntry).
        //
        // A draft is held ENTIRELY in memory until Save. That is the one real behavioural difference
        // from editing a saved entry, and it changes what the instant-save path does: for a saved
        // entry every sub-list change (a comment, a photo, a reorder) writes through to disk at once
        // via PersistEntrySilently, so Cancel cannot take it back; for a draft nothing is written at
        // all, so Cancel discards the whole thing and leaves no half-made entry behind.
        //
        // Attachment BYTES are the exception and are written immediately either way - a photo has to
        // be copied somewhere before it can be shown - into the folder for the id reserved below.
        // WorklogManager.AddEntryRecord moves that folder if the entry ends up with a different id.
        // Cancelling a draft therefore leaves those bytes in a folder no entry names; they are
        // cleaned up on Cancel, see DiscardDraftAttachments.
        // ###########################################################################################
        private bool thisIsDraftEntry;

        // The id a draft's attachment folder is named after, reserved from WorklogManager.
        // PeekNextEntryId when the draft is created. Only meaningful while thisIsDraftEntry.
        private int thisDraftReservedEntryId;

        // Set once AddEntryRecord has actually written this draft, and never cleared. It exists
        // solely so the Closing handler's recursive attachment delete has a second, independent
        // reason not to run - see that handler for why one flag is not enough for an operation that
        // destroys files.
        private bool thisWasSuccessfullySaved;

        // Set once the "discard unsaved work?" question has been answered with yes, so the Closing
        // handler does not ask it a second time on the way out. Also set by BOTH of Save's success
        // paths, which have no work to lose by definition.
        //
        // Save setting it is not redundant belt-and-braces that could be dropped: this is the one
        // flag meaning "the close question is already answered", and it is what keeps that answer
        // independent of whatever WouldDiscardEnteredWork happens to cover. The comment claimed Save
        // set it while only the discard paths actually did; the behaviour was right by accident,
        // through that predicate short-circuiting on a separate flag.
        private bool thisIsClosingConfirmed;

        // What the draft's sub-lists held the moment it was created, so WouldDiscardEnteredWork can
        // tell the user's own additions from what the editor seeded itself - see its own comment.
        private int thisDraftInitialCommentCount;
        private int thisDraftInitialWorkDoneCount;
        private int thisDraftInitialLinkCount;
        private int thisDraftInitialPhotoCount;
        private int thisDraftInitialFileCount;

        // The draft entry actually saved, for the caller that has to know what was written (the
        // Schematics tab refreshes its overlay against it). Null unless a draft reached disk.
        public WorklogEntryRecord? SavedNewEntry { get; private set; }

        // Guards against Initialize()'s own seeding of the direct fields (Title/Description text,
        // category/state selection) being mistaken for a user edit and enabling Save prematurely.
        private bool thisIsInitializing;

        // Set by PersistEntrySilently, so that even a Cancel/Escape close reports WasSaved = true
        // when a Links/Comments/Work-done/Photos/Files change already made it to disk.
        private bool thisHasPersistedChange;

        public bool WasSaved { get; private set; }

        // The window's last NON-maximized bounds, tracked continuously so they are correct however
        // the window is closed. Persisting this.Width/Height directly would store the maximized
        // size, and un-maximizing on the next open would then restore to full screen with no memory
        // of the size the user actually chose. Same approach as ComponentInfoWindow.
        private double thisNormalWidth;
        private double thisNormalHeight;

        // Nullable on purpose: null means "this window has never reported a normal-state position",
        // which is NOT the same as being at (0,0). Storing 0 for the unknown case made a first run
        // persist a top-left position and set the has-layout flag, so every later open was pinned to
        // the corner of the primary screen instead of centring on its owner - permanently, since
        // each close rewrote the same zeros.
        private int? thisNormalX;
        private int? thisNormalY;

        // ###########################################################################################
        // Which screen the window is on RIGHT NOW - tracked whenever Position changes, in EVERY
        // WindowState, unlike thisNormalX/Y above which only track Normal state.
        //
        // That difference is the point: dragging a maximized window to another monitor keeps it
        // maximized there without ever passing through Normal state, so thisNormalX/Y never
        // updates - it still names whichever monitor the window was last WINDOWED on, which can be
        // a different one entirely. Reported: maximize on monitor 2, Cancel, "Add worklog" again -
        // the new window opened maximized on monitor 1 instead. Restoring a maximized window has to
        // move it onto this screen FIRST and maximize second, the same two-step Main.axaml.cs's own
        // window placement uses for its own top-level window, for the same reason.
        //
        // Null until the first PositionChanged, which is enough: RestoreWindowPlacement only reads
        // it on the Maximized branch, and a window that has never moved has nothing to disagree
        // with the saved value about.
        private PixelPoint? thisCurrentScreenTopLeft;

        // ###########################################################################################
        // Test seam: when false, the window keeps the size and split its XAML declares instead of
        // restoring the user's saved placement, and does not persist on close.
        //
        // The headless UI tests build this window on a developer's real machine, where UserSettings
        // is the live settings file - so without this a test asserting anything about the layout is
        // really asserting whatever size and splitter position the developer last left the editor
        // in. That is not a hypothetical: it made the responsiveness and splitter tests pass alone
        // and fail in the suite. Tests that specifically exercise persistence set this to true and
        // point UserSettings at a temp file first.
        //
        // Defaults to true so the shipping app is unaffected; only a test ever turns it off.
        // ###########################################################################################
        internal static bool PersistWindowPlacement { get; private set; } = true;

        // ###########################################################################################
        // Turns placement persistence off for the duration of a using-block, then restores whatever
        // it was before.
        //
        // The flag used to be set directly and never put back, which made it a one-way latch: the
        // first test to disable it disabled it for every window built afterwards in the shared
        // headless session, so anything written later to exercise RestoreWindowPlacement or
        // TrackWindowPlacement would pass vacuously. A scope makes the off-state bounded even when
        // the body throws - the same discipline the ColumnDefinition restore in
        // WorklogEditorSplitterTests already uses.
        // ###########################################################################################
        internal static IDisposable SuppressWindowPlacementPersistence()
        {
            var scope = new PlacementPersistenceScope(PersistWindowPlacement);
            PersistWindowPlacement = false;
            return scope;
        }

        private sealed class PlacementPersistenceScope : IDisposable
        {
            private readonly bool thisPrevious;

            public PlacementPersistenceScope(bool previous) => this.thisPrevious = previous;

            public void Dispose() => PersistWindowPlacement = this.thisPrevious;
        }

        public WorklogEntryEditorWindow()
        {
            this.InitializeComponent();

            if (PersistWindowPlacement)
            {
                this.RestoreWindowPlacement();
                this.TrackWindowPlacement();
            }

            this.EditorLinksList.ItemsSource = this.thisLinkRows;
            this.EditorCommentsList.ItemsSource = this.thisCommentRows;
            this.EditorWorkDoneList.ItemsSource = this.thisWorkDoneRows;
            this.EditorPhotosList.ItemsSource = this.thisPhotoRows;
            this.EditorFilesList.ItemsSource = this.thisFileRows;
            this.EditorComponentList.ItemsSource = this.thisComponentRows;
            this.EditorCompletedComponentList.ItemsSource = this.thisCompletedComponentRows;

            this.InitializeListSections();

            // The marker's position is computed from EditorLocationPreviewGrid's OWN size, so the
            // redraw has to be driven by that grid rather than by the window. Dragging the
            // GridSplitter re-widths the preview column while the window's size never changes, so
            // a window-level SizeChanged does not fire and the marker kept coordinates computed
            // for the previous width - it drifted away from the area it is meant to mark, and only
            // snapped back when the window itself was resized.
            this.EditorLocationPreviewGrid.SizeChanged += (_, _) => this.RefreshLocationPreviewOverlay();

            this.AddHandler(KeyDownEvent, this.OnWindowPreviewKeyDown, RoutingStrategies.Tunnel);

            // The photo drag's move/release live on the LIST, not on the row that started it: the
            // dragged row is re-rendered as an empty placeholder the moment the drag begins, which
            // takes its own handlers out of the tree, and the row also moves out from under the
            // pointer as the list reorders. The list stays put for the whole gesture.
            // Tunnel so a release over a row's buttons still ends the drag rather than being eaten.
            this.EditorPhotosList.AddHandler(PointerMovedEvent, this.OnPhotoRowDragHandlePointerMoved, RoutingStrategies.Tunnel);
            this.EditorPhotosList.AddHandler(PointerReleasedEvent, this.OnPhotoRowDragHandlePointerReleased, RoutingStrategies.Tunnel);

            // The Files list drags through the same handlers - which list is being reordered comes
            // from the DragContext captured on press, not from which control raised the event.
            this.EditorFilesList.AddHandler(PointerMovedEvent, this.OnPhotoRowDragHandlePointerMoved, RoutingStrategies.Tunnel);
            this.EditorFilesList.AddHandler(PointerReleasedEvent, this.OnPhotoRowDragHandlePointerReleased, RoutingStrategies.Tunnel);

            // A release outside the list (dragged past the window edge, say) never reaches the
            // handlers above, which would strand the placeholder as a permanent empty slot. The
            // window-level handler commits the drop at wherever the placeholder currently sits.
            //
            // BUBBLE, not Tunnel. The window is the root, so on the tunnelling route it would fire
            // BEFORE the lists and commit every in-list drop itself - making the lists' own handlers
            // dead code and this "fallback" the actual primary path. Bubbling runs it last, so it
            // only ever sees a release the lists did not already handle, which is what the comment
            // above describes and what the release handler's early-return assumes.
            this.AddHandler(PointerReleasedEvent, this.OnPhotoRowDragHandlePointerReleased, RoutingStrategies.Bubble);

            // The thumbnails this window decoded hold unmanaged surfaces; without this the last set
            // survives the window itself. thisSchematicBitmap belongs to the caller and is not
            // touched here.
            this.Closed += (_, _) =>
            {
                foreach (var row in this.thisPhotoRows)
                {
                    row.Thumbnail?.Dispose();
                }

                // Teardown matches construction. A gesture still in flight when the window closes
                // (released outside it, so no release handler ever ran) leaves thisActiveDragContext
                // holding the entry's live lists and two bound delegates; clearing the collections
                // and the drag state drops those references with the window instead of after it.
                this.ResetPhotoDragState();

                this.thisPhotoRows.Clear();
                this.thisFileRows.Clear();
                this.thisLinkRows.Clear();
                this.thisCommentRows.Clear();
                this.thisWorkDoneRows.Clear();
                this.thisComponentRows.Clear();
                this.thisCompletedComponentRows.Clear();
            };
        }

        // ###########################################################################################
        // Restores the size, position and maximized state this window was last closed with.
        //
        // Position is only applied when something was actually saved: without it the window would
        // be placed at (0,0) on a first run instead of honouring WindowStartupLocation="CenterOwner".
        // The saved position is also range-checked against the available screens, so a window last
        // closed on a monitor that is no longer attached does not open off-screen where it cannot be
        // reached - it falls back to centring on the owner.
        // ###########################################################################################
        private void RestoreWindowPlacement()
        {
            this.thisNormalWidth = UserSettings.HasWorklogEntryWindowLayout
                ? UserSettings.WorklogEntryWindowWidth
                : this.Width;

            this.thisNormalHeight = UserSettings.HasWorklogEntryWindowLayout
                ? UserSettings.WorklogEntryWindowHeight
                : this.Height;

            if (!UserSettings.HasWorklogEntryWindowLayout)
            {
                return;
            }

            // Clamped to the window's own minimums, so a settings file carrying a smaller size (or
            // a hand-edited one) cannot produce a window too small to use.
            this.Width = Math.Max(this.MinWidth, UserSettings.WorklogEntryWindowWidth);
            this.Height = Math.Max(this.MinHeight, UserSettings.WorklogEntryWindowHeight);

            this.thisNormalWidth = this.Width;
            this.thisNormalHeight = this.Height;

            int savedX = UserSettings.WorklogEntryWindowX;
            int savedY = UserSettings.WorklogEntryWindowY;

            bool restoreMaximized = string.Equals(
                UserSettings.WorklogEntryWindowState, "Maximized", StringComparison.OrdinalIgnoreCase);

            if (this.IsSavedPositionOnAScreen(savedX, savedY))
            {
                this.thisNormalX = savedX;
                this.thisNormalY = savedY;
                this.WindowStartupLocation = WindowStartupLocation.Manual;

                // For the Maximized case this Normal-state position is about to be overwritten
                // below by the saved SCREEN's position - see restoreMaximized. Set unconditionally
                // anyway, or a window with no valid saved screen (see the fallback below) would
                // start unpositioned instead of at least landing on the last windowed spot.
                this.Position = new PixelPoint(savedX, savedY);
            }

            if (restoreMaximized)
            {
                // Move onto the saved SCREEN before maximizing, not the saved WINDOWED position -
                // see thisCurrentScreenTopLeft's comment for why the two can name different
                // monitors. +100,+100 only has to land inside the screen, matching the nudge
                // Main.axaml.cs's own window placement uses for the same reason.
                int screenX = UserSettings.WorklogEntryWindowScreenX;
                int screenY = UserSettings.WorklogEntryWindowScreenY;

                // Only moved when the saved screen was actually RECORDED. IsSavedPositionOnAScreen
                // deliberately accepts anything when the screen list is unavailable (see its header),
                // which is the right answer for a genuine saved position but the wrong one here: a
                // settings file upgraded from a build that saved the window state but not the screen
                // carries (0,0), and that is not a position the user chose. Nudging to (100,100)
                // from it maximizes the window on whichever monitor happens to contain that point,
                // rather than leaving the OS to maximize where the window already is.
                bool hasSavedScreen = screenX != 0 || screenY != 0;

                if (hasSavedScreen && this.IsSavedPositionOnAScreen(screenX, screenY))
                {
                    this.Position = new PixelPoint(screenX + 100, screenY + 100);
                }

                this.WindowState = WindowState.Maximized;
            }

            // The splitter, as the left column's share of the two content columns. Clamped so a
            // corrupt or hand-edited value cannot collapse either side to nothing - the MinWidths on
            // the columns would fight it, and the result is a splitter that will not move.
            double ratio = Math.Clamp(UserSettings.WorklogEntryWindowLeftColumnRatio, 0.15, 0.85);
            this.EditorSplitGrid.ColumnDefinitions[0].Width = new GridLength(ratio, GridUnitType.Star);
            this.EditorSplitGrid.ColumnDefinitions[2].Width = new GridLength(1.0 - ratio, GridUnitType.Star);
        }

        // ###########################################################################################
        // The left column's share of the two content columns, as laid out right now.
        //
        // Measured from the actual bounds rather than read back from the ColumnDefinitions: dragging
        // a GridSplitter rewrites those definitions, but reading the star VALUES back would mean
        // reconstructing the proportion from two numbers whose units depend on how the splitter left
        // them. The rendered widths are unambiguous. Returns the saved value unchanged when the
        // window has not been laid out (bounds still zero), so closing an unshown window cannot
        // overwrite a good setting with a meaningless one.
        // ###########################################################################################
        private double CurrentLeftColumnRatio()
        {
            double leftWidth = this.EditorSplitGrid.ColumnDefinitions[0].ActualWidth;
            double rightWidth = this.EditorSplitGrid.ColumnDefinitions[2].ActualWidth;
            double total = leftWidth + rightWidth;

            if (total <= 0.0)
            {
                return UserSettings.WorklogEntryWindowLeftColumnRatio;
            }

            return Math.Clamp(leftWidth / total, 0.15, 0.85);
        }

        // ###########################################################################################
        // True when the saved top-left lands inside one of the currently connected screens.
        //
        // Guards the monitor-unplugged case: a position saved on a second display would otherwise
        // put the window somewhere with no screen, where it cannot be moved or closed. Screens can
        // be unavailable this early in construction, in which case the position is accepted - the
        // OS will not place a window entirely off-screen on its own.
        // ###########################################################################################
        private bool IsSavedPositionOnAScreen(int x, int y)
        {
            var screens = this.Screens;
            if (screens == null || screens.ScreenCount == 0)
            {
                return true;
            }

            foreach (var screen in screens.All)
            {
                if (screen.Bounds.Contains(new PixelPoint(x, y)))
                {
                    return true;
                }
            }

            return false;
        }

        // ###########################################################################################
        // Keeps the normal-state bounds, and the current screen, current - then writes them out
        // when the window closes.
        //
        // The size/position trackers ignore anything but WindowState.Normal, which is what keeps a
        // maximized session from overwriting the restore size - see the fields above. The screen
        // tracker is the one exception and runs in EVERY state, including Maximized - see
        // thisCurrentScreenTopLeft's own comment for why.
        // ###########################################################################################
        private void TrackWindowPlacement()
        {
            this.SizeChanged += (_, _) =>
            {
                if (this.WindowState == WindowState.Normal)
                {
                    this.thisNormalWidth = this.Width;
                    this.thisNormalHeight = this.Height;
                }
            };

            this.PositionChanged += (_, _) =>
            {
                if (this.WindowState == WindowState.Normal)
                {
                    this.thisNormalX = this.Position.X;
                    this.thisNormalY = this.Position.Y;
                }

                this.UpdateCurrentScreenTopLeft();
            };

            // Closing, not Closed: the window's bounds are still meaningful here. It fires for every
            // route out - Save, Cancel, Escape and the title-bar close - so no exit path loses the
            // placement.
            this.Closing += (_, closingArgs) =>
            {
                // THE TITLE-BAR CLOSE does not go through OnCancelClick, so the discard guard has to
                // be applied here too - otherwise the one exit route that bypasses Cancel is also
                // the one that silently loses the work.
                //
                // Closing cannot await, so this CANCELS the close, asks, and closes again from the
                // dialog's continuation if the user confirms. thisIsClosingConfirmed stops that
                // second close re-entering this branch and asking forever.
                if (!this.thisIsClosingConfirmed && this.WouldDiscardEnteredWork())
                {
                    closingArgs.Cancel = true;

                    _ = Dispatcher.UIThread.InvokeAsync(async () =>
                    {
                        if (await this.ConfirmDiscardIfNeededAsync())
                        {
                            this.thisIsClosingConfirmed = true;
                            this.WasSaved = false;
                            this.Close(false);
                        }
                    });

                    return;
                }

                // The title-bar close does not go through OnCancelClick, so an abandoned draft's
                // attachment bytes would survive that one exit route. DiscardDraftAttachments is
                // idempotent and only fires while the entry is still a draft.
                //
                // thisWasSuccessfullySaved is checked ALONGSIDE the draft flag rather than relying
                // on Save having cleared that flag first. Save does clear it - but this handler
                // deletes an attachment folder recursively, so "correct because two lines in another
                // method happen to be in this order" is not a safe basis for it: reordering them, or
                // any future close path that does not clear the flag, would delete the photos and
                // files of an entry that had just saved successfully. Two independent conditions
                // mean either one being right is enough.
                if (this.thisIsDraftEntry && !this.thisWasSuccessfullySaved)
                {
                    this.DiscardDraftAttachments();
                }

                string state = this.WindowState == WindowState.Maximized ? "Maximized" : "Normal";

                // One last update in case the window closed WITHOUT PositionChanged ever firing
                // after its final move - a maximize-then-close on a monitor the window had never
                // visited before while windowed can do that.
                this.UpdateCurrentScreenTopLeft();
                var screenTopLeft = this.thisCurrentScreenTopLeft
                    ?? new PixelPoint(UserSettings.WorklogEntryWindowScreenX, UserSettings.WorklogEntryWindowScreenY);

                // Falls back to whatever is already stored when this window never reported a
                // normal-state position, rather than inventing (0,0) - see the fields above.
                UserSettings.SaveWorklogEntryWindowLayout(
                    state,
                    this.thisNormalWidth,
                    this.thisNormalHeight,
                    this.thisNormalX ?? UserSettings.WorklogEntryWindowX,
                    this.thisNormalY ?? UserSettings.WorklogEntryWindowY,
                    screenTopLeft.X,
                    screenTopLeft.Y,
                    this.CurrentLeftColumnRatio());
            };
        }

        // ###########################################################################################
        // Records the top-left of whichever screen the window's CURRENT position falls on - see
        // thisCurrentScreenTopLeft's own comment for why this has to run in every WindowState, not
        // only Normal.
        //
        // Matches by containment against this.Position, the same technique IsSavedPositionOnAScreen
        // uses for the inverse check. Left unset (not overwritten with a guess) when no screen
        // contains the point - headless/disconnected-monitor edge cases - so the fallback in the
        // Closing handler above can fall back to the last known-good value instead of persisting a
        // wrong one.
        // ###########################################################################################
        private void UpdateCurrentScreenTopLeft()
        {
            var screens = this.Screens;
            if (screens == null)
            {
                return;
            }

            var position = this.Position;

            foreach (var screen in screens.All)
            {
                if (screen.Bounds.Contains(position))
                {
                    this.thisCurrentScreenTopLeft = new PixelPoint(screen.Bounds.X, screen.Bounds.Y);
                    return;
                }
            }
        }

        // ###########################################################################################
        // Escape acts like Cancel. Plain Enter
        // in the single-line Title field saves and closes (Title has no use for a literal newline);
        // in the multi-line Description field (AcceptsReturn) plain Enter is left alone so it keeps
        // inserting a newline, and only Ctrl+Enter saves - same convention as WorklogAddCommentWindow.
        // Handled on the Tunnel route so this runs before Description's own AcceptsReturn handling
        // inserts a newline - a bubbling KeyDown handler would run too late to stop that. Save only
        // actually commits when it is enabled (a direct field has been edited); otherwise Enter/
        // Ctrl+Enter is a no-op, same as clicking a disabled Save button would be.
        // ###########################################################################################
        private void OnWindowPreviewKeyDown(object? sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
            {
                this.OnCancelClick(sender, e);
                e.Handled = true;
                return;
            }

            if (e.Key != Key.Enter)
                return;

            bool isDescriptionFocused = ReferenceEquals(e.Source, this.EditorDescriptionTextBox);
            if (isDescriptionFocused && !e.KeyModifiers.HasFlag(KeyModifiers.Control))
                return;

            bool isTitleFocused = ReferenceEquals(e.Source, this.EditorTitleTextBox);
            if (!isDescriptionFocused && !isTitleFocused)
                return;

            e.Handled = true;

            if (this.EditorSaveButton.IsEnabled)
            {
                this.OnSaveClick(sender, e);
            }
        }

        // ###########################################################################################
        // Must be called before showing the dialog: seeds every field/list from the given entry and
        // loads the schematic bitmap for the location preview. workbookId is needed separately since
        // WorklogEntryRecord itself does not know which workbook it belongs to.
        // ###########################################################################################
        public void Initialize(int workbookId, WorklogEntryRecord entry, Bitmap? schematicBitmap)
        {
            this.thisIsInitializing = true;

            this.thisWorkbookId = workbookId;
            this.thisEntry = CloneEntry(entry);
            this.thisSchematicBitmap = schematicBitmap;

            // "Update worklog" for a saved entry; InitializeForNewEntry overrides this to "Add
            // worklog" straight afterwards for a draft. Set here rather than left as the markup's
            // default so the two callers cannot drift from each other.
            this.EditorSaveButton.Content = "Update worklog";

            this.EditorIdText.Text = $"#{this.thisEntry.Id}";
            this.EditorTitleTextBox.Text = this.thisEntry.Title;
            this.EditorDescriptionTextBox.Text = this.thisEntry.Description;
            this.EditorLocationSchematicNameText.Text = this.thisEntry.SchematicName;
            this.EditorShowMarkedAreaCheckBox.IsChecked = this.thisEntry.ShowMarkedArea;

            this.thisSelectedCategory = string.IsNullOrWhiteSpace(this.thisEntry.Category) ? "Note" : this.thisEntry.Category;
            this.thisSelectedState = string.IsNullOrWhiteSpace(this.thisEntry.State) ? "Open" : this.thisEntry.State;
            this.UpdateCategoryChipVisuals();
            this.UpdateStatePillVisuals();

            // Heals duplicate/gapped DisplayOrder values left by older builds before anything is
            // rendered, so the list cannot show two rows in an arbitrary order. Working-copy only -
            // it reaches disk with the next save rather than writing on open.
            WorklogAttachmentStorage.NormalizeDisplayOrder(this.thisEntry.Photos);
            WorklogAttachmentStorage.NormalizeDisplayOrder(this.thisEntry.Files);

            this.RefreshLinkRows();
            this.RefreshCommentRows();
            this.RefreshWorkDoneRows();
            this.RefreshPhotoRows();
            this.RefreshFileRows();

            this.EditorLocationPreviewImage.Source = this.thisSchematicBitmap;
            this.RefreshLocationPreviewOverlay();

            // After the lists are built, so the sections fold over real row counts and the
            // empty-state lines settle correctly. Inside the initializing guard, so restoring the
            // user's saved folds cannot itself write them back to disk.
            this.RestoreCollapsedSections();
            this.RefreshListSectionEmptyStates();

            this.thisIsDirty = false;
            this.EditorSaveButton.IsEnabled = false;

            // The initializing guard is lifted on the dispatcher, not here. Setting TextBox.Text
            // above does not raise TextChanged synchronously - Avalonia posts it - so clearing the
            // flag inline let Initialize's OWN title and description assignments arrive afterwards
            // and mark the untouched window dirty. Every editor therefore opened with Save already
            // enabled, contradicting the "starts disabled" rule this class is built around, and a
            // straight open-and-close reported an edit that never happened.
            //
            // Posting at Background priority puts the lift behind those queued TextChanged jobs,
            // so they run while the guard is still up and are correctly ignored.
            Dispatcher.UIThread.Post(
                () =>
                {
                    this.thisIsInitializing = false;
                    this.thisIsDirty = false;
                    this.UpdateSaveButtonEnabled();
                },
                DispatcherPriority.Background);
        }

        // ###########################################################################################
        // Must be called before showing the dialog when the entry does NOT yet exist - the "Add
        // worklog" flow: the user picks "Add worklog" in the top bar, drags out an area on the
        // schematic, and this window opens on it directly.
        //
        // It replaced a small "New fault" card that asked for a title, description, category, state
        // and the component checklist, and then had to be reopened in THIS window to reach anything
        // else. Two dialogs for one entry, with the first one's fields a strict subset of the
        // second's - so the first was removed outright rather than kept as a shortcut.
        //
        // The entry is held in memory until Save (see thisIsDraftEntry): nothing is written on open,
        // so a Cancel here leaves the workbook exactly as it was. The id is RESERVED rather than
        // allocated - the attachment folder has to be named after something before Save - and
        // WorklogManager.AddEntryRecord re-allocates and moves that folder if another entry claimed
        // the number meanwhile.
        //
        // Category and state start at Note/Open, the same defaults the card it replaces used.
        // ###########################################################################################
        public void InitializeForNewEntry(
            int workbookId,
            string schematicName,
            Rect area,
            Bitmap? schematicBitmap)
        {
            this.thisIsDraftEntry = true;
            this.thisDraftReservedEntryId = WorklogManager.PeekNextEntryId(workbookId);

            var draft = new WorklogEntryRecord
            {
                Id = this.thisDraftReservedEntryId,
                SchematicName = schematicName ?? string.Empty,
                AreaX = area.X,
                AreaY = area.Y,
                AreaWidth = area.Width,
                AreaHeight = area.Height,
                Category = "Note",
                State = "Open",
                ShowMarkedArea = true,
                CreatedDate = DateTime.Now
            };

            // Every worklog starts its own history with the fact that it was created - the same
            // audit trail WorklogManager.AddEntry writes, added here because the draft bypasses it.
            WorklogManager.AppendAutomaticComment(draft.Comments, WorklogManager.CreatedCommentText);

            this.Initialize(workbookId, draft, schematicBitmap);

            // AFTER Initialize, which clones the record into thisEntry - the counts have to describe
            // the working copy the guard actually reads, not the local built above.
            this.thisDraftInitialCommentCount = this.thisEntry.Comments.Count;
            this.thisDraftInitialWorkDoneCount = this.thisEntry.WorkDoneItems.Count;
            this.thisDraftInitialLinkCount = this.thisEntry.Links.Count;
            this.thisDraftInitialPhotoCount = this.thisEntry.Photos.Count;
            this.thisDraftInitialFileCount = this.thisEntry.Files.Count;

            // Initialize is written for a saved entry and ends by clearing the dirty flag, which is
            // right there and wrong here: an empty draft has nothing worth saving yet, but the
            // moment the user types a title it must be saveable. UpdateSaveButtonEnabled already
            // gates on a non-blank title, so the flag is simply forced on and the title box decides.
            this.thisIsDirty = true;

            this.Title = "New worklog";
            this.EditorSaveButton.Content = "Add worklog";

            // Same reasoning as Initialize's own deferred lift - its posted job runs after this
            // method returns and would otherwise clear the flag straight back to false.
            Dispatcher.UIThread.Post(
                () =>
                {
                    this.thisIsDirty = true;
                    this.UpdateSaveButtonEnabled();
                    this.EditorTitleTextBox.Focus();
                },
                DispatcherPriority.Background);
        }

        // ###########################################################################################
        // Sets "Show marked area" on a brand-new entry. Call AFTER InitializeForNewEntry, which
        // defaults it to ticked.
        //
        // Exists for the worklog created from an oscilloscope capture, which has no drawn area at
        // all: leaving the box ticked would promise a rectangle that does not exist, and the entry
        // would draw as nothing. Unticked, it parks as a "#N" pill in the corner instead, which is
        // the state the app already supports for an entry with no area to sit on.
        // ###########################################################################################
        public void SetShowMarkedAreaForNewEntry(bool showMarkedArea)
        {
            this.EditorShowMarkedAreaCheckBox.IsChecked = showMarkedArea;
            this.thisEntry.ShowMarkedArea = showMarkedArea;
        }

        // ###########################################################################################
        // Attaches an already-captured image to the entry this window is editing, as if it had been
        // added through the Photos section - the "Create new worklog" answer to the oscilloscope
        // capture's attach dialog, where the user wants the capture filed into an entry that does
        // not exist yet.
        //
        // Call AFTER InitializeForNewEntry. That order matters: the draft's attachment folder is
        // named after the id InitializeForNewEntry reserves, so attaching before it would copy the
        // bytes into a folder belonging to whatever entry currently holds that number.
        //
        // For a draft, PersistEntrySilently deliberately writes nothing (see its own comment) - the
        // photo record lives in the working copy until Save, and WorklogManager.AddEntryRecord then
        // writes it and moves the reserved folder if the id has moved meanwhile. So a cancelled
        // draft discards this photo along with everything else, which is the correct outcome: the
        // capture itself is still safe in the oscilloscope image folder regardless.
        //
        // Returns false when the bytes could not be filed - an unresolvable attachments folder, or
        // a failed copy (a locked file, a full disk, permissions). The caller MUST act on that: the
        // editor otherwise opens with an empty Photos section and no message at all, which reads as
        // "the capture was attached" when it was not, and the comment the user typed into the attach
        // dialog is lost with it.
        // ###########################################################################################
        public bool AttachCapturedPhoto(string sourcePath, string comment)
        {
            if (string.IsNullOrWhiteSpace(sourcePath))
            {
                return false;
            }

            var section = this.PhotoAttachments;

            string? attachmentsFolder = WorklogManager.GetEntryAttachmentsFolder(this.thisWorkbookId, this.thisEntry.Id);

            var outcome = WorklogAttachmentWriter.Attach(
                sourcePath,
                attachmentsFolder,
                section.Records(),
                section.OwnerPrefix,
                comment,
                this.PersistEntrySilently,
                out _);

            if (outcome != WorklogAttachmentWriter.AttachOutcome.Added)
            {
                Logger.Warning($"Could not attach captured image to a new worklog entry: {outcome}");
                return false;
            }

            this.EnsureListSectionExpanded(section.HeaderKey);
            this.RefreshAttachmentRows(section);
            return true;
        }

        // ###########################################################################################
        // The Save button starts disabled and is only ever enabled by an edit to one of the direct
        // fields (Title, Description, category, state) - see OnDirectFieldTextChanged and the
        // category/state pointer handlers below. Everything else (links/comments/work done, and
        // delete/reorder on any sub-list) saves itself instantly via PersistEntrySilently, so losing
        // those was never a matter of forgetting to click Save.
        // ###########################################################################################
        private void MarkDirty()
        {
            if (this.thisIsInitializing)
                return;

            this.thisIsDirty = true;
            this.UpdateSaveButtonEnabled();
        }

        private bool thisIsDirty;

        // ###########################################################################################
        // Save is offered only when there is something to save AND the entry is valid - which here
        // means a non-blank title. A worklog with no title is unidentifiable in the worklog list and
        // on the board, where the "#N" badge would be all that distinguishes it.
        //
        // Whitespace does not count: SyncDirectFieldsToEntry Trim()s the title before writing it, so
        // a title of spaces would be persisted as an empty one and the gate has to agree with what
        // the save actually does.
        // ###########################################################################################
        private void UpdateSaveButtonEnabled()
        {
            bool hasTitle = this.HasValidTitle();

            this.EditorSaveButton.IsEnabled = this.thisIsDirty && hasTitle;

            // A disabled Save with no explanation reads as a broken button on a SAVED entry - say
            // why, and say it only when there is actually something waiting to be saved, so merely
            // opening an entry and clearing its title does not scold the user before they have done
            // anything.
            //
            // This matters more than it looks there: SyncDirectFieldsToEntry keeps the STORED title
            // when the box is blank (a blank title must never reach disk), so without a message the
            // window and the file would silently disagree about the title while an instant-save -
            // adding a comment, say - wrote every other field.
            //
            // A brand-new entry (thisIsDraftEntry) skips the message entirely: there is nothing on
            // disk yet to disagree with, and an empty title is simply the window's starting state,
            // not something that needs explaining before the user has typed anything.
            if (this.thisIsDirty && !hasTitle && !this.thisIsDraftEntry)
            {
                this.ShowSaveFailed(BlankTitleMessage);
            }
            else if (string.Equals(this.EditorSaveFailedText.Text, BlankTitleMessage, StringComparison.Ordinal))
            {
                // Only clears OUR message - a real save failure must stay on screen.
                this.EditorSaveFailedText.IsVisible = false;
            }
        }

        private const string BlankTitleMessage = "A worklog needs a title before it can be saved.";

        private bool HasValidTitle() => !string.IsNullOrWhiteSpace(this.EditorTitleTextBox.Text);

        private void OnDirectFieldTextChanged(object? sender, TextChangedEventArgs e)
        {
            this.MarkDirty();
        }

        // ###########################################################################################
        // "Show marked area" is a direct field like the title and category: it marks the window dirty
        // and reaches disk with Save, rather than saving itself the way the sub-lists do. It changes
        // what the board looks like, not what the entry records, so it belongs with the fields the
        // user can still abandon with Cancel.
        // ###########################################################################################
        private void OnShowMarkedAreaCheckedChanged(object? sender, RoutedEventArgs e)
        {
            this.EnsureMarkedAreaExistsWhenShown();
            this.MarkDirty();
        }

        // ###########################################################################################
        // Gives the entry a real, draggable marked area the first time "Show marked area" is ticked
        // on one that has never had a drawn area at all.
        //
        // A worklog created from an oscilloscope capture is stored with no area (see
        // ComponentInfoWindow's create branch) and parks as a corner pill instead. Ticking this box
        // on such an entry used to promise a rectangle that could not exist: a zero-sized rect draws
        // as nothing, or as a hairline, and can never be grabbed and dragged into place - so the
        // entry looked broken with no way to fix it from the UI.
        //
        // WorklogDefaultAreaGeometry places a visible square in the board's BOTTOM-right corner -
        // the opposite corner from the parked pills, so a freshly-placed area cannot be mistaken for
        // one of them - which the user then drags to where it belongs.
        //
        // Only ever ADDS an area, never replaces one: an entry that already has a drawn rectangle
        // keeps it through any number of tick/untick cycles.
        // ###########################################################################################
        private void EnsureMarkedAreaExistsWhenShown()
        {
            if (this.EditorShowMarkedAreaCheckBox.IsChecked != true || this.thisSchematicBitmap == null)
            {
                return;
            }

            var existing = new Rect(
                this.thisEntry.AreaX, this.thisEntry.AreaY, this.thisEntry.AreaWidth, this.thisEntry.AreaHeight);

            // ResolveAreaForShowing owns the "does this entry need an area inventing" decision, so
            // this method never re-derives it: it hands over what the entry has and takes back what
            // it should have. Re-deriving it here is what let the geometry class document itself as
            // the single decision point while the editor quietly kept its own copy of the rule.
            var area = WorklogDefaultAreaGeometry.ResolveAreaForShowing(
                existing,
                new Size(this.thisSchematicBitmap.PixelSize.Width, this.thisSchematicBitmap.PixelSize.Height));

            // Unchanged means the entry already had a real area, and it is never replaced. Still
            // unusable means the image has no size to place one on, so the entry stays parked.
            if (area == existing || WorklogDefaultAreaGeometry.IsUnset(area))
            {
                return;
            }

            this.thisEntry.AreaX = area.X;
            this.thisEntry.AreaY = area.Y;
            this.thisEntry.AreaWidth = area.Width;
            this.thisEntry.AreaHeight = area.Height;

            // The location preview draws from the entry's own area, so it has to be redrawn now that
            // the entry has one - otherwise the box is ticked and the preview still shows nothing.
            this.RefreshLocationPreviewOverlay();
        }

        // ###########################################################################################
        // Copies the direct fields (Title/Description/category/state) out of their controls and into
        // the working copy. Every write to disk must go through this first, because the working copy
        // is only ever updated here - the controls are the live value until it runs.
        // ###########################################################################################
        private void SyncDirectFieldsToEntry()
        {
            // The title is only taken from the box when it actually has one. The sub-lists
            // (links, comments, work done, photos, files) save themselves instantly through
            // PersistEntrySilently, which comes through here - so without this guard, adding a
            // comment while the title box happened to be cleared would write the blank straight
            // to disk, past the Save button that is disabled for exactly that reason.
            string typedTitle = this.EditorTitleTextBox.Text?.Trim() ?? string.Empty;
            if (typedTitle.Length > 0)
            {
                this.thisEntry.Title = typedTitle;
            }

            this.thisEntry.Description = this.EditorDescriptionTextBox.Text?.Trim() ?? string.Empty;
            this.thisEntry.Category = this.thisSelectedCategory;
            this.thisEntry.State = this.thisSelectedState;
            this.thisEntry.ShowMarkedArea = this.EditorShowMarkedAreaCheckBox.IsChecked ?? true;

            // Only when a scope was actually supplied. If the caller could not determine it, the
            // checklist was never shown and the rows are empty - writing that back would silently
            // clear a component list the user never saw, let alone chose to empty.
            //
            // Labels the checklist never offered are CARRIED OVER rather than dropped. The rows
            // come from the highlight rectangles as they are right now, so a label saved earlier
            // whose component has since been renamed or removed from the board data has no row to
            // tick - and this method runs on every instant-save (adding a photo, deleting a file,
            // any drag reorder), not just on Save. Without this the user would lose that label the
            // moment they touched anything unrelated, with no Save click and nothing to notice.
            if (this.thisHasComponentScope)
            {
                var offered = new HashSet<string>(
                    this.thisComponentRows.Select(r => r.BoardLabel),
                    StringComparer.OrdinalIgnoreCase);

                var keptFromBeforeOpening = (this.thisEntry.ComponentLabels ?? new List<string>())
                    .Where(label => !offered.Contains(label))
                    .ToList();

                this.thisEntry.ComponentLabels = this.thisComponentRows
                    .Where(r => r.IsChecked)
                    .Select(r => r.BoardLabel)
                    .Concat(keptFromBeforeOpening)
                    .ToList();

                // The completed list is written from the rows on screen, then narrowed to the scope
                // that was just written. The narrowing is what enforces the invariant that a
                // completed label is always a component the entry actually covers - including for
                // the labels carried over above, which have no row here to tick and so must not
                // survive as completed on the strength of an older save.
                this.thisEntry.CompletedComponentLabels = ComponentListBuilder.NarrowSelectionToScope(
                    this.thisCompletedComponentRows
                        .Where(r => r.IsChecked)
                        .Select(r => r.BoardLabel)
                        .ToList(),
                    this.thisEntry.ComponentLabels);
            }
        }

        // ###########################################################################################
        // Persists the working copy immediately, the same way Save does, but without touching
        // WasSaved or closing the window - used after every add/edit/delete on the Links/Comments/
        // Work done/Photos/Files sub-lists so none of that is lost if the window is later closed via
        // Cancel or Escape without the direct fields ever having been touched.
        //
        // It syncs the direct fields first, and MUST keep doing so. Without that, adding a comment
        // after retyping the headline wrote the record with the OLD headline and state, silently
        // reverting what the user had just typed - the sub-list write cannot save "only its half"
        // of the record, because UpdateEntry replaces the whole thing. A consequence worth knowing:
        // an instant-save therefore commits in-progress direct-field edits too, so Cancel can no
        // longer discard them. That is deliberate - it matches what is on screen, and Cancel already
        // could not undo an instant-saved sub-list change.
        // ###########################################################################################
        // Returns whether THIS save reached disk. thisHasPersistedChange cannot answer that - it is
        // sticky for the window's lifetime so Cancel can still report WasSaved - so a caller that
        // must not act on a failed save (deleting an attachment's bytes, say) reads this instead.
        private bool PersistEntrySilently()
        {
            this.SyncDirectFieldsToEntry();

            // A DRAFT has nothing on disk to update, and deliberately writes nothing until Save -
            // that is what lets Cancel discard a half-made new entry cleanly. The sub-list change
            // the caller just made is already in the working copy, which is the whole record Save
            // will write, so reporting success here is accurate: nothing was lost.
            //
            // What the true therefore MEANS is "the change is safely recorded", not "a file was
            // written" - and that is exactly what every caller reads it for. The attachment delete
            // paths gate their byte deletion on it, and for a draft that is still right: the row is
            // gone from the working copy Save will write, so the bytes are genuinely unreferenced.
            // Any caller added later that needs "reached disk" specifically has to ask
            // thisIsDraftEntry itself; there is no disk write here to report on.
            if (this.thisIsDraftEntry)
            {
                this.EditorSaveFailedText.IsVisible = false;
                return true;
            }

            if (WorklogManager.UpdateEntry(this.thisWorkbookId, this.thisEntry))
            {
                this.thisHasPersistedChange = true;
                this.EditorSaveFailedText.IsVisible = false;
                return true;
            }

            // "Silently" covers not closing the window and not touching WasSaved - not hiding a
            // failure. The sub-list change the user just made is only in the working copy.
            this.ShowSaveFailed(DefaultSaveFailedMessage);
            return false;
        }

        private const string DefaultSaveFailedMessage = "Could not save - see the log for details.";

        // ###########################################################################################
        // Shows a failure in the footer's status line. Always sets the text rather than only the
        // visibility, because the line is shared: an attachment failure writes its own wording, and
        // without rewriting it a later ordinary save failure would report the attachment's problem.
        // ###########################################################################################
        private void ShowSaveFailed(string message)
        {
            this.EditorSaveFailedText.Text = message;
            this.EditorSaveFailedText.IsVisible = true;
        }

        // ###########################################################################################
        // Deep-enough copy so editing in this window (including list add/delete) cannot mutate the
        // caller's record until Save explicitly commits it back via WorklogManager.UpdateEntry.
        //
        // Every sub-list is null-coalesced. WorklogManager.ReadEntries already normalizes what it
        // loads (see NormalizeEntryCollections there, and why System.Text.Json can produce nulls
        // despite the "= new()" initializers), but this takes a record from a caller rather than
        // straight from disk, and it runs in Initialize before the window is shown - so an
        // unguarded dereference here would throw before the user ever saw the editor.
        // ###########################################################################################
        private static WorklogEntryRecord CloneEntry(WorklogEntryRecord source)
        {
            return new WorklogEntryRecord
            {
                Id = source.Id,
                SchematicName = source.SchematicName,
                AreaX = source.AreaX,
                AreaY = source.AreaY,
                AreaWidth = source.AreaWidth,
                AreaHeight = source.AreaHeight,
                Title = source.Title,
                Description = source.Description,
                Category = source.Category,
                State = source.State,
                ComponentLabels = source.ComponentLabels?.ToList() ?? new(),
                CompletedComponentLabels = source.CompletedComponentLabels?.ToList() ?? new(),
                CollapsedSections = source.CollapsedSections?.ToList() ?? new(),
                ShowMarkedArea = source.ShowMarkedArea,
                CreatedDate = source.CreatedDate,
                Links = source.Links?.Select(l => new WorklogLinkRecord { Id = l.Id, Headline = l.Headline, Url = l.Url }).ToList() ?? new(),
                Comments = source.Comments?.Select(c => new WorklogCommentRecord { Id = c.Id, Text = c.Text, Date = c.Date }).ToList() ?? new(),
                WorkDoneItems = source.WorkDoneItems?.Select(w => new WorklogWorkDoneRecord { Id = w.Id, Text = w.Text, Date = w.Date, HoursSpent = w.HoursSpent, Cost = w.Cost, CurrencyCode = w.CurrencyCode }).ToList() ?? new(),
                Photos = source.Photos?.Select(p => new WorklogAttachmentRecord { Id = p.Id, FileName = p.FileName, Comment = p.Comment, DisplayOrder = p.DisplayOrder }).ToList() ?? new(),
                Files = source.Files?.Select(f => new WorklogAttachmentRecord { Id = f.Id, FileName = f.FileName, Comment = f.Comment, DisplayOrder = f.DisplayOrder }).ToList() ?? new(),
            };
        }

        // ###########################################################################################
        // Adds an automatic comment to the working copy, shows it, and writes it straight to disk.
        //
        // Persisted immediately rather than waiting for Save, matching every other sub-list change
        // in this window: a comment the user can see in the list but which vanishes on Cancel would
        // be the odd one out, and the audit trail is least useful if it can be discarded.
        //
        // PersistEntrySilently syncs the direct fields too, so the category or state that prompted
        // the comment reaches disk with it - the two can never disagree.
        // ###########################################################################################
        private void RecordAutomaticComment(string? text)
        {
            if (WorklogManager.AppendAutomaticComment(this.thisEntry.Comments, text) == null)
                return;

            // Deliberately does NOT expand the Comments section. Only a direct "Add comment" click
            // does that - the user asked to write a comment there, so they should see it land.
            // Flipping a status is not that request; unfolding a list they had folded away, every
            // time they touch a pill, is the app second-guessing them.
            this.RefreshCommentRows();
            this.PersistEntrySilently();
        }
    }
}
