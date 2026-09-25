using Avalonia;
using Avalonia.Controls;
using Avalonia.Interactivity;
using Handlers.DataHandling;
using System;
using System.Linq;
using Tabs.TabSchematics;

namespace CRT
{
    // ###########################################################################################
    // The Schematics tab's two auxiliary windows: the detached thumbnails window (detach/reattach,
    // open/close, per-board layout) and the fullscreen schematics window (enter/exit, placeholder
    // content, positioning, re-owning the thumbnails window while it is open). See Main.axaml.cs
    // for the file map of the whole partial class.
    // ###########################################################################################
    public partial class Main
    {
        // ###########################################################################################
        // Shows/hides the Schematics tab's inline thumbnail column and opens/closes the detached
        // thumbnails window to match the "Detach thumbnails into their own window" setting.
        //
        // Composes with F11 schematics fullscreen rather than excluding it: fullscreen collapses the
        // same column this mode does, so the two agree about the layout and only have to agree about
        // who restores the strip. EnterThumbnailsDetachedMode sets its flag and leaves the columns
        // alone while fullscreen is active, and ExitFullscreenMode checks that flag before bringing
        // the strip back - so the detached window stays usable beside a fullscreen schematic, and
        // leaving fullscreen lands on whichever state the setting actually describes.
        // ###########################################################################################
        // ###########################################################################################
        // The ONE way to turn "Detach thumbnails into their own window" on or off: persist the setting,
        // apply the layout, and resync the Configuration checkbox to match. Every surface that
        // offers the toggle goes through here - the checkbox itself, the thumbnail panel's
        // right-click menu, and the detached window's own OS close button - so a step added to the
        // feature is added once rather than found and updated at three call sites.
        //
        // Resyncing the checkbox is harmless when the checkbox IS the caller: it re-reads the
        // setting it just wrote, behind the suppress flag that stops it re-entering its own handler.
        // ###########################################################################################
        internal void SetThumbnailsDetached(bool isDetached)
        {
            string boardKey = this.GetCurrentBoardKey();

            // Per board rather than the single global flag: this board's own detach state is
            // what "Detach thumbnails into their own window" now means when it is toggled, matching
            // the checkbox that gates the setting rather than the global default it falls back
            // to for a board that has never had its own choice made yet (see
            // ResolveThumbnailsDetachedForCurrentBoard).
            //
            // With NO board selected there is no per-board record to write: GetCurrentBoardKey
            // documents an empty string for that state, and storing under it would give every
            // no-selection moment one shared "" entry that the next real board could inherit a
            // detach state from. The global flag is the honest place for the choice then, and it is
            // also what ResolveThumbnailsDetachedForCurrentBoard falls back to for the same key.
            if (UserSettings.RememberThumbnailWindowSettingsPerBoard && !string.IsNullOrEmpty(boardKey))
            {
                UserSettings.SetThumbnailWindowDetachedForBoard(boardKey, isDetached);
            }
            else
            {
                UserSettings.DetachSchematicsThumbnails = isDetached;
            }

            this.ApplyThumbnailsDetachedState();
            this.RefreshDetachThumbnailsCheckBox();
        }

        // ###########################################################################################
        // Resyncs the Configuration tab's "Detach thumbnails into their own window" checkbox with the
        // state actually in force for the current board. Hands the resolved value over rather than
        // letting the tab read a setting for itself, so the checkbox cannot end up showing a
        // different tier from the one the layout is following - see
        // ResolveThumbnailsDetachedForCurrentBoard.
        // ###########################################################################################
        internal void RefreshDetachThumbnailsCheckBox()
            => this.TabConfiguration.RefreshDetachSchematicsThumbnailsCheckBoxFromSettings(
                this.ResolveThumbnailsDetachedForCurrentBoard());

        // ###########################################################################################
        // Whether thumbnails should be shown detached for the CURRENT board: the per-board saved
        // choice while "Remember thumbnail window settings per board" is on (embedded, as asked
        // for, when this board has never had one saved), or the single global
        // DetachSchematicsThumbnails flag otherwise. The one place this decision is made, so the
        // checkbox, the board-switch restore and ApplyThumbnailsDetachedState cannot disagree.
        //
        // internal rather than private because the Configuration checkbox has to resync itself
        // against the SAME tier SetThumbnailsDetached writes to. Reading the global flag there while
        // a per-board one was written made the checkbox re-tick itself to the wrong state on every
        // toggle in per-board mode - see RefreshDetachSchematicsThumbnailsCheckBoxFromSettings.
        //
        // An empty board key (no selection - see GetCurrentBoardKey) falls back to the global flag
        // rather than looking up a shared "" record, matching what SetThumbnailsDetached writes in
        // that same state.
        // ###########################################################################################
        internal bool ResolveThumbnailsDetachedForCurrentBoard()
        {
            if (!UserSettings.RememberThumbnailWindowSettingsPerBoard)
                return UserSettings.DetachSchematicsThumbnails;

            string boardKey = this.GetCurrentBoardKey();
            if (string.IsNullOrEmpty(boardKey))
                return UserSettings.DetachSchematicsThumbnails;

            return UserSettings.GetThumbnailWindowSettingsForBoard(boardKey)?.IsDetached == true;
        }

        // ###########################################################################################
        // Re-applies the detached-thumbnails state for the board just switched to, while "Remember
        // thumbnail window settings per board" is on - a no-op otherwise, so board switches behave
        // exactly as before when the setting is off (the single global flag does not vary by board,
        // so there is nothing to re-apply).
        //
        // Always closes and reopens the window on a board change rather than trying to move it: two
        // boards can each want their OWN size/position (a maximized window for one, a small one for
        // another, as asked for), and SchematicsThumbnailsWindow resolves its layout once, in
        // Initialize - so the only way to pick up a different board's saved layout is a fresh
        // window. Closing first also means the CLOSING board's own layout is saved (via its Closing
        // handler) before the new board's window is opened, rather than losing whatever size/position
        // the user just left it at.
        //
        // Skipped entirely before Main itself has ever been shown: PopulateHardwareDropDown (in
        // StartAsync) selects the startup board, and so raises OnBoardSelectionChanged, BEFORE
        // App.axaml.cs calls Main.Show() - Window.Show(owner) throws "Cannot show window with
        // non-visible owner" if attempted this early, which is exactly why the global (non-per-board)
        // open is already deferred to OnWindowFirstOpened (see the constructor's own comment).
        // OnWindowFirstOpened's own detach-open call handles the startup board once the window is
        // actually open, so nothing is lost by skipping here.
        // ###########################################################################################
        private void ApplyThumbnailsDetachedStateForBoardChange()
        {
            if (!UserSettings.RememberThumbnailWindowSettingsPerBoard)
            {
                // An OPEN window still has to be dealt with even in global mode: unticking "Remember
                // thumbnail window settings per board" while one is open used to make this return
                // here forever, so the window kept the thisPerBoardKey it was opened with and wrote
                // the newly selected board's geometry into the PREVIOUS board's saved layout. The
                // window's key is resolved once, in its Initialize, so the only way to correct it is
                // a fresh window.
                //
                // With no window open there is genuinely nothing to re-apply: the global flag does
                // not vary by board, so a board change cannot change the answer.
                if (this._schematicsThumbnailsWindow == null)
                    return;
            }

            if (!this.thisWindowOpenedCompletionSource.Task.IsCompleted)
                return;

            this.CloseThumbnailsWindowForReopen();
            this.ApplyThumbnailsDetachedState();
        }

        // ###########################################################################################
        // Re-applies the detached-thumbnails state after "Remember thumbnail window settings per
        // board" has itself been toggled - the tier that decides both "is this board detached" and
        // "where does its window layout get saved" has just changed under an already-open window, so
        // the answer for the current board can flip and the open window's own board key is now stale.
        //
        // Goes through the same close-and-reopen the board change does rather than just calling
        // ApplyThumbnailsDetachedState: an open window resolved its layout tier once in Initialize
        // (see SchematicsThumbnailsWindow), so a window opened in one tier would keep saving into the
        // other one's record.
        // ###########################################################################################
        internal void ApplyThumbnailsDetachedStateForPerBoardModeChange()
        {
            if (!this.thisWindowOpenedCompletionSource.Task.IsCompleted)
                return;

            this.CloseThumbnailsWindowForReopen();
            this.ApplyThumbnailsDetachedState();
        }

        // ###########################################################################################
        // Closes the detached thumbnails window as BOOKKEEPING - so its Closed handler leaves the
        // detach setting alone instead of reading the close as the user dismissing the window - and
        // lets the caller reopen a fresh one. A no-op when no window is open.
        //
        // Close() raises Closing (saving the closing window's layout under the key it was opened
        // with) then Closed (nulling the field), so the caller's ApplyThumbnailsDetachedState opens a
        // fresh window resolved against the CURRENT board and tier.
        // ###########################################################################################
        private void CloseThumbnailsWindowForReopen()
        {
            var window = this._schematicsThumbnailsWindow;
            if (window == null)
                return;

            this._thumbnailsWindowClosingForReopen = window;
            try
            {
                window.Close();
            }
            finally
            {
                // Only ever clears the marker for THIS window. The Closed closure compares against
                // the instance it was created for, so clearing here cannot race it either way.
                if (ReferenceEquals(this._thumbnailsWindowClosingForReopen, window))
                    this._thumbnailsWindowClosingForReopen = null;
            }
        }

        public void ApplyThumbnailsDetachedState()
        {
            bool isEnabled = this.ResolveThumbnailsDetachedForCurrentBoard();

            if (isEnabled)
            {
                this.TabSchematicsControl.EnterThumbnailsDetachedMode();
                this.OpenSchematicsThumbnailsWindow();
            }
            else
            {
                this.CloseSchematicsThumbnailsWindow();
                this.TabSchematicsControl.ExitThumbnailsDetachedMode();

                // ExitThumbnailsDetachedMode restores column 0/2 to whatever widths were LIVE on the
                // grid when detach mode was entered - which, on a startup with the setting already
                // ticked, is still the XAML-declared default (detach happens in the constructor,
                // before OnBoardSelectionChanged has ever applied the per-board ratio to the live
                // grid). Re-applying the saved ratio here, the same way a board change does, is what
                // actually lands the splitter where the user last dragged it for this board rather
                // than at the default. A no-op when fullscreen owns the columns instead (the call
                // below applies its own skip).
                this.ApplySchematicsSplitterRatioForCurrentBoard(this.GetCurrentBoardKey());
            }

        }

        // ###########################################################################################
        // Applies the per-board Schematics/thumbnail splitter ratio (or the default) to the live
        // inner grid. Shared by OnBoardSelectionChanged (every board change) and
        // ApplyThumbnailsDetachedState (leaving detached-thumbnails mode) so the two cannot drift.
        //
        // Skipped while thumbnails are detached or the schematic is fullscreen: both modes collapse
        // columns 0/2 down to their own layout (see TabSchematics.ThumbnailsDetach.cs and
        // EnterFullscreenMode) - applying it while either is active would silently re-expand column 2
        // right back out from under it, leaving an empty widened column since the inline list itself
        // stays hidden either way.
        // ###########################################################################################
        private void ApplySchematicsSplitterRatioForCurrentBoard(string boardKey)
        {
            if (this.TabSchematicsControl.IsThumbnailsDetached || this.TabSchematicsControl.IsFullscreenModeActive)
                return;

            var innerGrid = this.TabSchematicsControl.FindControl<Grid>("SchematicsInnerGrid");

            if (innerGrid == null)
                return;

            if (UserSettings.HasSchematicsSplitterRatio(boardKey))
            {
                double ratio = UserSettings.GetSchematicsSplitterRatio(boardKey);
                innerGrid.ColumnDefinitions[0].Width = new GridLength(ratio * 100.0, GridUnitType.Star);
                innerGrid.ColumnDefinitions[2].Width = new GridLength((1.0 - ratio) * 100.0, GridUnitType.Star);
            }
            else
            {
                innerGrid.ColumnDefinitions[0].Width = new GridLength(1.0, GridUnitType.Star);
                innerGrid.ColumnDefinitions[2].Width = new GridLength(300.0, GridUnitType.Pixel);
            }
        }

        // ###########################################################################################
        // Opens the detached thumbnails window, bound to the same currentThumbnails collection and
        // hosted (now-hidden) SchematicsThumbnailList the inline gallery uses. A no-op if already open.
        // ###########################################################################################
        private void OpenSchematicsThumbnailsWindow()
        {
            if (this._schematicsThumbnailsWindow != null)
                return;

            var hostedThumbnailList = this.TabSchematicsControl.FindControl<ListBox>("SchematicsThumbnailList");
            if (hostedThumbnailList == null)
                return;

            // null (the app-wide layout tier) both when per-board mode is off AND when no board is
            // selected - GetCurrentBoardKey returns an empty string for the latter, and handing that
            // to the window would save every no-selection layout into one shared "" record that a
            // real board could then restore from. Matches the same guard in SetThumbnailsDetached
            // and ResolveThumbnailsDetachedForCurrentBoard.
            string currentBoardKey = this.GetCurrentBoardKey();
            string? perBoardKey =
                UserSettings.RememberThumbnailWindowSettingsPerBoard && !string.IsNullOrEmpty(currentBoardKey)
                    ? currentBoardKey
                    : null;

            var thumbnailsWindow = new SchematicsThumbnailsWindow();
            this._schematicsThumbnailsWindow = thumbnailsWindow;
            thumbnailsWindow.Initialize(
                this.TabSchematicsControl.currentThumbnails,
                hostedThumbnailList,
                this.TabSchematicsControl,
                perBoardKey);

            thumbnailsWindow.Closed += (_, _) =>
            {
                this._schematicsThumbnailsWindow = null;

                // Not on the way out: this window is OWNED (see ReownSchematicsThumbnailsWindow) by
                // whichever window was the visible schematics surface, so quitting the application
                // closes it too and lands here. Turning the setting off then meant the preference was
                // wiped on every exit and never survived to the next launch.
                if (this._isApplicationShuttingDown)
                    return;

                // A board switch (or a per-board-mode toggle) closing this window to reopen a fresh
                // one is bookkeeping, not the user dismissing it - the closing board's own detached
                // choice must survive for next time it is shown. Compared by INSTANCE, so this does
                // not depend on the marker still being set when Closed happens to fire - see
                // _thumbnailsWindowClosingForReopen.
                if (ReferenceEquals(this._thumbnailsWindowClosingForReopen, thumbnailsWindow))
                    return;

                // Closing the fullscreen window destroys this one along with it: while fullscreen is
                // open this window is OWNED by it (see ReownSchematicsThumbnailsWindow), and closing
                // an owner closes what it owns. That is the fullscreen window going away, not the
                // user dismissing the thumbnails - so the detach setting must survive, exactly as it
                // does on application shutdown above. ReopenThumbnailsWindowAfterFullscreenClosed
                // (in the fullscreen Closed handler) then brings it back owned by this window.
                if (this._isClosingFullscreenWindow)
                    return;

                // Closed via its own OS close button rather than the checkbox - the window is a
                // third way to turn the feature off, so it goes through the same one entry point
                // the other two do. The field above is already null, so the CloseSchematicsThumbnailsWindow
                // inside it finds nothing to close and this cannot recurse.
                //
                // Uses the board this window was opened FOR (perBoardKey), not whatever board is
                // currently selected - a board switch takes the path above instead of reaching here,
                // so by the time this handler runs for a genuine user-initiated close the two already
                // agree; captured explicitly anyway so this cannot silently turn off the WRONG
                // board's setting if that ever changes.
                bool stillDetachedForOpenedBoard = perBoardKey != null
                    ? UserSettings.GetThumbnailWindowSettingsForBoard(perBoardKey)?.IsDetached == true
                    : UserSettings.DetachSchematicsThumbnails;

                if (stillDetachedForOpenedBoard)
                {
                    this.SetThumbnailsDetached(false);
                }
            };

            // Owned by whichever window is CURRENTLY the visible schematics surface - the fullscreen
            // window if one is already open (e.g. ticking "Detach thumbnails" while in fullscreen),
            // this window otherwise. See ReownSchematicsThumbnailsWindow's header for why this must
            // track fullscreen rather than always being Main.
            this._schematicsThumbnailsWindow.Show(
                (Window?)this._schematicsFullscreenWindow ?? this);
        }

        // ###########################################################################################
        // Closes the detached thumbnails window, if open.
        // ###########################################################################################
        private void CloseSchematicsThumbnailsWindow()
        {
            // The field is NOT nulled here: Close() raises Closed synchronously and that handler
            // owns this field's lifetime. Nulling it again afterwards would mask a close that
            // actually refused (a future Closing handler setting e.Cancel), leaving a visible
            // window Main can no longer find or close - and the next toggle opening a second one
            // on top of it.
            this._schematicsThumbnailsWindow?.Close();
        }

        // ###########################################################################################
        // Creates placeholder content for the Schematics tab while fullscreen mode is active.
        // Keeps the thumbnail list available so another schematic can be selected without closing
        // the fullscreen window first.
        // ###########################################################################################
        private SchematicsFullscreenPlaceholder CreateSchematicsFullscreenPlaceholder()
        {
            double ratio = 0.70;
            var boardKey = this.GetCurrentBoardKey();
            if (!string.IsNullOrWhiteSpace(boardKey))
            {
                ratio = Math.Clamp(UserSettings.GetSchematicsSplitterRatio(boardKey), 0.1, 0.9);
            }

            var hostedThumbnailList = this.TabSchematicsControl.FindControl<ListBox>("SchematicsThumbnailList");

            var placeholder = new SchematicsFullscreenPlaceholder();
            placeholder.Initialize(this.TabSchematicsControl.currentThumbnails, hostedThumbnailList, ratio);
            return placeholder;
        }

        // ###########################################################################################
        // Opens the existing schematics viewer in a separate maximized window. internal (not
        // private) so SchematicsThumbnailsWindow's own F11 handler can reach it via TabSchematics.
        // MainWindow - see ToggleSchematicsFullscreenWindow's header for why F11 is a toggle from
        // both windows rather than an open-only action here.
        // ###########################################################################################
        internal void OpenSchematicsFullscreenWindow()
        {
            if (this._schematicsFullscreenWindow != null)
            {
                if (this._schematicsFullscreenWindow.WindowState == Avalonia.Controls.WindowState.Minimized)
                    this._schematicsFullscreenWindow.WindowState = Avalonia.Controls.WindowState.Normal;

                this._schematicsFullscreenWindow.WindowState = Avalonia.Controls.WindowState.Maximized;
                this._schematicsFullscreenWindow.Activate();
                this._schematicsFullscreenWindow.Focus();
                return;
            }

            this.TabSchematicsControl.EnterFullscreenMode();
            this.SchematicsTabItem.Content = this.CreateSchematicsFullscreenPlaceholder();

            this._schematicsFullscreenWindow = new SchematicsFullscreenWindow(
                this.TabSchematicsControl,
                this.RestoreSchematicsTabContent);

            // Set on Closing rather than around the Close() call sites, because this window closes
            // itself too: Escape and F11 are handled inside SchematicsFullscreenWindow and never pass
            // through Main. Closing runs before the owned thumbnails window is torn down, which is
            // what the thumbnails Closed handler needs in order to tell that cascade apart from the
            // user dismissing it.
            this._schematicsFullscreenWindow.Closing += (_, _) =>
            {
                this._isClosingFullscreenWindow = true;
            };

            this._schematicsFullscreenWindow.Closed += (_, _) =>
            {
                this._isClosingFullscreenWindow = false;
                this._schematicsFullscreenWindow = null;

                // Hand the detached thumbnails window back to THIS window now that the fullscreen
                // window it was following is gone - see ReownSchematicsThumbnailsWindow's header for
                // why it must always own the on-screen schematics surface.
                //
                // Re-owning ALONE is not enough here: closing the fullscreen window closes the
                // windows it owns, so by the time this runs the thumbnails window is usually already
                // gone and the field already null, leaving nothing to re-own. So this reopens it when
                // the setting still says detached, and re-owns it when it somehow survived.
                this.ReopenThumbnailsWindowAfterFullscreenClosed();
            };

            this.PositionFullscreenWindowOnSameScreen(this._schematicsFullscreenWindow);
            this._schematicsFullscreenWindow.WindowState = Avalonia.Controls.WindowState.Maximized;
            this._schematicsFullscreenWindow.Show();

            // The detached thumbnails window (if open) is normally owned by THIS window, which on
            // Windows means clicking into it raises its owner - this window - above the fullscreen
            // window it should stay behind. Re-own it to the fullscreen window instead.
            this.ReownSchematicsThumbnailsWindow();

            this.TabSchematicsControl.RefreshAfterHostChanged();
            this._schematicsFullscreenWindow.Focus();
        }

        // ###########################################################################################
        // Reported: with the thumbnails window detached to its own monitor, clicking into it while
        // schematics fullscreen was active brought THIS window (Main) to the front instead of the
        // fullscreen window, because the thumbnails window's Owner was still Main. Window ownership on
        // Windows is real Win32 owner/HWND parenting (WS_POPUP + GWL_HWNDPARENT), and activating an
        // owned window raises its owner along with it - so the thumbnails window must be owned by
        // whichever window is CURRENTLY the visible schematics surface: the fullscreen window while
        // one is open, this window otherwise. Avalonia's Owner property is live and reassignable after
        // Show() (no close/reopen needed), so this just flips it to match.
        //
        // A no-op when the thumbnails window is not open - there is nothing to re-own, and the next
        // OpenSchematicsThumbnailsWindow call already owns it correctly (Show(this) while fullscreen
        // is inactive, or the ApplyThumbnailsDetachedState call sequence below when it is).
        // ###########################################################################################
        // ###########################################################################################
        // Restores the detached thumbnails window after the schematics fullscreen window has closed.
        //
        // Reported shape of the bug this fixes: while fullscreen is open the thumbnails window is
        // OWNED by it, and closing an owner closes what it owns - so leaving fullscreen destroyed the
        // thumbnails window, and its Closed handler read that as the user dismissing it and silently
        // turned detached mode off. The compensating ReownSchematicsThumbnailsWindow() ran from the
        // fullscreen Closed handler, i.e. after the thumbnails field was already null, so it found
        // nothing to re-own and never prevented anything.
        //
        // The Closed handler now leaves the setting alone during that cascade (_isClosingFullscreenWindow),
        // which makes reopening here the whole of the restore: the setting still says detached, so
        // ApplyThumbnailsDetachedState opens a fresh window - owned by this window, since
        // _schematicsFullscreenWindow is already null by now. A window that somehow survived the
        // cascade is only re-owned, not reopened.
        // ###########################################################################################
        private void ReopenThumbnailsWindowAfterFullscreenClosed()
        {
            if (this._isApplicationShuttingDown)
                return;

            if (this._schematicsThumbnailsWindow != null)
            {
                this.ReownSchematicsThumbnailsWindow();
                return;
            }

            if (!this.thisWindowOpenedCompletionSource.Task.IsCompleted)
                return;

            this.ApplyThumbnailsDetachedState();
        }

        private void ReownSchematicsThumbnailsWindow()
        {
            if (this._schematicsThumbnailsWindow == null)
                return;

            Window desiredOwner = this._schematicsFullscreenWindow ?? (Window)this;

            if (!ReferenceEquals(this._schematicsThumbnailsWindow.Owner, desiredOwner))
            {
                this._schematicsThumbnailsWindow.SetOwnerWindow(desiredOwner);
            }
        }

        // ###########################################################################################
        // The ONE place F11 is interpreted, from EITHER window it can be pressed in (this window's
        // own OnMainKeyDownCloseSinglePopup, and SchematicsThumbnailsWindow's matching handler when
        // that detached window is the active one). Toggles: closes the fullscreen window if one is
        // already open, opens it otherwise - so F11 behaves the same as Escape-to-close while also
        // working as a single repeatable key, which is what was asked for. Escape's own handling in
        // SchematicsFullscreenWindow is untouched.
        // ###########################################################################################
        internal void ToggleSchematicsFullscreenWindow()
        {
            if (this._schematicsFullscreenWindow != null)
            {
                this._schematicsFullscreenWindow.Close();
                return;
            }

            this.OpenSchematicsFullscreenWindow();
        }

        // ###########################################################################################
        // Opens schematics fullscreen from the left-side button.
        // ###########################################################################################
        private void OnFullscreenSchematicsClick(object? sender, RoutedEventArgs e)
        {
            this.OpenSchematicsFullscreenWindow();
        }

        // ###########################################################################################
        // Restores the schematics control back into the normal tab after fullscreen closes.
        // ###########################################################################################
        private void RestoreSchematicsTabContent(Control hostedContent)
        {
            if (!ReferenceEquals(hostedContent, this.TabSchematicsControl))
                return;

            if (!ReferenceEquals(this.SchematicsTabItem.Content, hostedContent))
                this.SchematicsTabItem.Content = hostedContent;

            this.TabSchematicsControl.ExitFullscreenMode();
            this.TabSchematicsControl.RefreshAfterHostChanged();
        }

        // ###########################################################################################
        // Places the fullscreen window on the same screen as the main window before maximizing it.
        // ###########################################################################################
        private void PositionFullscreenWindowOnSameScreen(Window window)
        {
            window.WindowStartupLocation = WindowStartupLocation.Manual;

            double scaling = this.RenderScaling > 0 ? this.RenderScaling : 1.0;
            double w = this.Bounds.Width > 0 ? this.Bounds.Width : this._restoreWidth;
            double h = this.Bounds.Height > 0 ? this.Bounds.Height : this._restoreHeight;

            int centerX = this.Position.X + (int)((w * scaling) / 2);
            int centerY = this.Position.Y + (int)((h * scaling) / 2);

            var screen = this.Screens.All.FirstOrDefault(s =>
                centerX >= s.Bounds.X &&
                centerY >= s.Bounds.Y &&
                centerX < s.Bounds.X + s.Bounds.Width &&
                centerY < s.Bounds.Y + s.Bounds.Height)
                ?? this.Screens.Primary;

            if (screen != null)
            {
                window.Position = new PixelPoint(
                    screen.Bounds.X + 100,
                    screen.Bounds.Y + 100);
            }
        }
    }
}
