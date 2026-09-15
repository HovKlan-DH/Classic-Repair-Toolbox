using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Handlers.DataHandling;
using System;
using System.Collections.ObjectModel;
using Tabs.TabSchematics;

namespace CRT;

// ###########################################################################################
// The detached window shown when "Detach thumbnails into their own window" is enabled - hosts a
// single SchematicsThumbnailGallery bound to the Schematics tab's shared currentThumbnails
// collection. Size/state/position persistence mirrors ComponentInfoWindow's own constructor
// pattern (Handlers/Data/UserSettings.cs's SchematicsThumbnailsWindow* tier), minus the
// splitter-ratio/thumbnail-row-height fields that window has and this one has no equivalent of.
//
// With "Remember thumbnail window settings per board" also on, layout is read from and saved to
// UserSettings' per-board ThumbnailWindowSettingsByBoard tier instead of that single global set of
// fields - see Initialize's boardKey parameter. That decision is resolved once, at Initialize time
// (from Main.OpenSchematicsThumbnailsWindow, which only opens this window for the CURRENTLY
// selected board), not re-read afterwards - a window, once open, keeps saving to the board it was
// opened for even if both settings change while it is still showing.
// ###########################################################################################
public partial class SchematicsThumbnailsWindow : Window
{
    private double thisNormalWidth = 420.0;
    private double thisNormalHeight = 600.0;
    private int thisNormalX;
    private int thisNormalY;
    private TabSchematics? thisOwner;
    private string? thisPerBoardKey;

    public SchematicsThumbnailsWindow()
    {
        this.InitializeComponent();

        // F11 fullscreens the SCHEMATICS IMAGE (via Main.ToggleSchematicsFullscreenWindow, the
        // same toggle the main window's own F11 uses), not this window itself - this window is the
        // detached thumbnail gallery, and "fullscreen" has always meant the schematic. Tunnelled so
        // it is seen regardless of which child control has focus, matching how Main's own F11
        // handler is wired.
        this.AddHandler(KeyDownEvent, this.OnWindowKeyDownToggleSchematicsFullscreen, RoutingStrategies.Tunnel);

        // Belt and braces alongside the gallery's own OnDetachedFromVisualTree: a window that was
        // never shown - or closed before it laid out - never detaches, and the gallery's
        // subscription to the tab's application-lifetime thumbnail ListBox would then keep this
        // whole window alive for the rest of the session. Teardown is idempotent.
        this.Closed += (_, _) => this.Gallery.Teardown();
    }

    // ###########################################################################################
    // Wires the gallery to the Schematics tab's shared thumbnails collection and hosted (hidden)
    // ListBox, so selection and reorder stay in sync with the rest of the app. Also keeps the
    // owning tab for OnWindowKeyDownToggleSchematicsFullscreen, which needs its MainWindow, and
    // applies (and arms saving of) this window's remembered layout.
    //
    // boardKey is null for the ordinary single-global-layout case (every existing caller, and every
    // existing test's direct construction). Main.OpenSchematicsThumbnailsWindow passes the current
    // board's key instead when UserSettings.RememberThumbnailWindowSettingsPerBoard is on, which
    // switches both the initial restore AND the Closing-time save below to the per-board tier.
    // ###########################################################################################
    public void Initialize(
        ObservableCollection<SchematicThumbnail> thumbnails,
        ListBox hostedThumbnailList,
        TabSchematics owner,
        string? boardKey = null)
    {
        this.thisOwner = owner;
        this.thisPerBoardKey = boardKey;
        this.Gallery.Initialize(thumbnails, hostedThumbnailList, owner);

        bool hasLayout;
        string savedState;
        double savedWidth;
        double savedHeight;
        int savedX;
        int savedY;

        if (boardKey != null)
        {
            var perBoard = UserSettings.GetThumbnailWindowSettingsForBoard(boardKey);
            hasLayout = perBoard?.HasWindowLayout == true;
            savedState = perBoard?.WindowState ?? "Normal";
            savedWidth = perBoard?.WindowWidth ?? 420.0;
            savedHeight = perBoard?.WindowHeight ?? 600.0;
            savedX = perBoard?.WindowX ?? 0;
            savedY = perBoard?.WindowY ?? 0;
        }
        else
        {
            hasLayout = UserSettings.HasSchematicsThumbnailsWindowLayout;
            savedState = UserSettings.SchematicsThumbnailsWindowState;
            savedWidth = UserSettings.SchematicsThumbnailsWindowWidth;
            savedHeight = UserSettings.SchematicsThumbnailsWindowHeight;
            savedX = UserSettings.SchematicsThumbnailsWindowX;
            savedY = UserSettings.SchematicsThumbnailsWindowY;
        }

        this.thisNormalWidth = hasLayout ? savedWidth : 420.0;
        this.thisNormalHeight = hasLayout ? savedHeight : 600.0;
        this.thisNormalX = hasLayout ? savedX : 0;
        this.thisNormalY = hasLayout ? savedY : 0;

        if (hasLayout)
        {
            this.Width = savedWidth;
            this.Height = savedHeight;
            this.Position = new Avalonia.PixelPoint(this.thisNormalX, this.thisNormalY);

            if (string.Equals(savedState, "Maximized", StringComparison.OrdinalIgnoreCase))
                this.WindowState = WindowState.Maximized;
        }

        // Keep thisNormalWidth/thisNormalHeight up to date so they always reflect the last
        // non-maximized dimensions regardless of how the window is closed.
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
        };

        this.Closing += (_, _) =>
        {
            string state = this.WindowState == WindowState.Maximized ? "Maximized" : "Normal";

            if (this.thisPerBoardKey != null)
            {
                UserSettings.SaveThumbnailWindowLayoutForBoard(
                    this.thisPerBoardKey, state, this.thisNormalWidth, this.thisNormalHeight, this.thisNormalX, this.thisNormalY);
            }
            else
            {
                UserSettings.SaveSchematicsThumbnailsWindowLayout(
                    state, this.thisNormalWidth, this.thisNormalHeight, this.thisNormalX, this.thisNormalY);
            }
        };
    }

    // ###########################################################################################
    // F11 while this window is the active one fullscreens the schematic image, exactly as F11 does
    // from the main window - reached via the owning TabSchematics' own MainWindow reference since
    // this window has no direct link to Main. A no-op (not a throw) with no owner yet, which only
    // happens if F11 is somehow pressed before Initialize runs.
    // ###########################################################################################
    private void OnWindowKeyDownToggleSchematicsFullscreen(object? sender, KeyEventArgs e)
    {
        if (e.Key != Key.F11)
            return;

        this.thisOwner?.MainWindow?.ToggleSchematicsFullscreenWindow();
        e.Handled = true;
    }

    // ###########################################################################################
    // Lets Main re-parent this window to whichever window is currently the visible schematics
    // surface (itself, or the fullscreen window while one is open) - see Main.ReownSchematicsThumbnailsWindow's
    // header for why. WindowBase.Owner's setter is protected, so external code cannot assign it
    // directly; this is the one sanctioned way in from outside.
    // ###########################################################################################
    internal void SetOwnerWindow(Window owner) => this.Owner = owner;
}
