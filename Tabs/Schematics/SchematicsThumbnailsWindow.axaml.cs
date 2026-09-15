using Avalonia.Controls;
using Handlers.DataHandling;
using System;
using System.Collections.ObjectModel;
using Tabs.TabSchematics;

namespace CRT;

// ###########################################################################################
// The detached window shown when "Detach thumbnails to its own window" is enabled - hosts a
// single SchematicsThumbnailGallery bound to the Schematics tab's shared currentThumbnails
// collection. Size/state/position persistence mirrors ComponentInfoWindow's own constructor
// pattern (Handlers/Data/UserSettings.cs's SchematicsThumbnailsWindow* tier), minus the
// splitter-ratio/thumbnail-row-height fields that window has and this one has no equivalent of.
// ###########################################################################################
public partial class SchematicsThumbnailsWindow : Window
{
    private double thisNormalWidth = 420.0;
    private double thisNormalHeight = 600.0;
    private int thisNormalX;
    private int thisNormalY;

    public SchematicsThumbnailsWindow()
    {
        this.InitializeComponent();

        this.thisNormalWidth = UserSettings.HasSchematicsThumbnailsWindowLayout
            ? UserSettings.SchematicsThumbnailsWindowWidth
            : 420.0;

        this.thisNormalHeight = UserSettings.HasSchematicsThumbnailsWindowLayout
            ? UserSettings.SchematicsThumbnailsWindowHeight
            : 600.0;

        this.thisNormalX = UserSettings.HasSchematicsThumbnailsWindowLayout ? UserSettings.SchematicsThumbnailsWindowX : 0;
        this.thisNormalY = UserSettings.HasSchematicsThumbnailsWindowLayout ? UserSettings.SchematicsThumbnailsWindowY : 0;

        if (UserSettings.HasSchematicsThumbnailsWindowLayout)
        {
            this.Width = UserSettings.SchematicsThumbnailsWindowWidth;
            this.Height = UserSettings.SchematicsThumbnailsWindowHeight;
            this.Position = new Avalonia.PixelPoint(this.thisNormalX, this.thisNormalY);

            if (string.Equals(UserSettings.SchematicsThumbnailsWindowState, "Maximized", StringComparison.OrdinalIgnoreCase))
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

            UserSettings.SaveSchematicsThumbnailsWindowLayout(
                state, this.thisNormalWidth, this.thisNormalHeight, this.thisNormalX, this.thisNormalY);
        };

        // Belt and braces alongside the gallery's own OnDetachedFromVisualTree: a window that was
        // never shown - or closed before it laid out - never detaches, and the gallery's
        // subscription to the tab's application-lifetime thumbnail ListBox would then keep this
        // whole window alive for the rest of the session. Teardown is idempotent.
        this.Closed += (_, _) => this.Gallery.Teardown();
    }

    // ###########################################################################################
    // Wires the gallery to the Schematics tab's shared thumbnails collection and hosted (hidden)
    // ListBox, so selection and reorder stay in sync with the rest of the app.
    // ###########################################################################################
    public void Initialize(ObservableCollection<SchematicThumbnail> thumbnails, ListBox hostedThumbnailList, TabSchematics owner)
    {
        this.Gallery.Initialize(thumbnails, hostedThumbnailList, owner);
    }
}
