using Avalonia.Controls;

namespace CRT
{
    // ###########################################################################################
    // What the Schematics tab shows while the schematic itself is in its fullscreen window: a note
    // on the left, and on the right the same auto-fit thumbnail gallery the tab's own column shows,
    // so another schematic can be picked (or the thumbnails reordered) without leaving fullscreen.
    // ###########################################################################################
    public partial class SchematicsFullscreenPlaceholder : UserControl
    {
        public SchematicsFullscreenPlaceholder()
        {
            this.InitializeComponent();
        }

        // ###########################################################################################
        // Sizes the two columns as the tab had them, and wires the gallery to the tab. The gallery
        // listens to the tab only while this stand-in is on screen - one made by F11 from another
        // tab may never be.
        // ###########################################################################################
        public void Initialize(TabSchematics owner, double ratio)
        {
            this.RootGrid.ColumnDefinitions[0].Width = new GridLength(ratio * 100.0, GridUnitType.Star);
            this.RootGrid.ColumnDefinitions[2].Width = new GridLength((1.0 - ratio) * 100.0, GridUnitType.Star);

            this.ThumbnailGalleryPlaceholder.Initialize(owner);
        }
    }
}
