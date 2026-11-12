using Avalonia.Input;

namespace Handlers.Geometry
{
    // ###########################################################################################
    // HOW A PICTURE IS MOVED, EVERYWHERE IN CRT - THE SCHEMATICS TAB'S WAY (owner request, 2026-11-06: "I do think we
    // need to unify how pan works, and then have the same pan functionality as present in the 'Schematics' tab"): a
    // drag with the RIGHT or the MIDDLE button pans, and the left button is left to what it does there - selecting,
    // drawing, moving a marker or a box. The Schematics tab (all its tools, aligning KiCad data and marking components
    // included - TabSchematics.TryBeginPan), the component popup's picture, the Contribute tab's previews of the
    // schematic images (PicturePreview) and the photo viewer (PhotoZoomView) ask this one rule. The wheel zooms, at
    // the pointer, as each has always had it.
    // ###########################################################################################
    public static class PanGesture
    {
        // What the photo viewer says about it.
        public const string Hint = "Drag with the right or middle mouse button to move the picture, and scroll to zoom.";

        // And the Schematics tab's one-time tip - the board's pictures are "schematic images" in every word a person reads
        // (owner request, 2026-11-11).
        public const string SchematicImageHint = "Drag with the right or middle mouse button to move the schematic image, and scroll to zoom.";

        // A press that starts a pan: the right or the middle button.
        public static bool IsPanPress(PointerPointProperties properties) =>
            properties.IsRightButtonPressed || properties.IsMiddleButtonPressed;
    }
}
