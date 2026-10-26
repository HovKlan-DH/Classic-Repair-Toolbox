using System.Collections.Generic;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // WHAT THE PICTURE TOOLS SAY (CRT 4.0 step 2, slice 6 - "Contributor mode goes"): the bar over
    // the Schematics tab while "Mark components" or "Align KiCad" is on, the buttons that start them
    // on the Contribute tab's Pictures, and the menu a right-click on a component opens for everyone.
    // Words only, so they are tested without a display.
    // ###########################################################################################
    internal static class PictureToolWording
    {
        // ---- starting a tool, on the Contribute tab ---------------------------------------------------

        public const string MarkButton = "Mark components";
        public const string AlignButton = "Align KiCad";

        public const string MarkTip = "Opens the Schematics tab on this picture, to draw a rectangle around each component";
        public const string AlignTip = "Opens the Schematics tab on this picture, to line the KiCad copper up with it";
        public const string AlignNeedsKiCadTip = "Needs the board's KiCad data and a CAD name for this picture";

        // ---- the bar ------------------------------------------------------------------------------------

        public const string Done = "Done";
        public const string Cancel = "Cancel";
        public const string Close = "Close";

        public static string Opening(string picture) => $"Opening [{Trimmed(picture)}]...";

        public static string Title(PictureTool tool, string picture) => tool == PictureTool.MarkComponents
            ? $"Mark components on [{Trimmed(picture)}]"
            : $"Align KiCad on [{Trimmed(picture)}]";

        public static string Help(PictureTool tool) => tool == PictureTool.MarkComponents
            ? "Drag on the picture to draw a component's rectangle. Drag a rectangle to move it or its edges to resize it, " +
              "right-click it to remove it, and press D to copy the one selected. Arrow keys nudge it; Ctrl+Z undoes. " +
              "Done saves it all into your draft."
            : "Drag the box, its edges and its corners until the KiCad copper lies on the picture. Arrow keys move it a " +
              "pixel - with Shift they grow it, with Alt they shrink it. Done saves it into your draft; Esc cancels.";

        public static string CannotStart(PictureTool tool, string picture) => tool == PictureTool.MarkComponents
            ? $"Components cannot be marked on [{Trimmed(picture)}]"
            : $"KiCad cannot be aligned on [{Trimmed(picture)}]";

        // ---- why a tool cannot start --------------------------------------------------------------------

        public static string PictureNotFound(string picture) =>
            $"[{Trimmed(picture)}] is not a picture of the board shown.";

        public static string ImageMissing(string picture) =>
            $"The image of [{Trimmed(picture)}] could not be read.";

        public const string NoKiCadData =
            "This board has no KiCad data to align - add it under KiCad data on the Contribute tab.";

        public static string NoKiCadView(string picture) =>
            $"[{Trimmed(picture)}] names no view of the board's KiCad data - give it a CAD name under Pictures on the Contribute tab.";

        // ---- the menu a right-click on a component opens --------------------------------------------------

        public static string Highlight(string label) => $"Highlight [{Trimmed(label)}]";

        public static string TakeHighlightAway(string label) => $"Take the highlight away from [{Trimmed(label)}]";

        public const string EditComponent = "Edit this component";

        public const string EditComponentTip = "Opens the component on the Contribute tab, in your draft of this board";

        private static string Trimmed(string? text) => text?.Trim() ?? string.Empty;

        // Every sentence above, for the test that they are plain ASCII.
        internal static IEnumerable<string> All() =>
        [
            MarkButton, AlignButton, MarkTip, AlignTip, AlignNeedsKiCadTip, Done, Cancel, Close, Opening("A"),
            Title(PictureTool.MarkComponents, "A"), Title(PictureTool.AlignKiCad, "A"), Help(PictureTool.MarkComponents),
            Help(PictureTool.AlignKiCad), CannotStart(PictureTool.MarkComponents, "A"), CannotStart(PictureTool.AlignKiCad, "A"),
            PictureNotFound("A"), ImageMissing("A"), NoKiCadData, NoKiCadView("A"), Highlight("A"), TakeHighlightAway("A"),
            EditComponent, EditComponentTip,
        ];
    }
}
