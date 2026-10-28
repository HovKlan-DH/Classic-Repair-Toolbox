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

        // ---- marking from KiCad, on the Contribute tab (step 3, slice 3c) ---------------------------------

        public const string MarkAutomaticallyButton = "Mark all automatically";

        public const string MarkAutomaticallyTip =
            "Draws a rectangle around every component not marked on this picture yet, from the KiCad data it is aligned with";

        public const string MarkAutomaticallyNeedsAlignTip = "Needs the board's KiCad data, a CAD name for this picture, and KiCad aligned on it";

        public static string MarkedAutomatically(int count, string picture) => count switch
        {
            0 => $"Every component KiCad shows on [{Trimmed(picture)}] is marked already.",
            1 => $"Marked 1 component on [{Trimmed(picture)}] from the KiCad data.",
            _ => $"Marked {count.ToString(System.Globalization.CultureInfo.InvariantCulture)} components on [{Trimmed(picture)}] from the KiCad data.",
        };

        public static string NotAligned(string picture) =>
            $"Align KiCad on [{Trimmed(picture)}] first - the rectangles are drawn from that alignment.";

        public static string NotABoardView(string picture) =>
            $"[{Trimmed(picture)}] shows a KiCad schematic sheet, not the board - mark its components with Mark components.";

        public const string KiCadNotRead = "The board's KiCad data could not be read, so nothing was marked.";

        // Nothing to mark - told apart (code review, 2026-10-26: each read "marked already").
        public static string NoComponentsYet(string picture) =>
            $"The board has no components yet, so nothing was marked on [{Trimmed(picture)}] - add them first.";

        public static string NoFootprintOnPicture(string picture) =>
            $"None of the board's components has a KiCad footprint on [{Trimmed(picture)}], so nothing was marked.";

        // A picture tool's Done writes its picture's rectangles as it holds them - marks put in meanwhile
        // would go (code review, 2026-10-26).
        public const string PictureToolOn =
            "Mark components or Align KiCad is still on on the Schematics tab - finish it with Done or Cancel first.";

        // ---- the components not marked yet, beside the label editor (step 3, slice 3c) -------------------

        public static string UnmarkedHeading(int count) =>
            $"Not marked on this picture ({count.ToString(System.Globalization.CultureInfo.InvariantCulture)})";

        public const string UnmarkedHelp = "Pick one, then draw its rectangle - it is named for you.";

        public const string AllMarked = "Every component of the board is marked on this picture.";

        // ---- the bar ------------------------------------------------------------------------------------

        public const string Done = "Done";
        public const string Cancel = "Cancel";
        public const string Close = "Close";

        // While aligning, with another picture to align (step 3, slice 3d).
        public const string DoneNext = "Done - next picture";

        public static string DoneNextTip(string picture) => $"Saves this alignment, then aligns KiCad on [{Trimmed(picture)}]";

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
            EditComponent, EditComponentTip, MarkAutomaticallyButton, MarkAutomaticallyTip, MarkAutomaticallyNeedsAlignTip,
            MarkedAutomatically(0, "A"), MarkedAutomatically(1, "A"), MarkedAutomatically(2, "A"), NotAligned("A"), NotABoardView("A"),
            KiCadNotRead, UnmarkedHeading(3), UnmarkedHelp, AllMarked, NoComponentsYet("A"), NoFootprintOnPicture("A"), PictureToolOn,
            DoneNext, DoneNextTip("A"),
        ];
    }
}
