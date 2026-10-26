namespace Handlers.DataHandling
{
    // The two tools the Contribute tab starts on a picture (CRT 4.0 step 2, slice 6 - the plan:
    // "Mark components and Align KiCad open the Schematics tab in that tool, with a bar and a Done
    // button"): the label editor, and the KiCad calibration.
    internal enum PictureTool
    {
        MarkComponents,
        AlignKiCad,
    }

    // Where the picture's image is: still being read, read, or not readable.
    internal enum PictureImage
    {
        Loading,
        Shown,
        Missing,
    }

    // The board's KiCad data on the Schematics tab: none at all, still being read, or read.
    internal enum BoardKiCad
    {
        None,
        Loading,
        Loaded,
    }

    internal enum PictureToolStep
    {
        // Not yet: the picture, or the board's KiCad data, is still being read.
        Wait,
        Start,

        // Never on this picture - `Reason` says why.
        Refuse,
    }

    // ###########################################################################################
    // WHEN A TOOL ASKED FOR ON A PICTURE CAN START (slice 6). The Schematics tab reads a picture's
    // image and the board's KiCad data in the background, so a tool asked for from the Contribute tab
    // waits for what it works on: marking needs the picture shown, aligning needs the board's KiCad
    // data too, and a view of it for this picture. What can never come is refused with the reason,
    // rather than waited for. Pure, so it is tested without a display.
    // ###########################################################################################
    internal sealed record PictureToolStart(PictureToolStep Step, string Reason = "")
    {
        public static PictureToolStart Decide(
            PictureTool tool,
            string picture,
            bool pictureOnBoard,
            bool pictureSelected,
            PictureImage image,
            BoardKiCad kiCad,
            bool pictureHasKiCadView)
        {
            if (!pictureOnBoard)
                return new(PictureToolStep.Refuse, PictureToolWording.PictureNotFound(picture));

            if (!pictureSelected || image == PictureImage.Loading)
                return new(PictureToolStep.Wait);

            if (image == PictureImage.Missing)
                return new(PictureToolStep.Refuse, PictureToolWording.ImageMissing(picture));

            if (tool == PictureTool.MarkComponents)
                return new(PictureToolStep.Start);

            return kiCad switch
            {
                BoardKiCad.None => new(PictureToolStep.Refuse, PictureToolWording.NoKiCadData),
                BoardKiCad.Loading => new(PictureToolStep.Wait),
                _ when !pictureHasKiCadView => new(PictureToolStep.Refuse, PictureToolWording.NoKiCadView(picture)),
                _ => new(PictureToolStep.Start),
            };
        }
    }
}
