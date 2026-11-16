using Avalonia;
using Avalonia.Media.Imaging;
using Avalonia.Threading;
using Handlers.DataHandling;
using System;
using System.Threading.Tasks;

namespace CRT;

// ###########################################################################################
// The tab's pictures, read once and the last few kept (PictureReads, 2026-11-16, performance walkthrough): the selection
// takes a read made already or under way, and the board load starts reading the picture it will show while the
// thumbnails decode (StartReadingPicture). A picture let go of is disposed on the UI thread - never the one shown.
// ###########################################################################################
public partial class TabSchematics
{
    // The picture shown and the two before it: going back and forth between a board's top and bottom decodes neither,
    // and the board shown before keeps its picture. Some 50 MB each for a 4220x2941 schematic.
    internal const int PicturesKept = 3;

    private readonly PictureReads<Bitmap> thisPictureReads;

    private PictureReads<Bitmap> NewPictureReads() =>
        new(
            TabSchematics.PicturesKept,
            TabSchematics.ReadPicture,
            picture => Dispatcher.UIThread.Post(() =>
            {
                if (!ReferenceEquals(picture, this.currentFullResBitmap))
                {
                    picture.Dispose();
                }
            }));

    // ###########################################################################################
    // The picture at `path` begun reading now, off the UI thread, for the selection to take - the board load's, as its
    // thumbnails decode: the picture shown waited for every thumbnail until then. Safe from any thread.
    // ###########################################################################################
    internal void StartReadingPicture(string path)
    {
        if (!string.IsNullOrWhiteSpace(path))
        {
            _ = this.thisPictureReads.Read(path);
        }
    }

    // A picture decoded whole, with its own empty border measured (BlankBorderOf). Never throws.
    private static (Bitmap? Picture, Thickness Border) ReadPicture(string path)
    {
        try
        {
            var read = new Bitmap(path);
            return (read, TabSchematics.BlankBorderOf(read, path));
        }
        catch (Exception ex)
        {
            Logger.Warning($"Cannot load image file [{path}] - [{ex.Message}]");
            return (null, default);
        }
    }
}
