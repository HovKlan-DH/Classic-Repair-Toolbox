using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// PastedBitmapAlpha - a picture pasted from the Windows clipboard arrives from Avalonia labelled
// opaque with every alpha byte 0, and was drawn on the screen as nothing at all (owner report,
// 2026-10-15: a white box with only its corner handles). The bytes below are four-byte pixels,
// alpha last, as Avalonia's Bgra8888 and Rgba8888 both have it.
// ###########################################################################################
public sealed class PastedBitmapAlphaTests
{
    // Exactly what the reporting machine's clipboard gave: labelled opaque, alpha 0 everywhere.
    [Fact]
    public void A_bitmap_labelled_opaque_with_alpha_0_is_mended()
    {
        byte[] pixels = [10, 20, 30, 0, 40, 50, 60, 0];

        Assert.True(PastedBitmapAlpha.NeedsOpaqueAlpha(pixels, labelledOpaque: true));
    }

    // Skia's promise for an opaque bitmap is 255 everywhere - one pixel short is mended too.
    [Fact]
    public void A_bitmap_labelled_opaque_with_any_alpha_below_255_is_mended()
    {
        Assert.True(PastedBitmapAlpha.NeedsOpaqueAlpha([1, 2, 3, 255, 4, 5, 6, 254], labelledOpaque: true));
    }

    [Fact]
    public void A_bitmap_labelled_opaque_and_truly_opaque_is_left_alone()
    {
        Assert.False(PastedBitmapAlpha.NeedsOpaqueAlpha([1, 2, 3, 255, 4, 5, 6, 255], labelledOpaque: true));
    }

    // Not labelled opaque, but no alpha anywhere: a 32-bit DIB whose fourth byte nobody set. A
    // picture nobody can see is never what was pasted.
    [Fact]
    public void A_bitmap_with_no_alpha_anywhere_is_mended_whatever_its_label()
    {
        Assert.True(PastedBitmapAlpha.NeedsOpaqueAlpha([10, 20, 30, 0, 40, 50, 60, 0], labelledOpaque: false));
    }

    // A real picture with transparency - a logo on a clear background - keeps it.
    [Fact]
    public void A_picture_with_real_transparency_keeps_it()
    {
        Assert.False(PastedBitmapAlpha.NeedsOpaqueAlpha([0, 0, 0, 0, 40, 50, 60, 255, 0, 0, 0, 128], labelledOpaque: false));
    }

    [Fact]
    public void No_pixels_need_nothing()
    {
        Assert.False(PastedBitmapAlpha.NeedsOpaqueAlpha([], labelledOpaque: true));
        Assert.False(PastedBitmapAlpha.NeedsOpaqueAlpha([1, 2, 3], labelledOpaque: false));
    }

    [Fact]
    public void Mending_sets_every_alpha_to_255_and_leaves_the_colours()
    {
        byte[] pixels = [10, 20, 30, 0, 40, 50, 60, 7];

        PastedBitmapAlpha.SetOpaque(pixels);

        Assert.Equal([10, 20, 30, 255, 40, 50, 60, 255], pixels);
    }
}
