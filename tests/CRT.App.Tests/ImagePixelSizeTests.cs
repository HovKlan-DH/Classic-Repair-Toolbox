using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// ImagePixelSize - a picture's own pixel size from its header, for every type CRT shows (code
// review, 2026-10-26: only PNG and JPEG were read, so a GIF, BMP or WebP board picture had its
// marked components drawn at the size the preview decoded it to). The headers are built here byte
// by byte, as each format's specification lays them out.
// ###########################################################################################
public sealed class ImagePixelSizeTests
{
    private static byte[] Padded(params byte[] bytes)
    {
        var header = new byte[32];
        bytes.CopyTo(header, 0);
        return header;
    }

    [Fact]
    public void A_GIF_gives_its_logical_screen_size()
    {
        byte[] gif = Padded((byte)'G', (byte)'I', (byte)'F', (byte)'8', (byte)'9', (byte)'a', 0xA0, 0x0F, 0xB8, 0x0B);

        Assert.Equal((4000, 3000), ImagePixelSize.FromHeader(gif));
    }

    // The usual 40-byte header, and a picture stored top-down (a negative height).
    [Fact]
    public void A_BMP_gives_its_size_whichever_way_up_it_is_stored()
    {
        byte[] bmp = Padded((byte)'B', (byte)'M');
        BitConverter.GetBytes(40).CopyTo(bmp, 14);
        BitConverter.GetBytes(4000).CopyTo(bmp, 18);
        BitConverter.GetBytes(-3000).CopyTo(bmp, 22);

        Assert.Equal((4000, 3000), ImagePixelSize.FromHeader(bmp));

        byte[] core = Padded((byte)'B', (byte)'M');
        BitConverter.GetBytes(12).CopyTo(core, 14);
        BitConverter.GetBytes((ushort)640).CopyTo(core, 18);
        BitConverter.GetBytes((ushort)480).CopyTo(core, 20);

        Assert.Equal((640, 480), ImagePixelSize.FromHeader(core));
    }

    private static byte[] WebP(char kind)
    {
        byte[] webp = Padded((byte)'R', (byte)'I', (byte)'F', (byte)'F', 0, 0, 0, 0, (byte)'W', (byte)'E', (byte)'B', (byte)'P', (byte)'V', (byte)'P', (byte)'8', (byte)kind);
        return webp;
    }

    [Fact]
    public void A_WebP_gives_its_size_lossy_lossless_or_extended()
    {
        byte[] lossy = WebP(' ');
        lossy[23] = 0x9D;
        lossy[24] = 0x01;
        lossy[25] = 0x2A;
        BitConverter.GetBytes((ushort)4000).CopyTo(lossy, 26);
        BitConverter.GetBytes((ushort)3000).CopyTo(lossy, 28);
        Assert.Equal((4000, 3000), ImagePixelSize.FromHeader(lossy));

        // Lossless: 14 bits each, less one, packed little-endian after the 0x2F signature.
        byte[] lossless = WebP('L');
        lossless[20] = 0x2F;
        int bits = (4000 - 1) | ((3000 - 1) << 14);
        BitConverter.GetBytes(bits).CopyTo(lossless, 21);
        Assert.Equal((4000, 3000), ImagePixelSize.FromHeader(lossless));

        // Extended: the canvas, 24 bits each, less one.
        byte[] extended = WebP('X');
        extended[24] = (4000 - 1) & 0xFF;
        extended[25] = (4000 - 1) >> 8;
        extended[27] = (3000 - 1) & 0xFF;
        extended[28] = (3000 - 1) >> 8;
        Assert.Equal((4000, 3000), ImagePixelSize.FromHeader(extended));
    }

    // Anything else - or a header with no size in it - is not known, never guessed.
    [Fact]
    public void A_header_of_no_known_kind_or_with_no_size_gives_none()
    {
        Assert.Null(ImagePixelSize.FromHeader(Padded((byte)'%', (byte)'P', (byte)'D', (byte)'F')));
        Assert.Null(ImagePixelSize.FromHeader(Padded((byte)'G', (byte)'I', (byte)'F', (byte)'8', (byte)'9', (byte)'a')));
        Assert.Null(ImagePixelSize.FromHeader([(byte)'B', (byte)'M']));
        Assert.Null(ImagePixelSize.TryRead(null));
        Assert.Null(ImagePixelSize.TryRead(Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".gif")));
    }
}
