using System.Buffers.Binary;
using System.Text;
using Handlers.DataHandling;
using SkiaSharp;

namespace ClassicRepairToolbox.Tests.Collections;

// ###########################################################################################
// *** WHAT CRT WRITES, THE SERVER TAKES. *** A capture is saved by Avalonia's Bitmap.Save with PngBitmapEncoderOptions
// (TabOscilloscope.Capture.cs), which on Avalonia 12.2.0's Skia backend is exactly
// `pixmap.Encode(new SKPngEncoderOptions(SKPngEncoderFilterFlags.AllFilters, 6))` (Avalonia.Skia's ImageSavingHelper, read
// 2026-11-26) - and the server takes a collection's capture only as the plain PNG CollectionPng allows. Should Skia start
// writing a chunk of another kind, every share would be refused; this fails first, naming the chunks. Run on Skia itself, as
// the headless UI tests draw with stubs that encode nothing.
// ###########################################################################################
public sealed class CollectionPngEncoderTests
{
    private static SKBitmap Screen(SKColorSpace? colorSpace)
    {
        var bitmap = new SKBitmap(new SKImageInfo(800, 480, SKColorType.Bgra8888, SKAlphaType.Premul, colorSpace));

        for (int y = 0; y < 480; y += 3)
        {
            for (int x = 0; x < 800; x += 5)
                bitmap.SetPixel(x, y, new SKColor((byte)x, (byte)y, (byte)(x ^ y)));
        }

        return bitmap;
    }

    // As Avalonia's ImageSavingHelper encodes it.
    private static byte[] Encoded(SKBitmap bitmap)
    {
        using SKImage image = SKImage.FromBitmap(bitmap);
        using SKImage raster = image.ToRasterImage(true);
        using SKPixmap pixmap = raster.PeekPixels() ?? throw new InvalidOperationException("No pixels to encode.");
        using SKData data = pixmap.Encode(new SKPngEncoderOptions(SKPngEncoderFilterFlags.AllFilters, 6)) ?? throw new InvalidOperationException("Skia encoded nothing.");

        return data.ToArray();
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void A_screen_encoded_as_CRT_saves_it_is_a_capture_the_server_takes(bool sRgb)
    {
        using SKBitmap screen = Screen(sRgb ? SKColorSpace.CreateSrgb() : null);

        byte[] png = Encoded(screen);
        CollectionPng.Result result = CollectionPng.Check(png);

        Assert.True(result.IsCapture, $"{result.Problem} Its chunks: {ChunksOf(png)}");
        Assert.Equal((800, 480), (result.Width, result.Height));
    }

    // A scope's dump is often a BMP: decoded, then saved as PNG.
    [Fact]
    public void A_bmp_dump_decoded_and_saved_is_a_capture_the_server_takes()
    {
        using SKBitmap screen = Screen(SKColorSpace.CreateSrgb());
        using SKBitmap decoded = SKBitmap.Decode(Bmp(screen));

        byte[] png = Encoded(decoded);

        Assert.True(CollectionPng.Check(png).IsCapture, $"Its chunks: {ChunksOf(png)}");
    }

    // A 24-bit BMP of the screen, bottom row first - what a scope sends.
    private static byte[] Bmp(SKBitmap bitmap)
    {
        int rowBytes = (bitmap.Width * 3 + 3) & ~3;
        var bmp = new byte[54 + rowBytes * bitmap.Height];

        bmp[0] = (byte)'B';
        bmp[1] = (byte)'M';
        BinaryPrimitives.WriteInt32LittleEndian(bmp.AsSpan(2), bmp.Length);
        BinaryPrimitives.WriteInt32LittleEndian(bmp.AsSpan(10), 54);
        BinaryPrimitives.WriteInt32LittleEndian(bmp.AsSpan(14), 40);
        BinaryPrimitives.WriteInt32LittleEndian(bmp.AsSpan(18), bitmap.Width);
        BinaryPrimitives.WriteInt32LittleEndian(bmp.AsSpan(22), bitmap.Height);
        BinaryPrimitives.WriteInt16LittleEndian(bmp.AsSpan(26), 1);
        BinaryPrimitives.WriteInt16LittleEndian(bmp.AsSpan(28), 24);

        for (int y = 0; y < bitmap.Height; y++)
        {
            for (int x = 0; x < bitmap.Width; x++)
            {
                SKColor color = bitmap.GetPixel(x, bitmap.Height - 1 - y);
                int at = 54 + y * rowBytes + x * 3;

                (bmp[at], bmp[at + 1], bmp[at + 2]) = (color.Blue, color.Green, color.Red);
            }
        }

        return bmp;
    }

    // "IHDR 13, sRGB 1, IDAT 4096, IEND 0" - so a failure says what the encoder wrote.
    private static string ChunksOf(byte[] png)
    {
        var chunks = new List<string>();

        for (int at = 8; at + 8 <= png.Length;)
        {
            int length = BinaryPrimitives.ReadInt32BigEndian(png.AsSpan(at, 4));
            chunks.Add($"{Encoding.ASCII.GetString(png, at + 4, 4)} {length}");
            at += 12 + Math.Max(length, 0);
        }

        return string.Join(", ", chunks);
    }
}
