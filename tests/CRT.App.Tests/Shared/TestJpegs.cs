namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// JPEG headers built for a test - JpegExifOrientationTests reads them, and the UI tests use them as
// phone photos. Shared: compiled into CRT.App.Tests and CRT.App.UiTests alike.
// ###########################################################################################
internal static class TestJpegs
{
    // ###########################################################################################
    // A JPEG's header as a camera writes it: SOI, an APP1 "Exif" segment whose first IFD holds
    // another tag and then Orientation (0 leaves the segment out), the SOF0 frame header with
    // the size, and EOI - no image data, which neither this reader nor the size reader looks at.
    // ###########################################################################################
    internal static byte[] Jpeg(int width, int height, int orientation, bool bigEndian, ushort orientationType = 3)
    {
        var bytes = new List<byte> { 0xFF, 0xD8 };

        if (orientation != 0)
        {
            var tiff = new List<byte>();
            tiff.AddRange(bigEndian ? "MM"u8.ToArray() : "II"u8.ToArray());
            Add16(tiff, 42, bigEndian);
            Add32(tiff, 8, bigEndian);
            Add16(tiff, 2, bigEndian);

            // Make (ASCII, empty), then Orientation (SHORT, its value in the first two bytes).
            Add16(tiff, 0x010F, bigEndian); Add16(tiff, 2, bigEndian); Add32(tiff, 1, bigEndian); Add32(tiff, 0, bigEndian);
            Add16(tiff, 0x0112, bigEndian); Add16(tiff, orientationType, bigEndian); Add32(tiff, 1, bigEndian); Add16(tiff, orientation, bigEndian); Add16(tiff, 0, bigEndian);
            Add32(tiff, 0, bigEndian);

            byte[] payload = [.. "Exif\0\0"u8.ToArray(), .. tiff];
            int length = payload.Length + 2;
            bytes.AddRange([0xFF, 0xE1, (byte)(length >> 8), (byte)length]);
            bytes.AddRange(payload);
        }

        bytes.AddRange([0xFF, 0xC0, 0x00, 0x11, 0x08, (byte)(height >> 8), (byte)height, (byte)(width >> 8), (byte)width, 0x03, 1, 0x22, 0, 2, 0x11, 1, 3, 0x11, 1]);
        bytes.AddRange([0xFF, 0xD9]);

        return bytes.ToArray();
    }

    private static void Add16(List<byte> bytes, int value, bool bigEndian)
    {
        byte high = (byte)(value >> 8), low = (byte)value;
        bytes.AddRange(bigEndian ? [high, low] : [low, high]);
    }

    private static void Add32(List<byte> bytes, long value, bool bigEndian)
    {
        byte[] four = [(byte)(value >> 24), (byte)(value >> 16), (byte)(value >> 8), (byte)value];
        bytes.AddRange(bigEndian ? four : four.Reverse());
    }
}
