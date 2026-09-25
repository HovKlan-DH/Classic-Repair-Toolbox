using System.Text;
using Handlers.Oscilloscope;

namespace ClassicRepairToolbox.Tests;

// Characterisation tests for ScopePayloadParser - unwrapping the SCPI definite-length binary
// block a scope returns from a screen-dump, and reading the BMP header inside it.
//
// This is byte-offset code extracted from TabOscilloscope. It is exactly the kind of logic that
// fails silently on a malformed response, and it cannot be checked by looking at the UI.
public class ScopePayloadParserTests
{
    // Builds a well-formed definite-length block: '#', digit count, length, then the payload.
    private static byte[] Block(byte[] payload)
    {
        string length = payload.Length.ToString();
        var header = Encoding.ASCII.GetBytes("#" + length.Length + length);
        var buffer = new byte[header.Length + payload.Length];
        Buffer.BlockCopy(header, 0, buffer, 0, header.Length);
        Buffer.BlockCopy(payload, 0, buffer, header.Length, payload.Length);
        return buffer;
    }

    // -------------------------------------------------------------- TryExtractBinaryPayload

    [Fact]
    public void A_well_formed_block_yields_exactly_the_declared_payload()
    {
        var payload = new byte[] { 1, 2, 3, 4, 5 };

        Assert.True(ScopePayloadParser.TryExtractBinaryPayload(Block(payload), out byte[] extracted));
        Assert.Equal(payload, extracted);
    }

    // "#800001234" - eight length digits - is what most scopes actually send.
    [Fact]
    public void A_block_with_a_multi_digit_length_field_is_parsed()
    {
        var payload = Encoding.ASCII.GetBytes("BM-image-bytes");
        var raw = Encoding.ASCII.GetBytes("#8" + payload.Length.ToString("D8"));
        var buffer = new byte[raw.Length + payload.Length];
        Buffer.BlockCopy(raw, 0, buffer, 0, raw.Length);
        Buffer.BlockCopy(payload, 0, buffer, raw.Length, payload.Length);

        Assert.True(ScopePayloadParser.TryExtractBinaryPayload(buffer, out byte[] extracted));
        Assert.Equal(payload, extracted);
    }

    // Trailing bytes after the declared length (a terminator, say) are dropped, not returned.
    [Fact]
    public void Trailing_bytes_beyond_the_declared_length_are_discarded()
    {
        var block = Block(new byte[] { 9, 9, 9 });
        var withTerminator = block.Concat(new byte[] { (byte)'\n' }).ToArray();

        Assert.True(ScopePayloadParser.TryExtractBinaryPayload(withTerminator, out byte[] extracted));
        Assert.Equal(new byte[] { 9, 9, 9 }, extracted);
    }

    [Fact]
    public void A_zero_length_block_succeeds_with_an_empty_payload()
    {
        Assert.True(ScopePayloadParser.TryExtractBinaryPayload(Encoding.ASCII.GetBytes("#10"), out byte[] extracted));
        Assert.Empty(extracted);
    }

    // A truncated transfer must FAIL rather than hand back a short buffer that later gets decoded
    // as a corrupt image.
    [Fact]
    public void A_block_shorter_than_its_declared_length_is_rejected()
    {
        var truncated = Encoding.ASCII.GetBytes("#41000").Concat(new byte[] { 1, 2, 3 }).ToArray();

        Assert.False(ScopePayloadParser.TryExtractBinaryPayload(truncated, out byte[] extracted));
        Assert.Empty(extracted);
    }

    [Theory]
    [InlineData("")]                 // nothing at all
    [InlineData("#")]                // too short to hold a header
    [InlineData("ab")]               // too short, and no marker
    [InlineData("XYZ")]              // right length, wrong marker
    [InlineData("#0")]               // zero length-digits is not legal
    [InlineData("#X5")]              // length-digit count is not a digit
    public void A_malformed_block_is_rejected_and_returns_an_empty_payload(string raw)
    {
        Assert.False(ScopePayloadParser.TryExtractBinaryPayload(Encoding.ASCII.GetBytes(raw), out byte[] extracted));
        Assert.Empty(extracted);
    }

    // A response that is plain text (an error string, say) rather than a block is rejected on the
    // leading-'#' check, which is what stops an error message being saved as a .bmp.
    [Fact]
    public void A_plain_text_error_response_is_not_mistaken_for_a_block()
    {
        var text = Encoding.ASCII.GetBytes("-113,\"Undefined header\"");

        Assert.False(ScopePayloadParser.TryExtractBinaryPayload(text, out _));
    }

    // -------------------------------------------------------------- TryReadBmpMetadata

    // Builds a minimal 30-byte BMP head with width/height/bpp at offsets 18, 22 and 28,
    // then truncates to `size` so a short header can be tested too.
    private static byte[] Bmp(int width, int height, short bitsPerPixel, int size = 30)
    {
        var bytes = new byte[30];
        bytes[0] = (byte)'B';
        bytes[1] = (byte)'M';
        BitConverter.GetBytes(width).CopyTo(bytes, 18);
        BitConverter.GetBytes(height).CopyTo(bytes, 22);
        BitConverter.GetBytes(bitsPerPixel).CopyTo(bytes, 28);
        return bytes.Take(size).ToArray();
    }

    [Fact]
    public void A_valid_BMP_header_yields_its_width_height_and_bit_depth()
    {
        Assert.True(ScopePayloadParser.TryReadBmpMetadata(Bmp(800, 480, 24), out int w, out int h, out short bpp));

        Assert.Equal(800, w);
        Assert.Equal(480, h);
        Assert.Equal((short)24, bpp);
    }

    // A negative height is legal BMP - it means the rows are stored top-down. The parser reports
    // it as-is rather than taking the absolute value, so callers see the real header.
    [Fact]
    public void A_top_down_BMP_reports_its_negative_height_unchanged()
    {
        Assert.True(ScopePayloadParser.TryReadBmpMetadata(Bmp(640, -400, 32), out _, out int h, out _));

        Assert.Equal(-400, h);
    }

    // Exactly 30 bytes is the minimum accepted, because offset 28 needs two bytes to read.
    [Fact]
    public void A_header_of_exactly_thirty_bytes_is_accepted()
    {
        Assert.True(ScopePayloadParser.TryReadBmpMetadata(Bmp(1, 1, 1), out _, out _, out _));
    }

    [Fact]
    public void A_header_shorter_than_thirty_bytes_is_rejected()
    {
        Assert.False(ScopePayloadParser.TryReadBmpMetadata(Bmp(800, 480, 24, size: 29), out int w, out int h, out short bpp));

        Assert.Equal(0, w);
        Assert.Equal(0, h);
        Assert.Equal((short)0, bpp);
    }

    // A PNG dump (or any non-BMP format) fails the "BM" check, so the caller knows not to report
    // BMP dimensions for it.
    [Fact]
    public void A_payload_that_is_not_a_BMP_is_rejected()
    {
        var png = new byte[40];
        png[0] = 0x89;
        png[1] = (byte)'P';

        Assert.False(ScopePayloadParser.TryReadBmpMetadata(png, out _, out _, out _));
    }

    [Fact]
    public void An_empty_payload_is_rejected()
    {
        Assert.False(ScopePayloadParser.TryReadBmpMetadata(Array.Empty<byte>(), out _, out _, out _));
    }

    // -------------------------------------------------------------- the two together

    // The real flow: unwrap the SCPI block, then read the BMP inside it.
    [Fact]
    public void A_scope_screen_dump_unwraps_to_a_readable_BMP_header()
    {
        var raw = Block(Bmp(800, 480, 24));

        Assert.True(ScopePayloadParser.TryExtractBinaryPayload(raw, out byte[] payload));
        Assert.True(ScopePayloadParser.TryReadBmpMetadata(payload, out int w, out int h, out short bpp));

        Assert.Equal(800, w);
        Assert.Equal(480, h);
        Assert.Equal((short)24, bpp);
    }
}
