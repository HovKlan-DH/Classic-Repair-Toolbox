using System;

namespace Handlers.Oscilloscope
{
    // ###########################################################################################
    // Parsers for the raw bytes a scope returns from a screen-dump command: the SCPI definite
    // length binary block wrapper, and the BMP header inside it.
    //
    // Extracted from TabOscilloscope. This is exactly the kind of byte-offset code that is easy to
    // get subtly wrong and impossible to check by looking at the UI, so it lives here with tests.
    // ###########################################################################################
    public static class ScopePayloadParser
    {
        // ###########################################################################################
        // Extracts the payload bytes from a raw SCPI definite-length binary block response.
        //
        // The format is "#" then one digit giving how many digits the length itself occupies, then
        // that many length digits, then the payload: "#800001234" introduces 1234 payload bytes.
        // Anything shorter than its own declared length is rejected rather than truncated.
        // ###########################################################################################
        public static bool TryExtractBinaryPayload(byte[] rawData, out byte[] payload)
        {
            payload = Array.Empty<byte>();

            if (rawData.Length < 3 || rawData[0] != (byte)'#')
            {
                return false;
            }

            int lengthDigits = rawData[1] - (byte)'0';
            if (lengthDigits < 1 || rawData.Length < 2 + lengthDigits)
            {
                return false;
            }

            string lengthText = System.Text.Encoding.ASCII.GetString(rawData, 2, lengthDigits);
            if (!int.TryParse(lengthText, out int payloadLength) ||
                payloadLength < 0 ||
                rawData.Length < 2 + lengthDigits + payloadLength)
            {
                return false;
            }

            payload = new byte[payloadLength];
            Buffer.BlockCopy(rawData, 2 + lengthDigits, payload, 0, payloadLength);
            return true;
        }

        // ###########################################################################################
        // Reads basic BMP metadata from the dumped image payload when the format is BMP.
        //
        // Width, height and bit depth sit at fixed offsets 18, 22 and 28 of the BITMAPINFOHEADER.
        // A negative height is legal in BMP (top-down rows) and is returned as-is, not normalised.
        // ###########################################################################################
        public static bool TryReadBmpMetadata(byte[] imageBytes, out int width, out int height, out short bitsPerPixel)
        {
            width = 0;
            height = 0;
            bitsPerPixel = 0;

            if (imageBytes.Length < 30)
            {
                return false;
            }

            if (imageBytes[0] != (byte)'B' || imageBytes[1] != (byte)'M')
            {
                return false;
            }

            width = BitConverter.ToInt32(imageBytes, 18);
            height = BitConverter.ToInt32(imageBytes, 22);
            bitsPerPixel = BitConverter.ToInt16(imageBytes, 28);
            return true;
        }
    }
}
