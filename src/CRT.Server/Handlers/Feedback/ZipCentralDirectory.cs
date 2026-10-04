using System.Buffers.Binary;

namespace CRT.Server.Handlers.Feedback
{
    // ###########################################################################################
    // HOW MANY ENTRIES A ZIP'S CENTRAL DIRECTORY HOLDS, counted WITHOUT ZipArchive (code review,
    // 2026-10-04). ZipArchive builds an object for every central directory record the first time
    // its Entries are read, and only then compares the count with what the end record declared -
    // so a 250 MB zip that is nothing but empty records (~46 bytes each, five million of them) is
    // more than a gigabyte of objects before any limit of FeedbackFlow's could look at it. On the
    // anonymous feedback route that is a way to take the service down.
    //
    // This walks the records the way ZipArchive reads them - from the start the end record names,
    // for as long as each begins with a header's signature - and stops one past `limit`. Holding no
    // more than one record at a time, it costs a read of the directory and nothing else.
    //
    // A zip whose structure cannot be found is NOT refused here (null): ZipArchive then says what
    // is wrong with it, in the words the mail already uses.
    // ###########################################################################################
    public static class ZipCentralDirectory
    {
        private const uint EndSignature = 0x06054b50;
        private const uint Zip64LocatorSignature = 0x07064b50;
        private const uint Zip64EndSignature = 0x06064b50;
        private const uint HeaderSignature = 0x02014b50;

        private const int EndSize = 22;
        private const int Zip64LocatorSize = 20;
        private const int HeaderSize = 46;

        // The end record sits in the last 22 bytes plus a comment of up to 65,535.
        private const int EndSearchSize = ZipCentralDirectory.EndSize + ushort.MaxValue;

        // ###########################################################################################
        // The records counted, at most `limit + 1` (one past it says "too many"); null when the file
        // cannot be read or has no end record to start from.
        // ###########################################################################################
        public static long? CountRecords(string zipPath, long limit)
        {
            try
            {
                using var file = new FileStream(zipPath, FileMode.Open, FileAccess.Read, FileShare.Read, 64 * 1024);

                return ZipCentralDirectory.TryFindDirectoryStart(file, out long start)
                    ? ZipCentralDirectory.Count(file, start, limit)
                    : null;
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                return null;
            }
        }

        private static long Count(FileStream file, long start, long limit)
        {
            if (start < 0 || start >= file.Length)
                return 0;

            file.Seek(start, SeekOrigin.Begin);

            Span<byte> header = stackalloc byte[ZipCentralDirectory.HeaderSize];
            long count = 0;

            while (file.ReadAtLeast(header, header.Length, throwOnEndOfStream: false) == header.Length &&
                   BinaryPrimitives.ReadUInt32LittleEndian(header) == ZipCentralDirectory.HeaderSignature)
            {
                count++;

                if (count > limit)
                    return count;

                int variable =
                    BinaryPrimitives.ReadUInt16LittleEndian(header[28..]) +
                    BinaryPrimitives.ReadUInt16LittleEndian(header[30..]) +
                    BinaryPrimitives.ReadUInt16LittleEndian(header[32..]);

                file.Seek(variable, SeekOrigin.Current);
            }

            return count;
        }

        // ###########################################################################################
        // Where the central directory starts: the end record's offset, or the ZIP64 end record's
        // when the plain one is too small to hold it (0xFFFFFFFF) and a ZIP64 locator sits before it.
        // ###########################################################################################
        private static bool TryFindDirectoryStart(FileStream file, out long start)
        {
            start = 0;

            if (file.Length < ZipCentralDirectory.EndSize)
                return false;

            int size = (int)Math.Min(file.Length, ZipCentralDirectory.EndSearchSize);
            byte[] tail = new byte[size];

            file.Seek(file.Length - size, SeekOrigin.Begin);

            if (file.ReadAtLeast(tail, size, throwOnEndOfStream: false) != size)
                return false;

            // The LAST end signature, as ZipArchive looks for it - searching back from the end.
            int end = -1;

            for (int at = size - ZipCentralDirectory.EndSize; at >= 0; at--)
            {
                if (BinaryPrimitives.ReadUInt32LittleEndian(tail.AsSpan(at)) == ZipCentralDirectory.EndSignature)
                {
                    end = at;
                    break;
                }
            }

            if (end < 0)
                return false;

            uint offset = BinaryPrimitives.ReadUInt32LittleEndian(tail.AsSpan(end + 16));
            start = offset;

            if (offset != uint.MaxValue)
                return true;

            // ZIP64: the locator is the 20 bytes before the end record, and names the ZIP64 end
            // record, whose directory offset is at +48.
            long endInFile = file.Length - size + end;
            long locatorAt = endInFile - ZipCentralDirectory.Zip64LocatorSize;

            if (locatorAt < 0)
                return true;

            Span<byte> locator = stackalloc byte[ZipCentralDirectory.Zip64LocatorSize];
            file.Seek(locatorAt, SeekOrigin.Begin);

            if (file.ReadAtLeast(locator, locator.Length, throwOnEndOfStream: false) != locator.Length ||
                BinaryPrimitives.ReadUInt32LittleEndian(locator) != ZipCentralDirectory.Zip64LocatorSignature)
            {
                return true;
            }

            long zip64End = BinaryPrimitives.ReadInt64LittleEndian(locator[8..]);

            if (zip64End < 0 || zip64End > file.Length - 56)
                return true;

            Span<byte> record = stackalloc byte[56];
            file.Seek(zip64End, SeekOrigin.Begin);

            if (file.ReadAtLeast(record, record.Length, throwOnEndOfStream: false) == record.Length &&
                BinaryPrimitives.ReadUInt32LittleEndian(record) == ZipCentralDirectory.Zip64EndSignature)
            {
                start = BinaryPrimitives.ReadInt64LittleEndian(record[48..]);
            }

            return true;
        }
    }
}
