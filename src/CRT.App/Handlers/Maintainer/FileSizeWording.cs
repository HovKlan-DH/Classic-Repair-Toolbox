using System.Globalization;

namespace Handlers.MaintainerHandling
{
    // ###########################################################################################
    // A file's size in words - 812 bytes, 42.1 KB, 8.7 MB: 1024-based, one decimal, the same in every
    // locale. The one formatter for the Maintainer tab: the Unused files list and every file tree
    // (owner request, 2026-10-04: sizes "everywhere") say a size the same way.
    // ###########################################################################################
    public static class FileSizeWording
    {
        public static string Format(long bytes)
        {
            if (bytes < 1024)
                return bytes == 1 ? "1 byte" : $"{bytes.ToString(CultureInfo.InvariantCulture)} bytes";

            double kb = bytes / 1024.0;

            if (kb < 1024)
                return $"{kb.ToString("0.0", CultureInfo.InvariantCulture)} KB";

            return $"{(kb / 1024.0).ToString("0.0", CultureInfo.InvariantCulture)} MB";
        }
    }
}
