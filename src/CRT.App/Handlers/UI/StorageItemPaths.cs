using Avalonia.Platform.Storage;
using System.Collections.Generic;
using System.Linq;

namespace CRT
{
    // ###########################################################################################
    // The local paths of dropped or pasted storage items - the files CRT can read from this
    // computer. An item with no local file (a virtual item from an archive or a remote location)
    // is left out, as is nothing at all. One copy for every drop target that takes them all (code
    // review, 2026-10-16: the Feedback tab and BoardFilesWindow each had their own).
    // ###########################################################################################
    internal static class StorageItemPaths
    {
        public static List<string> LocalPaths(IEnumerable<IStorageItem>? items) =>
            items?.Select(item => item.TryGetLocalPath())
                .Where(path => !string.IsNullOrWhiteSpace(path))
                .Select(path => path!)
                .ToList() ?? new List<string>();
    }
}
