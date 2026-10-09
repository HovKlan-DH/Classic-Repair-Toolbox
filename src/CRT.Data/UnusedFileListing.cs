using System.Collections.Generic;
using System.Linq;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // The administrator's "Unused files" list for one data tree (2026-09-25): every file nothing
    // uses (DataTreeUsage), with its size - or, when the tree could not be read completely, why
    // none can be named. The server writes it and the Maintainer tab reads it as the same
    // record, so the two cannot drift apart.
    //
    // `PublicDataUrl` (2026-10-04) is where the tree is published, so the list - drawn as the same
    // folder tree as a board's Files - can show a file and open it, exactly as CRT downloads it.
    // Null from a server older than 4.4.0: the tree then opens nothing.
    // ###########################################################################################
    public sealed record UnusedFileListing(
        string Tree,
        bool IsComplete,
        IReadOnlyList<string> Problems,
        int MasterCount,
        int BoardWorkbookCount,
        int FileCount,
        IReadOnlyList<UnusedFileEntry> Files,
        string? PublicDataUrl = null)
    {
        public long TotalBytes => this.Files.Sum(file => file.SizeBytes);
    }

    public sealed record UnusedFileEntry(string Path, long SizeBytes);
}
