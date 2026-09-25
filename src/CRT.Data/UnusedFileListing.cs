using System.Collections.Generic;
using System.Linq;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // The administrator's "Unused files" list for one data tree (2026-09-25): every file nothing
    // uses (DataTreeUsage), with its size - or, when the tree could not be read completely, why
    // none can be named. The server writes it and the maintainer application reads it as the same
    // record, so the two cannot drift apart.
    // ###########################################################################################
    public sealed record UnusedFileListing(
        string Tree,
        bool IsComplete,
        IReadOnlyList<string> Problems,
        int MasterCount,
        int BoardWorkbookCount,
        int FileCount,
        IReadOnlyList<UnusedFileEntry> Files)
    {
        public long TotalBytes => this.Files.Sum(file => file.SizeBytes);
    }

    public sealed record UnusedFileEntry(string Path, long SizeBytes);
}
