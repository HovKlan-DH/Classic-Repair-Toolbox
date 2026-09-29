using System.Threading.Tasks;

namespace CRT
{
    // Which side of a file cell - the file as published, or as this draft or submission names it.
    public enum BoardTableFileSide
    {
        Published,
        Current
    }

    // ###########################################################################################
    // WHERE THE TABLE'S FILE PREVIEW GETS ITS BYTES (owner request, 2026-09-26) - the host's
    // business, since the two hosts keep files in different places: the Drafts tab reads the local
    // data folder and the draft's own copies, the Maintainer tab asks the server.
    //
    // WHICH file each side is, is CRT.Data's (BoardTableFileCells); what it looks like is
    // BoardTableFilePreview's. A host with no source set shows no preview at all.
    // ###########################################################################################
    public interface IBoardTableFileSource
    {
        // How the two sides are named in the preview - "Published" / "Your draft", or the maintainer
        // application's "Before (published)" / "After (submitted)".
        string PublishedLabel { get; }

        string CurrentLabel { get; }

        // Whether a file whose two sides hold the same bytes is headed "Unchanged". It is worth
        // saying where the cell could hide a file replaced under its own name; the maintainer
        // application says no for a NEW system, whose two sides are both the submission, so the
        // word said nothing there (owner request, 2026-09-26).
        bool SaysUnchanged => true;

        // The file's bytes, or null when there is no such file. Never throws: a file that cannot be
        // read is shown as missing, never as a crash under the pointer.
        Task<byte[]?> ReadAsync(string path, BoardTableFileSide side);

        // Opens the file outside the table, for a look at what cannot be drawn (a PDF). Returns why
        // it could not, or null when it opened.
        Task<string?> OpenAsync(string path, BoardTableFileSide side);
    }
}
