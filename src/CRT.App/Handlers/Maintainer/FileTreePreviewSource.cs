using System;
using System.Threading.Tasks;
using Handlers.DataHandling;

namespace Handlers.MaintainerHandling
{
    // ###########################################################################################
    // WHAT A FILE TREE'S HOST DOES WITH ONE OF ITS FILES (owner request, 2026-09-28) - reads it for
    // the hover card, and opens it in the program the computer uses for it. The tree (FileTreeView)
    // only draws; its host knows where a file's bytes are - see Main/FileTreeFiles.
    // ###########################################################################################
    public interface IFileTreeFiles
    {
        // The file's bytes for the hover card, or null when there are none. Never throws: a file
        // that cannot be read is shown as missing, never as a crash under the pointer.
        Task<byte[]?> ReadAsync(SystemFileEntry file);

        // Fetches the file and opens it. Null when it opened, otherwise the sentence to show.
        Task<string?> OpenAsync(SystemFileEntry file);
    }

    // ###########################################################################################
    // THE TABLE'S HOVER CARD, FOR ONE FILE OF A TREE (owner request, 2026-09-28: "the same hover
    // functionality per file, as in the table format, so I can view an image or open a PDF etc. It
    // should not show any text, if it has changed or whatever ... only the relative path"). So the
    // card is the table's BoardTableFilePreview with one side - the file as the tree opens it - and no
    // headline: its row already says new, changed or removed.
    //
    // Only a picture is read on hover. Anything else is a link, fetched when it is clicked.
    // ###########################################################################################
    public sealed class FileTreePreviewSource : CRT.IBoardTableFileSource
    {
        private readonly IFileTreeFiles thisFiles;
        private readonly SystemFileEntry thisFile;

        public FileTreePreviewSource(IFileTreeFiles files, SystemFileEntry file)
        {
            this.thisFiles = files ?? throw new ArgumentNullException(nameof(files));
            this.thisFile = file ?? throw new ArgumentNullException(nameof(file));
        }

        // One side, so neither is ever shown.
        public string PublishedLabel => string.Empty;

        public string CurrentLabel => string.Empty;

        public bool SaysUnchanged => false;

        // The card for this file: one side, its path. The path the card asks for is always this
        // file's - see Cell.
        public BoardTableFileCell Cell => new(PublishedPath: null, CurrentPath: this.thisFile.Path);

        public Task<byte[]?> ReadAsync(string path, CRT.BoardTableFileSide side) => this.thisFiles.ReadAsync(this.thisFile);

        public Task<string?> OpenAsync(string path, CRT.BoardTableFileSide side) => this.thisFiles.OpenAsync(this.thisFile);
    }
}
