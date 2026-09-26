using System;
using System.IO;
using System.Threading.Tasks;
using CRT;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // WHERE THE DRAFTS TAB'S TABLE READS A FILE CELL'S FILE for its hover card (owner request,
    // 2026-09-26) - see BoardTableEditor.FilePreview.cs.
    //
    // The published side is the local data folder, which holds what was synced - the same board
    // the table's colours compare against. The draft side is the draft's own copy when it has one,
    // else the published file: DraftFileResolver's draft-first rule, the one the board on screen
    // uses, so the card shows the picture the Schematics tab would.
    //
    // *** THE PATH IS WHATEVER WAS TYPED INTO THE CELL. *** So it is checked for shape
    // (SubmissionPathRules - no "..", nothing absolute) before it touches the disk, and the
    // published side is resolved through the same rule's containment check. Opening goes through
    // ExternalTargetLauncher, the only sanctioned way to hand a file to the operating system, with
    // the root the file was really found under.
    // ###########################################################################################
    internal sealed class DraftTableFileSource : IBoardTableFileSource
    {
        private readonly string thisDataRoot;
        private readonly string thisDraftSystemFolder;

        public DraftTableFileSource(string dataRoot, string draftSystemFolder)
        {
            this.thisDataRoot = dataRoot ?? string.Empty;
            this.thisDraftSystemFolder = draftSystemFolder ?? string.Empty;
        }

        public string PublishedLabel => "Published";

        public string CurrentLabel => "Your draft";

        public async Task<byte[]?> ReadAsync(string path, BoardTableFileSide side)
        {
            DraftFileResolution? file = this.Resolve(path, side);

            if (file is null)
            {
                return null;
            }

            try
            {
                return await File.ReadAllBytesAsync(file.FullPath);
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                return null;
            }
        }

        public Task<string?> OpenAsync(string path, BoardTableFileSide side)
        {
            DraftFileResolution? file = this.Resolve(path, side);

            string? problem = file is null
                ? "There is no file at this path."
                : ExternalTargetLauncher.TryOpen(file.FullPath, file.Root)
                    ? null
                    : "This file could not be opened.";

            return Task.FromResult(problem);
        }

        // ###########################################################################################
        // The file a side names, with the root it was found under - or null when the path is not a
        // safe relative path, or there is no such file.
        // ###########################################################################################
        internal DraftFileResolution? Resolve(string? path, BoardTableFileSide side)
        {
            string trimmed = path?.Trim() ?? string.Empty;

            if (!SubmissionPathRules.IsSafelyShaped(trimmed, out _))
            {
                return null;
            }

            if (side == BoardTableFileSide.Current)
            {
                return this.thisDraftSystemFolder.Length == 0
                    ? null
                    : DraftFileResolver.ResolveWithSource(this.thisDataRoot, this.thisDraftSystemFolder, trimmed);
            }

            if (this.thisDataRoot.Length == 0 ||
                !SubmissionPathRules.TryResolve(this.thisDataRoot, trimmed, out string fullPath, out _) ||
                !File.Exists(fullPath))
            {
                return null;
            }

            return new DraftFileResolution(fullPath, this.thisDataRoot, IsDrafted: false);
        }
    }
}
