using System;
using System.IO;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // Resolves one BoardData entry's stored relative "File" value (ComponentImageEntry.File,
    // ComponentLocalFileEntry.File, BoardLocalFileEntry.File - all relative to a data root the same
    // way BoardSchematicEntry.SchematicImageFile is) to the absolute path its bytes actually live
    // at (NewContributeStrategy.md Phase 2, session 2c, task 8's component/file/link attachment
    // piece).
    //
    // Every existing consumer of one of these File values resolves it as a single
    // Path.Combine(DataManager.DataRoot, entry.File) - there is no "also check Drafts/" fallback
    // anywhere (confirmed by research before writing this: ComponentInfoWindow.axaml.cs x2,
    // TabOverview.axaml.cs, Main.BoardSelection.cs each reimplement the same one-root combine by
    // hand). A drafted attachment's bytes cannot live under Data/ - that tree is sync-owned and
    // freely overwritten - so a second root has to exist and every consumer has to know to try it.
    //
    // This is that second root, and the one place "where do a drafted attachment's bytes go" is
    // decided: ComponentDraftWriter copies a picked file's bytes here at save time (import-by-copy,
    // the same "copy into the app's own tree" pattern WorklogAttachmentWriter already uses for
    // worklog photos), under the SAME relative path the entry's own File value names - so nothing
    // about the stored entry needs to be draft-aware, only how it is opened does.
    //
    // Pure and testable: takes plain root paths in, does no DataManager/DraftManager lookups of its
    // own, and touches the filesystem only via File.Exists.
    // ###########################################################################################
    public static class DraftFileResolver
    {
        // The subfolder under a system's draft folder that holds copied attachment bytes, mirroring
        // how the file lived relative to its OWN root (Data/ or the draft) - so a file drafted as
        // "Datasheets/U8.pdf" is found at "<draftsRoot>/<system>/Files/Datasheets/U8.pdf".
        public const string DraftFilesFolderName = "Files";

        // ###########################################################################################
        // Resolves relativeFile (as stored in a BoardData entry's File property, "/"-separated) to
        // an absolute path, preferring the officially published copy under dataRoot and falling
        // back to a drafted copy under draftSystemFolder/Files/ when the official one is missing -
        // the ordinary case for a file that exists only because it was just attached in a draft and
        // has never been synced. Returns null when relativeFile is blank or neither copy exists, so
        // callers keep their existing "missing file" handling (skip the thumbnail, log a warning)
        // unchanged; this only adds a second place to look, not new required-file behaviour.
        // ###########################################################################################
        public static string? Resolve(string dataRoot, string draftSystemFolder, string? relativeFile) =>
            DraftFileResolver.ResolveWithSource(dataRoot, draftSystemFolder, relativeFile)?.FullPath;

        // ###########################################################################################
        // The same resolution as Resolve, plus WHICH ROOT the file was found under.
        //
        // *** A CALLER THAT OPENS THE FILE NEEDS THE ROOT, AND MUST NOT GUESS IT. ***
        // ExternalTargetLauncher only opens a local file that sits inside the root it is handed,
        // so a caller has to pass the root the path really came from. The component popup used to
        // work that out for itself with the old official-first assumption ("drafted only when the
        // official copy is missing"). After the order flipped to draft-first, a file present in
        // BOTH trees - every file a seeded draft copies - resolved to the draft path but was
        // tagged with the data root, the containment check refused it, and the datasheet button
        // silently did nothing on every board with a draft. Answering both halves here, from the
        // one place the order is decided, means the two can never disagree again.
        // ###########################################################################################
        public static DraftFileResolution? ResolveWithSource(string dataRoot, string draftSystemFolder, string? relativeFile)
        {
            string trimmed = relativeFile?.Trim() ?? string.Empty;
            if (trimmed.Length == 0)
            {
                return null;
            }

            string normalizedRelative = trimmed.Replace('/', Path.DirectorySeparatorChar);

            // ###########################################################################################
            // *** THE DRAFT IS TRIED FIRST, AND THE ORDER FLIPPED IN PHASE 6 (2026-09-23). ***
            //
            // It used to be official-first, drafted-second, because a draft was an OVERLAY: the
            // published bytes were the real file and the draft folder held only what had been newly
            // attached. Trying the draft second was correct then, and would be wrong now.
            //
            // A draft is a full copy of the board, so when one exists its files ARE the board's
            // files - including replacements for images that also exist in Data/. Official-first
            // would hand back the published bytes for an image the contributor has replaced, and
            // they would edit a board that kept showing the old picture.
            //
            // The published tree is still the fallback, which is what lets seeding skip SHARED
            // files (a manufacturer "Shared files" image belongs to many boards, so it is not
            // copied into any one draft - see DraftFolderLayout.GetReferencedFilePath).
            // ###########################################################################################
            if (!string.IsNullOrWhiteSpace(draftSystemFolder))
            {
                // *** THE SYSTEM'S OWN PREFIX IS STRIPPED. *** A stored path is relative to the
                // DATA ROOT ("Commodore/C64/250407/top.png"), while the draft folder already IS
                // those three segments - so combining them unstripped looks for
                // "<draft>/Commodore/C64/250407/Commodore/C64/250407/top.png".
                string? withinSystem = DraftFolderLayout.RelativeToSystemFolder(
                    DraftFileResolver.SystemKeyFromFolder(draftSystemFolder),
                    trimmed);

                if (withinSystem is not null)
                {
                    string draftedPath = Path.Combine(
                        draftSystemFolder,
                        withinSystem.Replace('/', Path.DirectorySeparatorChar));

                    if (File.Exists(draftedPath))
                    {
                        return new DraftFileResolution(draftedPath, draftSystemFolder, IsDrafted: true);
                    }
                }
            }

            if (!string.IsNullOrWhiteSpace(dataRoot))
            {
                string officialPath = Path.Combine(dataRoot, normalizedRelative);
                if (File.Exists(officialPath))
                {
                    return new DraftFileResolution(officialPath, dataRoot, IsDrafted: false);
                }
            }

            return null;
        }

        // ###########################################################################################
        // The absolute path a NEWLY attached file's bytes should be copied to, given the relative
        // path (FileLocation + file name, "/"-separated) the drafted entry is about to store. Does
        // not create the file or its directory - the caller copies the bytes and is the one place
        // that should decide when directory creation happens (see ComponentDraftWriter).
        // ###########################################################################################
        public static string BuildDraftFileDestination(string draftSystemFolder, string relativeFile)
        {
            string trimmed = relativeFile?.Trim() ?? string.Empty;

            // Same prefix-stripping as Resolve above, and for the same reason - a file written to
            // the unstripped path would be written somewhere Resolve never looks.
            string? withinSystem = DraftFolderLayout.RelativeToSystemFolder(
                DraftFileResolver.SystemKeyFromFolder(draftSystemFolder),
                trimmed);

            string relative = (withinSystem ?? trimmed).Replace('/', Path.DirectorySeparatorChar);

            return Path.Combine(draftSystemFolder, relative);
        }

        // ###########################################################################################
        // The system identity a draft folder represents, rebuilt from the folder's own last three
        // segments plus a placeholder file name.
        //
        // RelativeToSystemFolder takes an ExcelDataFile and ignores everything after the last "/",
        // so only the folder segments matter - and those are exactly what this folder path already
        // carries. This exists so the two methods above do not each have to be handed the identity
        // separately; every existing caller passes a folder, and changing that signature would
        // touch call sites this change has no other reason to visit.
        // ###########################################################################################
        private static string SystemKeyFromFolder(string draftSystemFolder)
        {
            string[] segments = (draftSystemFolder ?? string.Empty)
                .Split([Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar],
                    StringSplitOptions.RemoveEmptyEntries);

            if (segments.Length < 3)
            {
                return string.Empty;
            }

            return string.Join('/', segments[^3..]) + "/placeholder.xlsx";
        }
    }

    // ###########################################################################################
    // Where one stored File value was found: the absolute path, and the ROOT it sits under (the
    // draft system folder or the data root) - the containment root ExternalTargetLauncher must be
    // given to open it. IsDrafted says which of the two it was.
    // ###########################################################################################
    public sealed record DraftFileResolution(string FullPath, string Root, bool IsDrafted);
}
