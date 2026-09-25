using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using CRT;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // Naming and validation for worklog entry photo/file attachments - the decisions made when a
    // user picks or drops a file, kept out of WorklogEntryEditorWindow so they can be tested
    // without a window, a file dialog or a drop event.
    //
    // Only the naming and vetting lives here; the actual byte copy is CopyAttachmentIntoFolder,
    // which is the single method in this class that touches the disk.
    // ###########################################################################################
    public static class WorklogAttachmentStorage
    {
        // ###########################################################################################
        // Why an attachment was refused. Distinguishing these lets the caller say what is actually
        // wrong instead of a single unhelpful "could not add photo".
        // ###########################################################################################
        public enum AttachmentProblem
        {
            None,
            NoFileSelected,
            FileNotFound,
            NotDisplayableImage,
            NotOpenableFile
        }

        // ###########################################################################################
        // Which list an attachment is destined for, which decides what file types it may be.
        //
        // Photos are drawn in-app with Avalonia's Bitmap, so they are limited to what Bitmap can
        // decode. Files are handed to the OS shell instead, so they are limited to what
        // ExternalTargetLauncher will open - a wider set that still excludes executables, scripts
        // and shortcuts, because the shell would RUN those rather than display them. An image is
        // valid as either; a PDF only as a File.
        // ###########################################################################################
        public enum AttachmentKind
        {
            Photo,
            File
        }

        // ###########################################################################################
        // Vets a path the user picked or dropped, before anything is copied. What counts as valid
        // depends on the kind:
        //
        //  - Photo: only what Avalonia's Bitmap can decode, since the app draws these itself
        //    (ContributionPackaging's set, reused rather than copied).
        //  - File: only what ExternalTargetLauncher will open, since these are handed to the OS
        //    shell - which RUNS an executable, script or shortcut rather than displaying it. That
        //    set is wider than the Photo one (a PDF is fine) and is the launcher's own, so the two
        //    cannot drift into a file that attaches but will not open.
        //
        // The file picker's own filter is only a suggestion - a name typed into its file box gets
        // straight past it, and a drag-and-drop never consults it at all - so the format is checked
        // here rather than trusted, exactly as ComponentContribution's picker does.
        // ###########################################################################################
        public static AttachmentProblem ValidateSourceFile(string? sourcePath, AttachmentKind kind)
        {
            string trimmed = sourcePath?.Trim() ?? string.Empty;

            if (string.IsNullOrWhiteSpace(trimmed))
            {
                return AttachmentProblem.NoFileSelected;
            }

            if (kind == AttachmentKind.Photo && !ContributionPackaging.IsDisplayableImageFile(trimmed))
            {
                return AttachmentProblem.NotDisplayableImage;
            }

            if (kind == AttachmentKind.File && !ExternalTargetLauncher.IsOpenableFile(trimmed))
            {
                return AttachmentProblem.NotOpenableFile;
            }

            try
            {
                if (!File.Exists(trimmed))
                {
                    return AttachmentProblem.FileNotFound;
                }
            }
            catch
            {
                // A malformed path (illegal characters, too long) throws rather than returning
                // false, and is just as unusable as a missing one.
                return AttachmentProblem.FileNotFound;
            }

            return AttachmentProblem.None;
        }

        // ###########################################################################################
        // The message shown for a refused attachment. Kept beside the enum so a new problem cannot
        // be added without a matching sentence.
        // ###########################################################################################
        public static string DescribeProblem(AttachmentProblem problem) => problem switch
        {
            AttachmentProblem.NoFileSelected => "Select a file first.",
            AttachmentProblem.FileNotFound => "That file could no longer be found.",
            AttachmentProblem.NotDisplayableImage =>
                "That file is not an image the application can display. Use PNG, JPG, GIF, BMP or WEBP.",
            AttachmentProblem.NotOpenableFile =>
                "That file type cannot be opened from the application. Use a document, image or data file - not a program, script or shortcut.",
            _ => string.Empty
        };

        // ###########################################################################################
        // The name an attachment is stored under: its list's prefix, an underscore, the owning
        // record's id, another underscore, then the original file name - "photo_3_IMG_1234.jpg"
        // for photo #3. (Attachments stored by earlier builds used "photo3_..." with no separating
        // underscore before the id; GetDisplayFileName still reads those back correctly, but no
        // attachment is ever written in that older form again - existing files on disk are not
        // renamed.)
        //
        // Attachments share one folder per entry, so two photos picked from different folders that
        // are both "IMG_1234.jpg" would otherwise collide and the second would overwrite the first
        // one's bytes while both rows pointed at the same file. Prefixing with the record id makes
        // the name unique by construction - ids are unique within their list - instead of hunting
        // for a free " (2)" variant, and it makes the file's owner readable straight off the name,
        // which is what lets a deleted photo's file be found and removed with confidence.
        //
        // Photos and Files number their ids independently, so photo #3 and file #3 would both want
        // "3_..."; the caller passes a prefix that separates them (see PhotoFilePrefix).
        //
        // The original name is kept rather than replaced by the id alone so the folder stays
        // readable to someone looking at it outside the app.
        // ###########################################################################################
        public static string BuildStoredFileName(string? sourcePath, string ownerPrefix, int recordId)
        {
            string candidate;

            try
            {
                candidate = Path.GetFileName(sourcePath?.Trim() ?? string.Empty);
            }
            catch
            {
                candidate = string.Empty;
            }

            if (string.IsNullOrWhiteSpace(candidate))
            {
                candidate = "attachment";
            }

            candidate = SanitizeFileName(candidate);

            return $"{ownerPrefix}_{recordId.ToString(CultureInfo.InvariantCulture)}_{candidate}";
        }

        // ###########################################################################################
        // The stored name with its owner prefix and id stripped back off - "photo_3_board.png" reads
        // as "board.png", "file_2_manual.pdf" as "manual.pdf". The prefix exists to keep names unique
        // on disk and is noise to the user, so lists show this instead.
        //
        // Only the prefix built for THIS record is removed, which is why the record's id is a
        // parameter. A file the user named "2_schematic.png" that became photo #1 is stored as
        // "photo_1_2_schematic.png" and correctly reads back as "2_schematic.png" - one segment
        // dropped, not two.
        //
        // A name that does not carry this record's exact prefix is returned unchanged. That covers
        // attachments recorded before the scheme existed, and the genuinely ambiguous case:
        // "file_2_backup.pdf" is both a plausible user-chosen name and exactly what file #2 would be
        // stored as, so only the record actually numbered 2 has it stripped.
        // ###########################################################################################
        public static string GetDisplayFileName(string? storedFileName, string ownerPrefix, int recordId)
        {
            string trimmed = storedFileName?.Trim() ?? string.Empty;
            if (string.IsNullOrWhiteSpace(trimmed))
            {
                return string.Empty;
            }

            // Strips only the exact prefix THIS record's name would have been built with, rather
            // than any prefix-plus-digits that happens to look right.
            //
            // Guessing from the string alone cannot work: "file_2_backup.pdf" is both a name a user
            // may legitimately have chosen and exactly what file #2 is stored as, and nothing in the
            // text distinguishes them. Matching the owning record's own id does: file #7 strips
            // "file_7_" and leaves "file_2_backup.pdf" untouched, so the only name ever removed is
            // one this class demonstrably created for that record.
            string id = recordId.ToString(CultureInfo.InvariantCulture);
            string expected = $"{ownerPrefix}_{id}_";

            if (TryStrip(trimmed, expected, out string displayName))
            {
                return displayName;
            }

            // The OLDER stored form, from before an underscore was added between the prefix and the
            // id ("photo3_board.png" rather than "photo_3_board.png") - see BuildStoredFileName's own
            // comment. Existing attachments are not renamed, so this class must keep reading that
            // form back correctly or every attachment stored before the change would show its raw
            // storage name forever the moment it upgrades.
            if (TryStrip(trimmed, $"{ownerPrefix}{id}_", out displayName))
            {
                return displayName;
            }

            // Photos attached by the build where PhotoFilePrefix was "" are stored bare, as
            // "3_board.png". Those names are still on disk and in entries.json, and without this
            // they would show their raw storage form forever - the exact prefix noise this method
            // exists to hide. The id still has to match the owning record, so the guarantee above
            // is unchanged: a user's own "3_notes.png" on photo #7 is left alone.
            if (ownerPrefix.Length > 0 && TryStrip(trimmed, $"{id}_", out displayName))
            {
                return displayName;
            }

            return trimmed;
        }

        // ###########################################################################################
        // Picks the id for a new attachment: one past the highest currently in the list, but never
        // one whose stored file is already sitting in the folder.
        //
        // Max(Id) + 1 alone is not safe, because it REUSES an id as soon as the highest-numbered
        // attachment is removed. Deleting attachment #2 and adding another gives #2 again, and the
        // stored name is built from the id - so the new file is named exactly what the old one was.
        // That matters when the old bytes are still there, which happens whenever the metadata save
        // that accompanied the delete failed: the record is gone from the list but the file is not,
        // and CopyAttachmentIntoFolder overwrites with overwrite: true. The user's new attachment
        // silently replaces an orphan, or two records end up sharing a prefix.
        //
        // Skipping ids whose file already exists closes that, and keeps BuildStoredFileName's
        // "unique by construction" claim actually true. existingFileNames is what the folder holds
        // right now; passing an empty set degrades to plain Max(Id) + 1.
        // ###########################################################################################
        public static int AllocateAttachmentId(
            IReadOnlyCollection<WorklogAttachmentRecord> attachments,
            string ownerPrefix,
            IReadOnlyCollection<string> existingFileNames)
        {
            int candidate = 1;
            if (attachments != null && attachments.Count > 0)
            {
                foreach (var attachment in attachments)
                {
                    if (attachment.Id >= candidate)
                    {
                        candidate = attachment.Id + 1;
                    }
                }
            }

            if (existingFileNames == null || existingFileNames.Count == 0)
            {
                return candidate;
            }

            // Compared on the "<prefix>_<id>_" stem rather than a whole filename, since the rest of
            // the name comes from whatever file the user is attaching and is not known here.
            //
            // EVERY stem GetDisplayFileName reads back has to be tested, not only the current one.
            // A file stored by an older build as "photo3_board.png" is just as much attachment #3's
            // as "photo_3_board.png" is, so handing id 3 out again because only the new stem was
            // checked reintroduces exactly the orphan-overwrite this scan exists to prevent - and
            // leaves the orphan and the live row rendering the same display name, since
            // GetDisplayFileName strips both forms.
            var taken = new HashSet<string>(existingFileNames, StringComparer.OrdinalIgnoreCase);

            while (taken.Any(name => IsStoredUnderId(name, ownerPrefix, candidate)))
            {
                candidate++;
            }

            return candidate;
        }

        // ###########################################################################################
        // Whether a stored file name belongs to the given record id - in ANY of the forms this class
        // has written over time, matching GetDisplayFileName's read-back exactly:
        //
        //  - "photo_3_board.png"  the current form
        //  - "photo3_board.png"   before the separating underscore was added
        //  - "3_board.png"        the build where PhotoFilePrefix was empty
        //
        // Kept beside AllocateAttachmentId because the two must agree: a form that reads back as
        // this record's must also block the id, or the same bytes are both displayed as one record's
        // and overwritten as another's.
        //
        // The bare form was only ever written for photos, so testing it for a FILE allocation can
        // only ever skip an id it did not have to - which costs nothing, and is the safe direction:
        // the alternative is handing out an id whose bytes are already on disk.
        // ###########################################################################################
        private static bool IsStoredUnderId(string name, string ownerPrefix, int recordId)
        {
            string id = recordId.ToString(CultureInfo.InvariantCulture);

            if (name.StartsWith($"{ownerPrefix}_{id}_", StringComparison.OrdinalIgnoreCase) ||
                name.StartsWith($"{ownerPrefix}{id}_", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            return ownerPrefix.Length > 0 && name.StartsWith($"{id}_", StringComparison.OrdinalIgnoreCase);
        }

        // ###########################################################################################
        // The file names currently in an attachments folder, or an empty list when the folder does
        // not exist or cannot be read. Feeds AllocateAttachmentId above; a folder that cannot be
        // listed simply degrades that to Max(Id) + 1 rather than failing the attach.
        // ###########################################################################################
        public static IReadOnlyCollection<string> ListAttachmentFileNames(string? folder)
        {
            if (string.IsNullOrWhiteSpace(folder) || !Directory.Exists(folder))
            {
                return Array.Empty<string>();
            }

            try
            {
                return Directory.GetFiles(folder).Select(Path.GetFileName).Where(n => n != null).Select(n => n!).ToArray();
            }
            catch (Exception ex)
            {
                Logger.Warning($"Failed to list worklog attachments folder [{folder}]: {ex.Message}");
                return Array.Empty<string>();
            }
        }

        // ###########################################################################################
        // Removes prefix from name when it is present and something is left afterwards. A name that
        // is nothing BUT the prefix is refused, since an empty display name is worse than the raw
        // one - the row would show no filename at all.
        // ###########################################################################################
        private static bool TryStrip(string name, string prefix, out string stripped)
        {
            stripped = string.Empty;

            if (!name.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            string candidate = name.Substring(prefix.Length);
            if (string.IsNullOrWhiteSpace(candidate))
            {
                return false;
            }

            stripped = candidate;
            return true;
        }

        // ###########################################################################################
        // Prefixes keeping the two attachment lists' ids from colliding in the folder they share.
        //
        // Both lists carry one, so a folder holding attachments of both kinds says which is which:
        // "photo_3_board.png" beside "file_2_manual.pdf". The prefix is what makes the name unique -
        // the two lists number their ids independently, so photo #3 and file #3 both exist and
        // would otherwise both want "3_board.png", one overwriting the other.
        //
        // Photos briefly used an empty prefix, when Files could not yet store anything and a bare
        // id read more cleanly. Now that both write into the folder, naming them symmetrically is
        // worth more than the shorter name.
        // ###########################################################################################
        public const string PhotoFilePrefix = "photo";
        public const string FileFilePrefix = "file";

        // ###########################################################################################
        // Strips anything the filesystem will not accept in a name. A dropped file's name comes from
        // wherever it was dragged from, so it is not assumed to be a legal name on this platform -
        // and the stored name is later combined into a path, so a name carrying a directory
        // separator would otherwise write outside the attachments folder.
        // ###########################################################################################
        private static string SanitizeFileName(string fileName)
        {
            var invalid = Path.GetInvalidFileNameChars();
            var builder = new System.Text.StringBuilder(fileName.Length);

            foreach (char character in fileName)
            {
                builder.Append(Array.IndexOf(invalid, character) >= 0 ? '_' : character);
            }

            string sanitized = builder.ToString().Trim();
            return string.IsNullOrWhiteSpace(sanitized) ? "attachment" : sanitized;
        }

        // ###########################################################################################
        // Renumbers DisplayOrder densely from 0, preserving the order the attachments already sort
        // into. Called when a list is loaded so that duplicate or gapped values - written by an
        // older build whose add path started at 1 while its reorder renumbered from 0, or by a
        // hand-edited entries.json - cannot leave two rows sharing an order and sorting arbitrarily.
        //
        // Returns true when anything actually changed, so the caller can avoid a pointless save.
        // ###########################################################################################
        public static bool NormalizeDisplayOrder(List<WorklogAttachmentRecord> attachments)
        {
            if (attachments == null || attachments.Count == 0)
            {
                return false;
            }

            var ordered = attachments.OrderBy(a => a.DisplayOrder).ToList();
            bool changed = false;

            for (int i = 0; i < ordered.Count; i++)
            {
                if (ordered[i].DisplayOrder != i)
                {
                    ordered[i].DisplayOrder = i;
                    changed = true;
                }
            }

            return changed;
        }

        // ###########################################################################################
        // Reorders attachments by moving the dragged one to sit at targetIndex in display order,
        // renumbering DisplayOrder from 0 so the result is a dense, gap-free sequence.
        //
        // targetIndex is the position the row should end up at in the list as the user sees it. It
        // is clamped rather than rejected, because a drag released past the end of the list means
        // "put it last", which is what the user is asking for and not an error.
        //
        // targetIndex is interpreted as the final position in the reordered list, so it is used
        // directly after the removal rather than adjusted: dragging row 0 to index 2 of four leaves
        // [1,2,3], and inserting at 2 puts the row third, which is where it was dropped.
        // ###########################################################################################
        public static void ReorderAttachment(List<WorklogAttachmentRecord> attachments, int id, int targetIndex)
        {
            if (attachments == null || attachments.Count < 2)
            {
                return;
            }

            var ordered = attachments.OrderBy(a => a.DisplayOrder).ToList();

            int currentIndex = ordered.FindIndex(a => a.Id == id);
            if (currentIndex < 0)
            {
                return;
            }

            int clampedTarget = Math.Clamp(targetIndex, 0, ordered.Count - 1);
            if (clampedTarget == currentIndex)
            {
                return;
            }

            var moved = ordered[currentIndex];
            ordered.RemoveAt(currentIndex);
            ordered.Insert(Math.Clamp(clampedTarget, 0, ordered.Count), moved);

            for (int i = 0; i < ordered.Count; i++)
            {
                ordered[i].DisplayOrder = i;
            }
        }

        // ###########################################################################################
        // Deletes one attachment's bytes, reporting whether the folder no longer holds that file.
        // An already-missing file counts as success: the caller's goal is that it is gone, and a
        // photo whose file was removed outside the app must still be deletable from the list.
        // ###########################################################################################
        public static bool DeleteAttachmentFile(string? attachmentsFolder, string? fileName)
        {
            if (string.IsNullOrWhiteSpace(attachmentsFolder) || string.IsNullOrWhiteSpace(fileName))
            {
                return false;
            }

            try
            {
                string path = Path.Combine(attachmentsFolder, fileName);

                // Guards against a stored name that escapes the folder (a "..\" that predates the
                // sanitizing above, or a hand-edited entries.json) - deleting outside the
                // attachments folder is never intended here.
                string fullPath = Path.GetFullPath(path);
                string fullFolder = Path.GetFullPath(attachmentsFolder);

                if (!fullPath.StartsWith(fullFolder + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase))
                {
                    Logger.Warning($"Refused to delete worklog attachment outside its folder: [{fileName}]");
                    return false;
                }

                if (File.Exists(fullPath))
                {
                    File.Delete(fullPath);
                }

                return true;
            }
            catch (Exception ex)
            {
                Logger.Warning($"Failed to delete worklog attachment [{fileName}]: {ex.Message}");
                return false;
            }
        }

        // ###########################################################################################
        // Deletes one attachment's bytes via DeleteAttachmentFile, then removes the entry's shared
        // "worklog_{id}" folder too if that was the last thing in it - so removing an entry's
        // only photo or file leaves no empty folder behind on disk.
        //
        // Photos and Files share ONE folder per entry (see WorklogManager.GetEntryAttachmentsFolder),
        // so "empty" means empty of BOTH kinds, not just the one just deleted - checking Directory
        // .EnumerateFileSystemEntries after the delete answers that directly, without either caller
        // having to know about the other's rows.
        //
        // Only for the "this attachment is gone for good" callers (a user's Delete, or an add whose
        // save failed and is being undone) - NOT for TryReplaceAttachmentFile's swap, which always
        // leaves a fresh file behind and has no reason to call this.
        // ###########################################################################################
        public static bool DeleteAttachmentFileAndFolderIfEmpty(string? attachmentsFolder, string? fileName)
        {
            bool deleted = DeleteAttachmentFile(attachmentsFolder, fileName);

            if (string.IsNullOrWhiteSpace(attachmentsFolder))
            {
                return deleted;
            }

            try
            {
                if (Directory.Exists(attachmentsFolder) && !Directory.EnumerateFileSystemEntries(attachmentsFolder).Any())
                {
                    Directory.Delete(attachmentsFolder);
                }
            }
            catch (Exception ex)
            {
                // The attachment itself is already gone either way - an empty folder left behind on
                // a failure here is untidy, not incorrect, so this does not change the return value.
                Logger.Warning($"Failed to remove empty worklog attachments folder [{attachmentsFolder}]: {ex.Message}");
            }

            return deleted;
        }

        // ###########################################################################################
        // Copies a vetted source file into the entry's attachments folder under storedFileName, and
        // reports whether the bytes actually landed. Returns false rather than throwing when the
        // copy fails (source vanished between picking and copying, destination not writable, disk
        // full), so the caller can refuse to add a row that would point at a file which is not
        // there - a row whose image can never load is worse than a refused add.
        //
        // Overwrites deliberately: stored names are built from the owning record's id, so replacing
        // a photo's image reuses that photo's existing name on purpose and the old bytes are meant
        // to go. There is no accidental-collision case left for a refusal to protect against.
        // ###########################################################################################
        public static bool CopyAttachmentIntoFolder(string sourcePath, string attachmentsFolder, string storedFileName)
        {
            try
            {
                Directory.CreateDirectory(attachmentsFolder);
                File.Copy(sourcePath, Path.Combine(attachmentsFolder, storedFileName), overwrite: true);
                return true;
            }
            catch (Exception ex)
            {
                Logger.Warning($"Failed to copy worklog attachment [{storedFileName}]: {ex.Message}");
                return false;
            }
        }

        // ###########################################################################################
        // Deletes the file an attachment was sourced FROM, once its bytes have safely landed in the
        // entry's attachments folder - turning what CopyAttachmentIntoFolder did into a MOVE.
        //
        // Deliberately a separate step rather than File.Move inside the copy. The attach sequence
        // can still fail after the bytes land (WorklogAttachmentWriter rolls the copy back out when
        // the metadata persist fails), and a move would already have destroyed the original by
        // then - so a full disk would cost the user the file itself rather than just the filing.
        // Copy first, delete only once nothing can still be rolled back.
        //
        // Reports whether the source is gone. An ALREADY-missing file counts as success: the goal
        // is that it is no longer there, and a source removed by hand between the copy and this
        // call has met it. A genuine failure (the file still open elsewhere, a read-only folder) is
        // logged and reported false but is NOT an error the user needs to act on - the attachment
        // itself is safely filed, and what is left behind is a duplicate, not a loss.
        // ###########################################################################################
        public static bool DeleteSourceFileAfterAttach(string? sourcePath)
        {
            if (string.IsNullOrWhiteSpace(sourcePath))
            {
                return false;
            }

            try
            {
                if (!File.Exists(sourcePath))
                {
                    return true;
                }

                File.Delete(sourcePath);
                return true;
            }
            catch (Exception ex)
            {
                Logger.Warning($"Attached worklog file [{sourcePath}] could not be removed from its original folder: {ex.Message} - the copy in the worklog is unaffected");
                return false;
            }
        }

        // ###########################################################################################
        // Swaps an attachment's file for a newly chosen one, leaving exactly one file behind, and
        // returns the name to store on the record (unchanged when the copy failed).
        //
        // The two cases look the same to the user but are not on disk. Stored names are built from
        // the record's id, so re-picking a file with the same name and extension produces the same
        // stored name and the copy overwrites in place - deleting "the old file" afterwards would
        // delete the one just written. Change the extension or pick a differently-named file and the
        // new name differs, so the previous file has to be removed or it sits there orphaned,
        // invisible to the app but taking up space forever.
        //
        // Ordering matters: the old file is only removed once the replacement has actually landed.
        // A failed copy leaves the original untouched and the record still pointing at it, rather
        // than a row whose image can never load again.
        // ###########################################################################################
        public static bool TryReplaceAttachmentFile(
            string sourcePath,
            string attachmentsFolder,
            string previousFileName,
            string newStoredFileName,
            out string storedFileName)
        {
            storedFileName = previousFileName;

            if (!CopyAttachmentIntoFolder(sourcePath, attachmentsFolder, newStoredFileName))
            {
                return false;
            }

            storedFileName = newStoredFileName;

            if (!string.Equals(previousFileName, newStoredFileName, StringComparison.OrdinalIgnoreCase))
            {
                DeleteAttachmentFile(attachmentsFolder, previousFileName);
            }

            return true;
        }
    }
}
