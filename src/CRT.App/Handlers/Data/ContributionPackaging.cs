using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // One file-backed contribution row that may carry an attachable source file.
    // SectionFolder is the zip sub-folder for the row's section (e.g. "ComponentImages"), and
    // ResolvedSourcePath is the verified on-disk path of the file, or null when the row's file
    // could not be resolved (nothing gets attached for it, and its zip entry stays empty).
    // ###########################################################################################
    public sealed class ContributionFileReference
    {
        public string SectionFolder { get; init; } = string.Empty;
        public string? ResolvedSourcePath { get; init; }
    }

    // ###########################################################################################
    // The result of planning the contribution zip: per input row the zip entry name its file
    // lives under (empty string when it has no attachable file), plus the distinct list of
    // files to actually write into the archive. Two rows referencing the same source file share
    // one zip entry, so Attachments can be shorter than EntryNames.
    // ###########################################################################################
    public sealed class ContributionZipPlan
    {
        public IReadOnlyList<string> EntryNames { get; init; } = Array.Empty<string>();
        public IReadOnlyList<ContributionAttachment> Attachments { get; init; } = Array.Empty<ContributionAttachment>();
    }

    public sealed class ContributionAttachment
    {
        public string SourcePath { get; init; } = string.Empty;
        public string ZipEntryName { get; init; } = string.Empty;
    }

    // ###########################################################################################
    // Pure packaging logic for the component contribution upload: resolving referenced files,
    // assigning deterministic zip entry names that the payload records per row (so the server
    // can locate each submitted file exactly instead of guessing by file name), and building
    // the plain-text summary sent alongside the zip.
    // ###########################################################################################
    public static class ContributionPackaging
    {
        public const string ReferencedFilesRootFolder = "ReferencedFiles";

        // Marker the contribution endpoint (Assets/Webserver/app-contribution/api/index.php) puts
        // in its response body when it rejects a submission because the application is older than
        // the newest released version. The token is followed by that newest version number.
        public const string OutdatedVersionToken = "OUTDATED_VERSION";

        // ###########################################################################################
        // Detects the server's "application too old" rejection in a contribution upload response.
        // Returns true when the token is present; newestVersion carries the version number the
        // server named right after the token, or an empty string when none could be read.
        // ###########################################################################################
        public static bool TryParseOutdatedVersionResponse(string? responseBody, out string newestVersion)
        {
            newestVersion = string.Empty;

            string body = responseBody?.Trim() ?? string.Empty;
            int tokenIndex = body.IndexOf(OutdatedVersionToken, StringComparison.OrdinalIgnoreCase);
            if (tokenIndex < 0)
            {
                return false;
            }

            string remainder = body.Substring(tokenIndex + OutdatedVersionToken.Length).Trim();
            string[] parts = remainder.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries);

            if (parts.Length > 0 && parts[0].Length > 0 && char.IsDigit(parts[0][0]))
            {
                newestVersion = parts[0];
            }

            return true;
        }

        // ###########################################################################################
        // The image formats the application can actually put on screen. Every component image is
        // drawn through Avalonia's Bitmap (Skia), so this is the set Bitmap can decode. It is
        // deliberately narrower than the ExternalTargetLauncher allowlist, which also permits ".svg"
        // because those files are handed to the OS shell rather than drawn by the app. A component
        // image outside this set uploads perfectly happily and then shows as an empty frame, so the
        // contribution editor refuses one up front instead.
        // ###########################################################################################
        public static readonly IReadOnlyList<string> DisplayableImageExtensions = new[]
        {
            ".png", ".jpg", ".jpeg", ".gif", ".bmp", ".webp"
        };

        private static readonly HashSet<string> DisplayableImageExtensionSet =
            new(DisplayableImageExtensions, StringComparer.OrdinalIgnoreCase);

        // ###########################################################################################
        // True when the given file name or path carries an extension the application can display as
        // a component image. Blank input, a name with no extension at all, and every non-image type
        // are false (fail closed) - the caller is deciding whether to accept a contributed file.
        // ###########################################################################################
        public static bool IsDisplayableImageFile(string? pathValue)
        {
            string trimmed = pathValue?.Trim() ?? string.Empty;
            if (string.IsNullOrWhiteSpace(trimmed))
            {
                return false;
            }

            try
            {
                return DisplayableImageExtensionSet.Contains(Path.GetExtension(trimmed));
            }
            catch
            {
                return false;
            }
        }

        // ###########################################################################################
        // What is wrong with the file on a component image row, if anything. A row is allowed to sit
        // there with no file while it is still being filled in, but by submission time every row must
        // carry a file the application can display - otherwise the contribution ships a component
        // image that renders as an empty frame for everybody who downloads the data.
        // ###########################################################################################
        public enum ComponentImageFileProblem
        {
            None,
            NoFileSelected,
            NotDisplayable
        }

        // ###########################################################################################
        // Judges the file value held by one component image row. storedPath is the row's own file
        // path as edited - absolute when it came from the file picker, relative to the data root when
        // it came from existing board data - and is not required to exist on disk here; this decides
        // whether the row is submittable at all, not whether the file resolves.
        // ###########################################################################################
        public static ComponentImageFileProblem ValidateComponentImageFile(string? storedPath)
        {
            if (string.IsNullOrWhiteSpace(storedPath))
            {
                return ComponentImageFileProblem.NoFileSelected;
            }

            return IsDisplayableImageFile(storedPath)
                ? ComponentImageFileProblem.None
                : ComponentImageFileProblem.NotDisplayable;
        }

        // ###########################################################################################
        // What stops a brand-new component from being submitted, if anything. Two fields have to be
        // right before a contributed component is of any use:
        //
        //   Board label - the component is identified by nothing else, and it must not name one the
        //                 board already has (see ValidateNewComponent for why a duplicate is refused
        //                 rather than merged).
        //   Category .. - the main window builds its category filter from the categories present in
        //                 the data and skips blank ones (ComponentListBuilder), so a component
        //                 without one is merged into the board data and then never reachable in the
        //                 UI at all. It is not an optional detail; it is what makes it visible.
        // ###########################################################################################
        public enum NewComponentProblem
        {
            None,
            BoardLabelMissing,
            BoardLabelAlreadyExists,
            CategoryMissing
        }

        // ###########################################################################################
        // Judges a new component against the board it is being added to. The board label is checked
        // first, because it is what the whole contribution is resolved by.
        //
        // A duplicate label is refused rather than merged: the server resolves a contribution by its
        // board label, so a submission reusing an existing one is diffed against that existing
        // component - and every image, file and link the new component does not happen to repeat
        // would come back as a proposed deletion of the existing one's data. Editing an existing
        // component is what the component list is for; this path only creates one that is not there.
        // Matching ignores case and surrounding whitespace, exactly as the board data is read.
        // ###########################################################################################
        public static NewComponentProblem ValidateNewComponent(
            string? boardLabel,
            string? category,
            IEnumerable<string>? existingBoardLabels)
        {
            string trimmedLabel = boardLabel?.Trim() ?? string.Empty;

            if (string.IsNullOrWhiteSpace(trimmedLabel))
            {
                return NewComponentProblem.BoardLabelMissing;
            }

            foreach (var existing in existingBoardLabels ?? Enumerable.Empty<string>())
            {
                string existingTrimmed = existing?.Trim() ?? string.Empty;
                if (string.IsNullOrWhiteSpace(existingTrimmed))
                {
                    continue;
                }

                if (string.Equals(existingTrimmed, trimmedLabel, StringComparison.OrdinalIgnoreCase))
                {
                    return NewComponentProblem.BoardLabelAlreadyExists;
                }
            }

            if (string.IsNullOrWhiteSpace(category))
            {
                return NewComponentProblem.CategoryMissing;
            }

            return NewComponentProblem.None;
        }

        // ###########################################################################################
        // The sentence shown once the contributor asks to delete a component: what goes with it.
        //
        // A component is not one row. Its images, schematic highlights, local files and links are
        // all resolved by its board label, so removing the component removes every one of them too -
        // and the contributor cannot see that while the sections are collapsed. Counting them out
        // loud is what turns "delete this component" from a guess into an informed decision.
        //
        // Sections with nothing in them are left out entirely rather than reported as "0 images":
        // the list exists to say what will be lost, and a zero is not a loss.
        // ###########################################################################################
        public static string BuildDeleteComponentSummary(
            string? boardLabel,
            int componentImages,
            int componentHighlights,
            int componentLocalFiles,
            int componentLinks)
        {
            string trimmedLabel = boardLabel?.Trim() ?? string.Empty;
            string subject = string.IsNullOrWhiteSpace(trimmedLabel)
                ? "This component"
                : $"The component [{trimmedLabel}]";

            var parts = new List<string>();
            AppendCount(parts, componentImages, "component image", "component images");
            AppendCount(parts, componentHighlights, "schematic highlight", "schematic highlights");
            AppendCount(parts, componentLocalFiles, "local file", "local files");
            AppendCount(parts, componentLinks, "link", "links");

            if (parts.Count == 0)
            {
                return $"{subject} will be removed from the board data. It carries no images, highlights, files or links.";
            }

            return $"{subject} will be removed from the board data, together with its {JoinWithAnd(parts)}.";
        }

        // ###########################################################################################
        // Adds "<n> <thing>" to the list, or nothing at all when the count is zero or negative.
        // ###########################################################################################
        private static void AppendCount(List<string> parts, int count, string singular, string plural)
        {
            if (count <= 0)
            {
                return;
            }

            parts.Add($"{count} {(count == 1 ? singular : plural)}");
        }

        // ###########################################################################################
        // "a", "a and b", "a, b and c" - the last separator is "and" rather than a comma, because
        // this is read as a sentence rather than scanned as a list.
        // ###########################################################################################
        private static string JoinWithAnd(IReadOnlyList<string> parts)
        {
            if (parts.Count <= 1)
            {
                return parts.Count == 0 ? string.Empty : parts[0];
            }

            string head = string.Join(", ", parts.Take(parts.Count - 1));

            return $"{head} and {parts[parts.Count - 1]}";
        }

        // ###########################################################################################
        // Assigns a zip entry name to every reference that has a resolved source file.
        // Entry names are "ReferencedFiles/<SectionFolder>/<NNN>_<filename>" with a single global
        // running number, and references to the same source path (case-insensitive) reuse the
        // first entry instead of packing the file twice.
        // ###########################################################################################
        public static ContributionZipPlan AssignZipEntries(IReadOnlyList<ContributionFileReference> references)
        {
            var entryNames = new List<string>(references.Count);
            var attachments = new List<ContributionAttachment>();
            var entryBySourcePath = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            foreach (var reference in references)
            {
                string sourcePath = reference.ResolvedSourcePath?.Trim() ?? string.Empty;

                if (string.IsNullOrWhiteSpace(sourcePath))
                {
                    entryNames.Add(string.Empty);
                    continue;
                }

                if (entryBySourcePath.TryGetValue(sourcePath, out string? existingEntry))
                {
                    entryNames.Add(existingEntry);
                    continue;
                }

                int number = attachments.Count + 1;
                string entryName = $"{ReferencedFilesRootFolder}/{reference.SectionFolder}/{number:D3}_{Path.GetFileName(sourcePath)}";

                entryBySourcePath[sourcePath] = entryName;
                attachments.Add(new ContributionAttachment
                {
                    SourcePath = sourcePath,
                    ZipEntryName = entryName
                });
                entryNames.Add(entryName);
            }

            return new ContributionZipPlan
            {
                EntryNames = entryNames,
                Attachments = attachments
            };
        }

        // ###########################################################################################
        // The data-root-relative folder for a picked file, or null when the file sits OUTSIDE the
        // data root (in which case the caller keeps whatever folder the user chose in the drop-down).
        //
        // Extracted from TabContribute's ApplySelectedFilePathToRow, which tested containment with a
        // bare dir.StartsWith(dataRoot) - no trailing separator and no normalization - so a SIBLING
        // folder whose name merely begins with the root's own name passed. With a root of
        // "...\Classic-Repair-Toolbox\Data", picking a file from "...\Classic-Repair-Toolbox\Data-backup"
        // was reported as inside it, and GetRelativePath then produced "../Data-backup/Commodore":
        // a traversing relative path written into the row and shipped in the uploaded payload, where
        // the server merge concatenates it onto ITS data root (see the ".." rejection in the
        // webserver's normalizeRelativeWorkbookPath, the other half of this contract).
        //
        // Both sides are normalized with GetFullPath and the root is compared WITH a trailing
        // separator, which is what makes "Data-backup" fail against "Data\" - the same approach
        // ExternalTargetLauncher.TryResolveDataRootScopedFilePath already uses for the same reason.
        //
        // Returns an empty string (not null) for a file sitting directly IN the root: that is inside
        // it, with no sub-folder, and the caller stores an empty FileLocation for exactly that case.
        // Null means "not inside", which is a different answer and drives a different branch.
        // ###########################################################################################
        public static string? TryGetDataRootRelativeFolder(string? dataRoot, string? fileDirectory)
        {
            if (string.IsNullOrWhiteSpace(dataRoot) || string.IsNullOrWhiteSpace(fileDirectory))
            {
                return null;
            }

            try
            {
                string normalizedRoot = Path.GetFullPath(dataRoot);
                string normalizedDirectory = Path.GetFullPath(fileDirectory);

                var comparison = OperatingSystem.IsWindows()
                    ? StringComparison.OrdinalIgnoreCase
                    : StringComparison.Ordinal;

                // The directory IS the root: inside, but with no sub-folder to name.
                if (string.Equals(normalizedDirectory, normalizedRoot, comparison))
                {
                    return string.Empty;
                }

                string rootWithSeparator = normalizedRoot.EndsWith(Path.DirectorySeparatorChar)
                    ? normalizedRoot
                    : normalizedRoot + Path.DirectorySeparatorChar;

                if (!normalizedDirectory.StartsWith(rootWithSeparator, comparison))
                {
                    return null;
                }

                string relative = Path.GetRelativePath(normalizedRoot, normalizedDirectory);

                if (relative == "." || relative.Length == 0)
                {
                    return string.Empty;
                }

                return relative.Replace('\\', '/');
            }
            catch
            {
                // GetFullPath throws on genuinely malformed input; treat that as "not inside"
                // rather than letting a bad pick take down the row.
                return null;
            }
        }

        // ###########################################################################################
        // Resolves an edited file path so it can be verified for existence and attached.
        // Accepts both relative paths (resolved against the data root) and external absolute
        // paths chosen through the file picker. Returns null when the file does not exist.
        // ###########################################################################################
        public static string? ResolveExistingFilePath(string dataRoot, string? pathValue)
        {
            string trimmed = pathValue?.Trim() ?? string.Empty;
            if (string.IsNullOrWhiteSpace(trimmed))
            {
                return null;
            }

            try
            {
                string normalizedInput = trimmed.Replace('/', Path.DirectorySeparatorChar);

                // 1. If the user selected an absolute path via the file picker anywhere on their PC, allow it!
                if (Path.IsPathRooted(normalizedInput))
                {
                    string fullPath = Path.GetFullPath(normalizedInput);
                    return File.Exists(fullPath) ? fullPath : null;
                }

                // 2. If it's a relative path, assume it lives strictly inside the current data-root
                if (!string.IsNullOrWhiteSpace(dataRoot))
                {
                    string normalizedDataRoot = Path.GetFullPath(dataRoot);
                    string combinedPath = Path.GetFullPath(Path.Combine(normalizedDataRoot, normalizedInput));
                    return File.Exists(combinedPath) ? combinedPath : null;
                }

                return null;
            }
            catch
            {
                return null;
            }
        }

        // ###########################################################################################
        // Builds the plain-text summary sent alongside the zipped JSON payload. The server relies
        // on the "Mandatory change comment:" marker line when formatting the notification email,
        // so that exact wording is part of the upload contract.
        //
        // The request type is stated on EVERY submission rather than only on a deletion. A line
        // that appears only in the dangerous case is a line whose absence has to be noticed, and a
        // deletion arriving as a mail that reads exactly like an ordinary edit is the one mistake
        // this summary exists to prevent.
        // ###########################################################################################
        public const string DeleteRequestTypeText = "DELETE COMPONENT";
        public const string UpdateRequestTypeText = "Component update";

        public static string BuildFeedbackText(
            string hardwareName,
            string boardName,
            string componentDisplayText,
            string componentUuidV4,
            string region,
            string comment,
            bool isDeleteRequest = false)
        {
            var builder = new StringBuilder();
            builder.AppendLine($"Request type: {(isDeleteRequest ? DeleteRequestTypeText : UpdateRequestTypeText)}");
            builder.AppendLine($"Hardware: {hardwareName}");
            builder.AppendLine($"Board: {boardName}");
            builder.AppendLine($"Component: {componentDisplayText}");
            builder.AppendLine($"Component UUID v4: {componentUuidV4}");
            builder.AppendLine($"Region context: {region}");
            builder.AppendLine();
            builder.AppendLine("Mandatory change comment:");
            builder.AppendLine(comment);

            return builder.ToString();
        }
    }
}
