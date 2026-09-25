using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // Decides whether a path in a submitted manifest may be written. PURE STRING WORK: it resolves
    // and compares, and never touches the filesystem, so every rule below is a unit test.
    //
    // *** EVERY PATH HERE IS UNTRUSTED INPUT. *** It arrives over the network from anyone with an
    // account and names a location the server will create a file at. This is the single most
    // security-sensitive function in Phase 4, and the trap list in NewContributeStrategy.md names
    // it explicitly: reject absolute paths, traversal, and anything resolving outside the target
    // system folder.
    //
    // THE CENTRAL RULE, AND WHY THE OBVIOUS IMPLEMENTATION IS WRONG: this RESOLVES the path and
    // then checks containment, rather than scanning for ".." or "/" patterns. Pattern-matching a
    // path is a losing game - "a/../../b", "a/./../../b", a symlinked component, a UNC prefix and
    // several encodings all defeat a naive scan, and the list of tricks grows. Resolving first and
    // asking "is the answer inside the folder" is the one formulation that cannot be worked
    // around, and it is the same reasoning OnlineServices.TryResolveValidatedLocalPath uses for
    // manifest entries.
    //
    // *** CASE IS PRESERVED, NEVER NORMALISED. *** From Phase 3 the data tree lives on Linux, where
    // CRT.Data's exact-case File.Exists means a row naming "foo.pdf" when the file is "foo.PDF"
    // loads on the contributor's Windows machine and fails only on the server. Mixed case is real
    // in the shipped tree (27 .PDF alongside 85 .pdf). Lower-casing paths here would hide
    // contributed data that is genuinely wrong, and would still leave Windows clients disagreeing
    // with the server about which file is meant. The strategy document says this outright: do NOT
    // "fix" it by making lookups case-insensitive.
    //
    // CONTAINMENT IS CHECKED CASE-SENSITIVELY for the same reason, because the server is the
    // filesystem that matters. On Windows that is stricter than the OS requires, which is the safe
    // direction to err.
    // ###########################################################################################
    public static class SubmissionPathRules
    {
        // A submitted path may not be longer than this. Well under any filesystem's limit, and
        // long enough for the deepest legitimate path in the shipped tree several times over.
        public const int MaximumPathLength = 200;

        // Windows reserves these as DEVICE names regardless of extension, so "CON.txt" is still
        // the console. A server on Linux would happily create such a file; a Windows CLIENT then
        // cannot sync it, and the failure appears on a machine that never saw the submission.
        // Rejecting them at the door keeps the tree portable, which is the whole point of a data
        // tree every platform syncs.
        private static readonly HashSet<string> ReservedWindowsNames = new(StringComparer.OrdinalIgnoreCase)
        {
            "CON", "PRN", "AUX", "NUL",
            "COM1", "COM2", "COM3", "COM4", "COM5", "COM6", "COM7", "COM8", "COM9",
            "LPT1", "LPT2", "LPT3", "LPT4", "LPT5", "LPT6", "LPT7", "LPT8", "LPT9"
        };

        // ###########################################################################################
        // Is this path safe to write inside the system folder?
        //
        // systemFolder is the absolute path of the folder this submission may write into.
        // relativePath is the untrusted value from the manifest.
        //
        // On success, resolvedPath is the absolute path to write - ALREADY RESOLVED, so the caller
        // must use this value rather than re-combining the inputs itself. A caller that re-joins
        // the original strings reintroduces exactly the hole this closes.
        // ###########################################################################################
        public static bool TryResolve(
            string systemFolder,
            string relativePath,
            out string resolvedPath,
            out string failureReason)
        {
            resolvedPath = string.Empty;
            failureReason = string.Empty;

            if (string.IsNullOrWhiteSpace(systemFolder))
            {
                failureReason = "The target system folder is not set.";
                return false;
            }

            if (string.IsNullOrWhiteSpace(relativePath))
            {
                failureReason = "A file path in the submission is empty.";
                return false;
            }

            if (relativePath.Length > SubmissionPathRules.MaximumPathLength)
            {
                failureReason =
                    $"The path is longer than {SubmissionPathRules.MaximumPathLength} characters: [{Trim(relativePath)}]";
                return false;
            }

            // A NUL byte truncates the path in any C library underneath, so "safe.txt\0../../evil"
            // can pass a managed containment check and then be written somewhere else entirely.
            // Other control characters have no legitimate place in a file name either.
            foreach (char character in relativePath)
            {
                if (char.IsControl(character))
                {
                    failureReason = $"The path contains a control character: [{Trim(relativePath)}]";
                    return false;
                }
            }

            // Backslash is a separator on Windows and a legal FILE NAME CHARACTER on Linux. A
            // client sending "a\..\..\b" would be one path on Windows and a single oddly-named
            // file on Linux - so the contract is forward slashes only, and anything else is
            // refused rather than translated. Translating would mean the same submission produced
            // different trees on different servers.
            if (relativePath.Contains('\\', StringComparison.Ordinal))
            {
                failureReason =
                    $"The path must use forward slashes only: [{Trim(relativePath)}]";
                return false;
            }

            if (Path.IsPathRooted(relativePath))
            {
                failureReason = $"The path must be relative, not absolute: [{Trim(relativePath)}]";
                return false;
            }

            // A leading "//" or a "C:" style prefix can survive IsPathRooted on a non-Windows
            // runtime, where the colon is an ordinary character. Checked explicitly so the answer
            // does not depend on which platform the server happens to run.
            if (relativePath.StartsWith('/') || relativePath.Contains(':', StringComparison.Ordinal))
            {
                failureReason = $"The path must be relative, not absolute: [{Trim(relativePath)}]";
                return false;
            }

            string[] segments = relativePath.Split('/');

            foreach (string segment in segments)
            {
                if (segment.Length == 0)
                {
                    failureReason = $"The path has an empty folder name: [{Trim(relativePath)}]";
                    return false;
                }

                if (segment == "." || segment == "..")
                {
                    failureReason = $"The path may not contain '.' or '..': [{Trim(relativePath)}]";
                    return false;
                }

                // A trailing dot or space is silently STRIPPED by Windows when creating a file, so
                // "report.pdf " becomes "report.pdf" - two manifest entries that differ only in a
                // trailing space would collide into one file, and the row referencing the other
                // would then name something that does not exist.
                if (segment.EndsWith('.') || segment.EndsWith(' ') || segment.StartsWith(' '))
                {
                    failureReason =
                        $"A path segment may not start or end with a space, or end with a dot: [{Trim(relativePath)}]";
                    return false;
                }

                string withoutExtension = Path.GetFileNameWithoutExtension(segment);

                if (SubmissionPathRules.ReservedWindowsNames.Contains(withoutExtension))
                {
                    failureReason =
                        $"[{segment}] is a reserved device name on Windows and cannot be used as a file name.";
                    return false;
                }
            }

            // Now resolve and check containment. Everything above is a fast refusal of input that
            // is obviously wrong; THIS is the check that actually guarantees the result lands
            // inside the folder.
            try
            {
                string root = Path.GetFullPath(systemFolder);
                string candidate = Path.GetFullPath(Path.Combine(root, relativePath.Replace('/', Path.DirectorySeparatorChar)));

                string rootWithSeparator = root.EndsWith(Path.DirectorySeparatorChar)
                    ? root
                    : root + Path.DirectorySeparatorChar;

                // Resolving to the folder itself is not a file path.
                if (string.Equals(candidate, root, StringComparison.Ordinal))
                {
                    failureReason = $"The path resolves to the system folder itself: [{Trim(relativePath)}]";
                    return false;
                }

                // Ordinal, not OrdinalIgnoreCase - see the header. The server's filesystem is the
                // one that matters, and it is case-sensitive.
                if (!candidate.StartsWith(rootWithSeparator, StringComparison.Ordinal))
                {
                    failureReason = $"The path escapes the system folder: [{Trim(relativePath)}]";
                    return false;
                }

                resolvedPath = candidate;
                return true;
            }
            catch (Exception ex) when (ex is ArgumentException or NotSupportedException or PathTooLongException)
            {
                failureReason = $"The path cannot be resolved: [{Trim(relativePath)}]";
                return false;
            }
        }

        // ###########################################################################################
        // Is this a well-formed SHA-256 as the manifest carries it?
        //
        // Lowercase hex, exactly 64 characters. Strict rather than forgiving because this value is
        // used as a STORE KEY: a hash accepted in two spellings would store the same blob twice
        // and, worse, let two manifest entries disagree about which one they meant.
        // ###########################################################################################
        public static bool IsValidHash(string? hash)
        {
            if (hash is null || hash.Length != SubmissionFormat.HashLength)
                return false;

            foreach (char character in hash)
            {
                bool isLowerHex = (character >= '0' && character <= '9')
                    || (character >= 'a' && character <= 'f');

                if (!isLowerHex)
                    return false;
            }

            return true;
        }

        // ###########################################################################################
        // Checks every path in a manifest, returning a finding per problem.
        //
        // Every problem at once rather than the first: a contributor whose client produced one bad
        // path probably produced several, and fixing them one round trip at a time is intolerable
        // over a slow upload.
        //
        // DUPLICATE PATHS ARE AN ERROR, compared case-SENSITIVELY for the reason in the header, but
        // ALSO reported when two paths differ only by case. The second is not a security problem;
        // it is a portability one - both files can exist on the server and only one can exist on a
        // contributor's Windows machine, so the tree would be un-syncable for half the audience.
        //
        // ###########################################################################################
        // *** `containmentRoot` IS THE DATA ROOT, NOT THE SYSTEM'S OWN FOLDER (renamed 2026-09-23).
        // ***
        //
        // A submitted path is data-root-relative ("Commodore/C64/250407/Sheet1.png"), so containing
        // it to the system folder was doubly wrong: it resolved every path one level too deep, and
        // it would have refused any SHARED file - "Commodore/Shared files/Component images/6526.png"
        // legitimately sits beside the manufacturer rather than inside one board.
        //
        // The parameter is renamed rather than just re-pointed so a caller cannot pass the old
        // thing and still compile. PublishPlan makes the identical resolve against the identical
        // base; the two must agree, or a submission validates against one location and publishes
        // to another.
        // ###########################################################################################
        public static IReadOnlyList<ValidationFinding> ValidateManifestPaths(
            SubmissionManifest manifest,
            string containmentRoot)
        {
            ArgumentNullException.ThrowIfNull(manifest);

            var findings = new List<ValidationFinding>();
            var seenExact = new HashSet<string>(StringComparer.Ordinal);
            var seenIgnoringCase = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            foreach (SubmissionFile file in manifest.Files)
            {
                if (!SubmissionPathRules.TryResolve(containmentRoot, file.Path, out _, out string failureReason))
                {
                    findings.Add(new ValidationFinding
                    {
                        Severity = ValidationSeverity.Error,
                        Code = "path.rejected",
                        Subject = file.Path,
                        Message = failureReason
                    });

                    continue;
                }

                if (!SubmissionPathRules.IsValidHash(file.Sha256))
                {
                    findings.Add(new ValidationFinding
                    {
                        Severity = ValidationSeverity.Error,
                        Code = "hash.malformed",
                        Subject = file.Path,
                        Message = "The file's SHA-256 is missing or is not 64 lowercase hex characters."
                    });
                }

                if (file.SizeBytes < 0 || file.SizeBytes > SubmissionFormat.MaximumBlobBytes)
                {
                    findings.Add(new ValidationFinding
                    {
                        Severity = ValidationSeverity.Error,
                        Code = "file.too_large",
                        Subject = file.Path,
                        Message =
                            $"The file is {file.SizeBytes} bytes, which is outside the permitted range " +
                            $"(up to {SubmissionFormat.MaximumBlobBytes} bytes)."
                    });
                }

                if (!seenExact.Add(file.Path))
                {
                    findings.Add(new ValidationFinding
                    {
                        Severity = ValidationSeverity.Error,
                        Code = "path.duplicate",
                        Subject = file.Path,
                        Message = "The same path appears more than once in the submission."
                    });

                    continue;
                }

                if (seenIgnoringCase.TryGetValue(file.Path, out string? other))
                {
                    findings.Add(new ValidationFinding
                    {
                        Severity = ValidationSeverity.Error,
                        Code = "path.case_collision",
                        Subject = file.Path,
                        Message =
                            $"This path differs from [{other}] only by capitalisation. Both can exist on " +
                            "the server but only one can exist on a Windows machine, so the data would " +
                            "not sync for everyone."
                    });
                }
                else
                {
                    seenIgnoringCase[file.Path] = file.Path;
                }
            }

            if (manifest.Files.Count > SubmissionFormat.MaximumFilesPerSubmission)
            {
                findings.Add(new ValidationFinding
                {
                    Severity = ValidationSeverity.Error,
                    Code = "submission.too_many_files",
                    Subject = string.Empty,
                    Message =
                        $"The submission lists {manifest.Files.Count} files, more than the limit of " +
                        $"{SubmissionFormat.MaximumFilesPerSubmission}."
                });
            }

            return findings;
        }

        // Keeps a hostile path from flooding a log line or a UI message. The value is still shown
        // because a contributor needs to know WHICH path was refused.
        private static string Trim(string value)
        {
            return value.Length <= 80 ? value : value[..80] + "...";
        }
    }
}
