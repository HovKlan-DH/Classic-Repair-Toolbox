using Handlers.DataHandling;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;

namespace CRT
{
    public static class ExternalTargetLauncher
    {
        // ###########################################################################################
        // File extensions the launcher will hand to the OS shell. TryStart uses ShellExecute, and
        // the shell RUNS executables, scripts and shortcuts rather than displaying them - so a
        // *.exe/*.bat/*.lnk inside the (network-synced, community-contributed) data root must never
        // become code execution just because a workbook cell references it. Only the document,
        // image and data formats that board data actually contains are openable; anything else,
        // including a file with no extension, is rejected (fail closed). Extend this set when board
        // data legitimately gains a new non-executable file type.
        // ###########################################################################################
        private static readonly HashSet<string> AllowedFileExtensions = new(StringComparer.OrdinalIgnoreCase)
        {
            // Images
            ".png", ".jpg", ".jpeg", ".gif", ".bmp", ".webp", ".svg",
            // Documents
            ".pdf", ".txt", ".md", ".html", ".htm", ".csv",
            // Data
            ".json", ".xml", ".xlsx", ".xls",
            // Domain-specific files shipped with board data (scope captures, CAD schematics)
            ".fsc", ".sch", ".kicad_pcb", ".kicad_sch"
        };

        // ###########################################################################################
        // The same set, exposed so callers that ATTACH a file can refuse an unopenable one up front
        // rather than storing it and failing at open time - the worklog's Files section builds its
        // picker filter and its validation from this.
        //
        // Deliberately derived from the launcher's own set rather than a second hand-kept list: a
        // caller keeping its own would drift, and the drift would show up as a file that attaches
        // happily and then cannot be opened. Extending AllowedFileExtensions extends this too.
        //
        // Returns a COPY. IReadOnlyCollection is only a compile-time promise - the backing HashSet
        // implements it, so handing back the instance itself would let any caller cast it and
        // Add(".exe"), permanently defeating the executable/script/shortcut rejection this class
        // exists to enforce. The set is tiny and callers use it once to build a picker filter, so
        // copying costs nothing next to leaving a security allowlist writable.
        // ###########################################################################################
        public static IReadOnlyCollection<string> OpenableFileExtensions => AllowedFileExtensions.ToArray();

        // ###########################################################################################
        // True when the file's extension is one the launcher will hand to the OS shell. Blank input
        // and a name with no extension are false (fail closed), matching TryOpen's own behaviour.
        //
        // The extension is read from the NORMALIZED full path, exactly as HasAllowedFileExtension
        // does for the open path. Reading it from the raw string instead lets the two disagree
        // about the same file: Windows quirks like a trailing dot or an alternate data stream
        // ("notes.txt:evil.exe") change what the OS finally resolves, so a file could pass the
        // attach-time check here and be refused - or worse, resolve differently - at open time.
        // Since the whole point of exposing this is that the two checks cannot drift, they have to
        // examine the same string.
        // ###########################################################################################
        public static bool IsOpenableFile(string? pathValue)
        {
            string trimmed = pathValue?.Trim() ?? string.Empty;
            if (string.IsNullOrWhiteSpace(trimmed))
            {
                return false;
            }

            try
            {
                // GetFullPath resolves the trailing-dot and ADS forms the raw string hides. It
                // throws on genuinely malformed input, which the catch below turns into a refusal.
                return ExternalTargetLauncher.HasAllowedFileExtension(Path.GetFullPath(trimmed));
            }
            catch
            {
                return false;
            }
        }

        // ###########################################################################################
        // Opens a validated external target. Allowed URI schemes are HTTP/HTTPS/mailto, and local
        // files must resolve inside the configured data-root boundary and carry an extension from
        // the document/image/data allowlist above - never an executable, script or shortcut.
        // ###########################################################################################
        public static bool TryOpen(string target, string? dataRootOverride = null)
        {
            if (string.IsNullOrWhiteSpace(target))
                return false;

            if (ExternalTargetLauncher.TryCreateAllowedUri(target, out Uri allowedUri))
            {
                return ExternalTargetLauncher.TryStart(allowedUri.AbsoluteUri, $"URI [{allowedUri.AbsoluteUri}]");
            }

            string dataRoot = !string.IsNullOrWhiteSpace(dataRootOverride)
                ? dataRootOverride
                : DataManager.DataRoot;

            if (ExternalTargetLauncher.TryResolveDataRootScopedFilePath(target, dataRoot, out string localPath))
            {
                return ExternalTargetLauncher.TryStart(localPath, $"local file [{localPath}]");
            }

            // Logger, not Debug.WriteLine: the latter carries an implicit [Conditional("DEBUG")] and
            // is erased from RELEASE builds, which is precisely where a refused link needs to leave a
            // trace - a user reporting "the link does nothing" has only the log to send.
            Logger.Warning($"Rejected external target outside allowed scope: [{target}]");
            return false;
        }

        // ###########################################################################################
        // Validates that a target string is an allowed absolute URI.
        // ###########################################################################################
        private static bool TryCreateAllowedUri(string target, out Uri uri)
        {
            uri = null!;

            if (!Uri.TryCreate(target.Trim(), UriKind.Absolute, out Uri? candidateUri))
                return false;

            if (!string.Equals(candidateUri.Scheme, Uri.UriSchemeHttp, StringComparison.OrdinalIgnoreCase) &&
                !string.Equals(candidateUri.Scheme, Uri.UriSchemeHttps, StringComparison.OrdinalIgnoreCase) &&
                !string.Equals(candidateUri.Scheme, "mailto", StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            uri = candidateUri;
            return true;
        }

        // ###########################################################################################
        // Resolves a local file path and rejects anything outside the configured data-root, plus
        // any file whose extension is not on the openable-document allowlist.
        // Relative paths are resolved against data-root; absolute paths must still stay inside it.
        //
        // The containment check runs TWICE: once on the lexically-normalized path (cheap, and
        // enough to reject an ordinary traversal or an absolute path elsewhere on disk before ever
        // touching the filesystem for the file's existence), and again on each side's REAL path
        // via RealPathResolver, which follows any symlink or junction along the way to where it
        // actually points. Path.GetFullPath alone cannot see a symlinked directory inside the data
        // root that redirects outside it - see RealPathResolver's own header for how that was
        // proven against a real Windows junction. Currently unreachable in this application (see
        // that header), which is exactly why it is worth having: the day something changes that,
        // this check does not need to change with it.
        //
        // The path handed back is the REAL one, which is also what the shell is given - see the
        // comment at the assignment for why the lexical path must not be substituted there.
        // ###########################################################################################
        private static bool TryResolveDataRootScopedFilePath(string target, string dataRoot, out string localPath)
        {
            localPath = string.Empty;

            if (string.IsNullOrWhiteSpace(dataRoot) || string.IsNullOrWhiteSpace(target))
                return false;

            try
            {
                string normalizedDataRoot = Path.GetFullPath(dataRoot);
                string normalizedTargetInput = target.Trim().Replace('/', Path.DirectorySeparatorChar);

                string normalizedTarget = Path.IsPathRooted(normalizedTargetInput)
                    ? Path.GetFullPath(normalizedTargetInput)
                    : Path.GetFullPath(Path.Combine(normalizedDataRoot, normalizedTargetInput));

                StringComparison pathComparison = ExternalTargetLauncher.GetPathComparison();

                if (!ExternalTargetLauncher.IsContainedWithinRoot(normalizedTarget, normalizedDataRoot, pathComparison))
                    return false;

                if (!ExternalTargetLauncher.HasAllowedFileExtension(normalizedTarget))
                    return false;

                if (!File.Exists(normalizedTarget))
                    return false;

                // The existence check above is also what makes it safe to resolve real paths now:
                // RealPathResolver walks the filesystem, and calling it on something that might not
                // exist is the caller's job to guard, not its own (see its own header).
                //
                // BOTH sides must genuinely resolve. A resolved path compared against an unresolved
                // one is a verdict about neither: when the data root itself sits behind a link (an
                // AppData folder redirected to another volume, or --data-root= pointing at one) and
                // the root's own resolve fails while the target's succeeds, the two disagree and a
                // legitimate file is refused. So a failure to resolve either side refuses the open
                // outright rather than falling back to a comparison that cannot mean anything.
                if (!RealPathResolver.TryResolveRealPath(normalizedTarget, out string realTarget) ||
                    !ExternalTargetLauncher.TryGetRealDataRoot(normalizedDataRoot, out string realDataRoot))
                {
                    return false;
                }

                if (!ExternalTargetLauncher.IsContainedWithinRoot(realTarget, realDataRoot, pathComparison))
                    return false;

                // The RESOLVED path, not the lexical one. TryStart hands this to the shell, and
                // handing over a different string than the one just validated leaves a window in
                // which swapping a directory component for a link opens a file that never passed
                // the check. The two are identical whenever no link is involved, so this costs
                // nothing in the normal case.
                localPath = realTarget;
                return true;
            }
            catch
            {
                return false;
            }
        }

        // ###########################################################################################
        // The data root's own real path, resolved once per distinct root and remembered.
        //
        // Resolving it walks the root component by component, issuing a Directory.Exists plus a
        // Directory.ResolveLinkTarget per segment - roughly sixteen syscalls for a typical AppData
        // root, on EVERY link and file the user opens, for an answer that cannot change while the
        // app runs. DataManager.DataRoot is fixed for the process lifetime, and the only other
        // value that reaches here is a test's explicit override, so the root is keyed rather than
        // assumed single: a test pointing at a fresh temp folder must not be handed the previous
        // one's answer.
        //
        // Only SUCCESSFUL resolutions are cached. A failure is transient by nature (a permission
        // blip, a race with something being deleted), and remembering it would turn one bad moment
        // into every subsequent open being refused for the life of the process.
        // ###########################################################################################
        private static readonly ConcurrentDictionary<string, string> RealDataRootCache = new();

        private static bool TryGetRealDataRoot(string normalizedDataRoot, out string realDataRoot)
        {
            if (ExternalTargetLauncher.RealDataRootCache.TryGetValue(normalizedDataRoot, out string? cached))
            {
                realDataRoot = cached;
                return true;
            }

            if (!RealPathResolver.TryResolveRealPath(normalizedDataRoot, out realDataRoot))
            {
                return false;
            }

            ExternalTargetLauncher.RealDataRootCache[normalizedDataRoot] = realDataRoot;
            return true;
        }

        // ###########################################################################################
        // Whether normalizedTarget sits inside normalizedRoot - shared by the lexical and the
        // real-path passes above, so the two cannot disagree about what "contained" means.
        // ###########################################################################################
        private static bool IsContainedWithinRoot(string normalizedTarget, string normalizedRoot, StringComparison pathComparison)
        {
            if (string.Equals(normalizedTarget, normalizedRoot, pathComparison))
                return false;

            string normalizedRootWithSeparator = ExternalTargetLauncher.AppendDirectorySeparator(normalizedRoot);
            return normalizedTarget.StartsWith(normalizedRootWithSeparator, pathComparison);
        }

        // ###########################################################################################
        // Returns whether the normalized path carries an extension from the openable allowlist.
        // The extension is taken from the already-normalized full path, so Windows quirks like
        // trailing dots or alternate data streams cannot smuggle a second, executable extension.
        // ###########################################################################################
        private static bool HasAllowedFileExtension(string normalizedPath)
        {
            string extension = Path.GetExtension(normalizedPath);

            return !string.IsNullOrEmpty(extension) &&
                   ExternalTargetLauncher.AllowedFileExtensions.Contains(extension);
        }

        // ###########################################################################################
        // Starts a validated target through the operating system shell.
        // ###########################################################################################
        private static bool TryStart(string fileName, string description)
        {
            try
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = fileName,
                    UseShellExecute = true
                });

                return true;
            }
            catch (Exception ex)
            {
                Logger.Warning($"Failed to open {description} - [{ex.Message}]");
                return false;
            }
        }

        // ###########################################################################################
        // Appends a trailing directory separator when missing so StartsWith path checks stay safe.
        // ###########################################################################################
        private static string AppendDirectorySeparator(string path)
        {
            if (string.IsNullOrEmpty(path))
                return Path.DirectorySeparatorChar.ToString();

            if (path.EndsWith(Path.DirectorySeparatorChar.ToString(), StringComparison.Ordinal) ||
                path.EndsWith(Path.AltDirectorySeparatorChar.ToString(), StringComparison.Ordinal))
            {
                return path;
            }

            return path + Path.DirectorySeparatorChar;
        }

        // ###########################################################################################
        // Returns the correct filesystem path comparison for the current operating system.
        // ###########################################################################################
        private static StringComparison GetPathComparison()
        {
            return RuntimeInformation.IsOSPlatform(OSPlatform.Windows)
                ? StringComparison.OrdinalIgnoreCase
                : StringComparison.Ordinal;
        }
    }
}