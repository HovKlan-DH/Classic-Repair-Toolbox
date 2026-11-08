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
        // *** TryOpen FOR A PLACE THAT SAYS NOTHING OF ITS OWN WHEN IT FAILS (owner request, 2026-11-07: every error in the
        // UI). *** A link in a note, a file on the "Resources" tab, a help link: a click that opened nothing was a click
        // that seemed to do nothing, its reason in the log alone. Refused or failed, it is said on Main's "Something went
        // wrong" card - what it was - and why is the line above it in "Errors and warnings". A place that says it itself
        // (the component pane, the table's hover card) keeps TryOpen.
        // ###########################################################################################
        public static bool OpenOrSay(string target, string? dataRootOverride = null)
        {
            bool opened = ExternalTargetLauncher.TryOpen(target, dataRootOverride);

            if (!opened)
            {
                Logger.Critical($"[{ExternalTargetLauncher.ShortNameOf(target)}] could not be opened");
            }

            return opened;
        }

        // ###########################################################################################
        // *** A STORED LINK ROW, OPENED BY LinkTarget's RULE (code review, 2026-11-08). *** A link is stored as typed -
        // "www.c64-wiki.com" as likely as "https://..." - and the exported PDF and the Contribute tab open it through
        // LinkTarget.For. A device's links and a worklog's handed it over as typed: tried as a file and refused, the error
        // card said it could not be opened. `openInstead` stands in for the opening in a test; one that is no link CRT
        // opens is said as a refused target is.
        // ###########################################################################################
        public static bool OpenLinkOrSay(string? url, Action<string>? openInstead = null)
        {
            if (LinkTarget.For(url) is not { } target)
            {
                Logger.Critical($"[{ExternalTargetLauncher.ShortNameOf(url)}] could not be opened");
                return false;
            }

            if (openInstead is not null)
            {
                openInstead(target);
                return true;
            }

            return ExternalTargetLauncher.OpenOrSay(target);
        }

        // A web address as it is, a file by its name.
        internal static string ShortNameOf(string? target)
        {
            string text = target?.Trim() ?? string.Empty;

            if (text.Length == 0 || text.Contains("://", StringComparison.Ordinal) || text.StartsWith("mailto:", StringComparison.OrdinalIgnoreCase))
            {
                return text;
            }

            string name = Path.GetFileName(text.Replace('\\', '/').TrimEnd('/'));

            return name.Length > 0 ? name : text;
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
        // *** REVEAL: THE FILE SHOWN SELECTED IN THE SYSTEM'S FILE MANAGER *** (owner request, 2026-11-07: 'a button named
        // "Reveal" and then open and select the file in File Explorer') - Explorer's "/select", the Finder's "open -R", and
        // on Linux the freedesktop FileManager1 interface most file managers answer (Nautilus, Dolphin, Nemo, Thunar,
        // Caja...), its folder opened instead where none does. The file must lie inside `root` as TryOpen's must, through
        // no link out of it - its extension is free, as nothing is run: the file manager only shows it.
        // ###########################################################################################
        public static bool TryReveal(string path, string root)
        {
            if (!ExternalTargetLauncher.TryResolveDataRootScopedFilePath(path, root, out string localPath, requireOpenableExtension: false))
            {
                Logger.Warning($"Rejected a file to reveal outside the allowed folder: [{path}]");
                return false;
            }

            // Windows: the shell's own call, which selects the file every time (owner report, 2026-11-07: the folder
            // opened, "but not every time select the file ... it is the second time I press the button that it will reveal
            // it") - explorer.exe's "/select" is known to open a NEW window on the folder and lose the selection while that
            // window is still being made. Explorer's command is what is left when the call fails.
            if (OperatingSystem.IsWindows())
            {
                WindowsFileSelection.Select(localPath);
                return true;
            }

            RevealDesktop desktop = ExternalTargetLauncher.CurrentDesktop();
            ProcessStartInfo start = ExternalTargetLauncher.RevealCommandFor(desktop, localPath);
            Process? process;

            try
            {
                process = Process.Start(start);
            }
            catch (Exception ex)
            {
                Logger.Warning($"Failed to reveal [{localPath}] - [{ex.Message}]");
                process = null;
            }

            if (process is null)
            {
                // ###########################################################################################
                // Linux: dbus-send not there to ask the file manager (a package of its own on some systems) - its folder
                // opened, as when no file manager answers (code review, 2026-11-08: nothing was opened at all).
                // ###########################################################################################
                return ExternalTargetLauncher.RevealFallsBackToFolder(desktop) && ExternalTargetLauncher.OpenFolderOf(localPath);
            }

            // Linux: no file manager answering the interface - its folder opened, as near as Linux comes. The Finder's
            // "open" is let go of at once: it hands the request over and ends.
            if (desktop == RevealDesktop.Linux)
            {
                _ = ExternalTargetLauncher.OpenFolderIfNotAnsweredAsync(process, localPath);
            }
            else
            {
                process.Dispose();
            }

            return true;
        }

        // Whether a reveal that could not start falls back to opening the file's folder - Linux's alone, where the file
        // manager is asked through a program that may not be there.
        internal static bool RevealFallsBackToFolder(RevealDesktop desktop) => desktop == RevealDesktop.Linux;

        // The folder `path` lies in, opened - false when there is none or it would not open.
        private static bool OpenFolderOf(string path) =>
            Path.GetDirectoryName(path) is { Length: > 0 } folder && ExternalTargetLauncher.TryStart(folder, $"folder [{folder}]");

        // The three ways of revealing a file.
        internal enum RevealDesktop
        {
            Windows,
            MacOs,
            Linux,
        }

        private static RevealDesktop CurrentDesktop() =>
            OperatingSystem.IsWindows() ? RevealDesktop.Windows : OperatingSystem.IsMacOS() ? RevealDesktop.MacOs : RevealDesktop.Linux;

        // ###########################################################################################
        // What is started to reveal `path` (a full path, checked) - its own method, so the tests read it without starting
        // anything. Explorer reads "/select," and the path as one argument, quoted for spaces (a Windows path holds no
        // quote); the Finder and dbus-send take theirs as separate arguments.
        // ###########################################################################################
        internal static ProcessStartInfo RevealCommandFor(RevealDesktop desktop, string path)
        {
            switch (desktop)
            {
                case RevealDesktop.Windows:
                    return new ProcessStartInfo { FileName = "explorer.exe", Arguments = $"/select,\"{path}\"", UseShellExecute = false };

                case RevealDesktop.MacOs:
                {
                    var finder = new ProcessStartInfo { FileName = "open", UseShellExecute = false };
                    finder.ArgumentList.Add("-R");
                    finder.ArgumentList.Add(path);
                    return finder;
                }

                default:
                {
                    var dbus = new ProcessStartInfo { FileName = "dbus-send", UseShellExecute = false };

                    foreach (string argument in new[]
                             {
                                 "--session",
                                 "--print-reply",
                                 "--dest=org.freedesktop.FileManager1",
                                 "--type=method_call",
                                 "/org/freedesktop/FileManager1",
                                 "org.freedesktop.FileManager1.ShowItems",
                                 "array:string:" + ExternalTargetLauncher.FileUriForDbusArray(path),
                                 "string:",
                             })
                    {
                        dbus.ArgumentList.Add(argument);
                    }

                    return dbus;
                }
            }
        }

        // ###########################################################################################
        // `path` - a Linux path, this being Linux's way alone - as a file URI for dbus-send's "array:string:", which splits
        // its items at every comma with no way to escape one (code review, 2026-11-07: "Schematic, rev B.pdf" went as two
        // broken addresses, and the file manager selected nothing). Each name between the slashes is escaped as data, so a
        // comma goes as "%2C" and a "%" as "%25" - Uri left the comma as it was, and took "%41" in a path for an escaped "A"
        // (code review, 2026-11-08: "a%41b.pdf" went as "aAb.pdf", a file that is not there). Whoever reads the URI reads
        // the name exactly.
        // ###########################################################################################
        internal static string FileUriForDbusArray(string path) =>
            "file://" + string.Join('/', path.Split('/').Select(Uri.EscapeDataString));

        // ###########################################################################################
        // Windows: SHOpenFolderAndSelectItems - the folder opened in Explorer, or the window already on it used, and the
        // file selected there. On a thread of its own (the call wants an apartment of its own, and opening Explorer can
        // take a second), never holding the window up; explorer.exe's "/select" when the shell refuses.
        // ###########################################################################################
        [System.Runtime.Versioning.SupportedOSPlatform("windows")]
        private static class WindowsFileSelection
        {
            [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
            private static extern int SHParseDisplayName(string name, IntPtr bindingContext, out IntPtr pidl, uint attributesWanted, out uint attributes);

            [DllImport("shell32.dll")]
            private static extern int SHOpenFolderAndSelectItems(IntPtr pidlFolder, uint count, IntPtr children, uint flags);

            [DllImport("shell32.dll")]
            private static extern void ILFree(IntPtr pidl);

            internal static void Select(string path)
            {
                // Neither the shell nor Explorer showing it is said on CRT's error card (code review, 2026-11-08): the
                // caller was answered at once - the work is on this thread - so nothing else says it.
                var thread = new System.Threading.Thread(() =>
                {
                    if (!WindowsFileSelection.TrySelect(path) && !WindowsFileSelection.RevealWithExplorer(path))
                    {
                        Logger.Critical($"[{ExternalTargetLauncher.ShortNameOf(path)}] could not be shown in Explorer");
                    }
                })
                {
                    IsBackground = true,
                    Name = "Reveal a file",
                };

                thread.SetApartmentState(System.Threading.ApartmentState.STA);
                thread.Start();
            }

            private static bool TrySelect(string path)
            {
                IntPtr item = IntPtr.Zero;

                try
                {
                    int parsed = WindowsFileSelection.SHParseDisplayName(path, IntPtr.Zero, out item, 0, out _);

                    if (parsed != 0 || item == IntPtr.Zero)
                    {
                        Logger.Warning($"Reveal: the shell could not read the path [{path}] (0x{parsed:X8}) - Explorer is asked instead");
                        return false;
                    }

                    // No children: the item itself is selected in its folder.
                    int opened = WindowsFileSelection.SHOpenFolderAndSelectItems(item, 0, IntPtr.Zero, 0);

                    if (opened != 0)
                    {
                        Logger.Warning($"Reveal: the shell could not show [{path}] (0x{opened:X8}) - Explorer is asked instead");
                        return false;
                    }

                    Logger.Info($"Revealed [{path}] in Explorer");
                    return true;
                }
                catch (Exception ex) when (ex is DllNotFoundException or EntryPointNotFoundException or COMException)
                {
                    Logger.Warning($"Reveal: the shell call failed for [{path}] - [{ex.Message}] - Explorer is asked instead");
                    return false;
                }
                finally
                {
                    if (item != IntPtr.Zero)
                    {
                        WindowsFileSelection.ILFree(item);
                    }
                }
            }

            // Explorer's "/select" - false when it could not be started.
            private static bool RevealWithExplorer(string path)
            {
                try
                {
                    using Process? process = Process.Start(ExternalTargetLauncher.RevealCommandFor(RevealDesktop.Windows, path));
                    return process is not null;
                }
                catch (Exception ex)
                {
                    Logger.Warning($"Failed to reveal [{path}] - [{ex.Message}]");
                    return false;
                }
            }
        }

        // ###########################################################################################
        // Linux: the file's folder opened when no file manager answered ShowItems in a few seconds - one that never answers
        // too (code review, 2026-11-07: a FileManager1 name registered by a service that hung left dbus-send waiting its own
        // 25 seconds; the wait gave up after 5 and returned, and "Reveal" opened nothing at all). dbus-send is stopped then.
        // ###########################################################################################
        private static async System.Threading.Tasks.Task OpenFolderIfNotAnsweredAsync(Process process, string path)
        {
            bool answered = false;

            try
            {
                answered = await ExternalTargetLauncher.AnsweredInTimeAsync(
                    async token =>
                    {
                        await process.WaitForExitAsync(token).ConfigureAwait(false);
                        return process.ExitCode;
                    },
                    TimeSpan.FromSeconds(5)).ConfigureAwait(false);

                if (!answered && !process.HasExited)
                {
                    process.Kill();
                }
            }
            catch (Exception ex)
            {
                Logger.Warning($"Revealing [{path}] through the file manager failed - [{ex.Message}]");
            }
            finally
            {
                process.Dispose();
            }

            if (!answered)
            {
                ExternalTargetLauncher.OpenFolderOf(path);
            }
        }

        // ###########################################################################################
        // Whether the file manager answered: `exitCode` - dbus-send's, once it ends - came within `patience`, and was 0.
        // Not ending in time is no answer, the same as a refusal: the folder is opened then.
        // ###########################################################################################
        internal static async System.Threading.Tasks.Task<bool> AnsweredInTimeAsync(
            Func<System.Threading.CancellationToken, System.Threading.Tasks.Task<int>> exitCode,
            TimeSpan patience)
        {
            using var waited = new System.Threading.CancellationTokenSource(patience);

            try
            {
                return await exitCode(waited.Token).ConfigureAwait(false) == 0;
            }
            catch (OperationCanceledException)
            {
                return false;
            }
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
        private static bool TryResolveDataRootScopedFilePath(string target, string dataRoot, out string localPath, bool requireOpenableExtension = true)
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

                if (requireOpenableExtension && !ExternalTargetLauncher.HasAllowedFileExtension(normalizedTarget))
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