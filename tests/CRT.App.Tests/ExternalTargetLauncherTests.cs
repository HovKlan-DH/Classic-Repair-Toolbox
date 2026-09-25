using System.Reflection;
using CRT;

namespace ClassicRepairToolbox.Tests;

// Characterisation tests for ExternalTargetLauncher - the only sanctioned way the UI opens
// an external link or a local file. Board data is community-contributed, so this is a real
// trust boundary: it must accept http/https/mailto and files inside the data root, and
// reject everything else.
//
// IMPORTANT - why this file uses reflection:
// TryOpen's SUCCESS path calls Process.Start(UseShellExecute = true), which would really
// launch a browser or open a file on whatever machine runs these tests. So the accept cases
// are exercised against the two private predicates that decide the outcome, and only the
// REJECT cases go through the public TryOpen (they return false without starting anything).
//
// If these reflection lookups ever fail, the fix is to make the predicates `internal` and
// add [InternalsVisibleTo] to the app project - not to call TryOpen with a valid target.
public sealed class ExternalTargetLauncherTests : IDisposable
{
    private readonly string thisDataRoot;
    private readonly string thisOutsideRoot;

    public ExternalTargetLauncherTests()
    {
        string baseDir = Path.Combine(Path.GetTempPath(), "crt-tests-" + Guid.NewGuid().ToString("N"));
        this.thisDataRoot = Path.GetFullPath(Path.Combine(baseDir, "Data"));
        this.thisOutsideRoot = Path.GetFullPath(Path.Combine(baseDir, "Outside"));

        Directory.CreateDirectory(Path.Combine(this.thisDataRoot, "Commodore", "C64"));
        Directory.CreateDirectory(this.thisOutsideRoot);

        File.WriteAllText(Path.Combine(this.thisDataRoot, "Commodore", "C64", "notes.txt"), "inside");
        File.WriteAllText(Path.Combine(this.thisOutsideRoot, "secrets.txt"), "outside");
    }

    public void Dispose()
    {
        try
        {
            string baseDir = Path.GetDirectoryName(this.thisDataRoot)!;
            if (Directory.Exists(baseDir))
            {
                Directory.Delete(baseDir, recursive: true);
            }
        }
        catch
        {
            // A leftover temp folder is not worth failing a test run over.
        }
    }

    // ------------------------------------------------------------- private predicates

    private static MethodInfo GetPrivateMethod(string name)
    {
        MethodInfo? method = typeof(ExternalTargetLauncher)
            .GetMethod(name, BindingFlags.NonPublic | BindingFlags.Static);

        Assert.True(method is not null,
            $"ExternalTargetLauncher.{name} not found - if it was renamed, update these tests.");

        return method!;
    }

    private static bool TryResolveDataRootScopedFilePath(string target, string dataRoot, out string localPath)
    {
        object?[] args = { target, dataRoot, null };
        bool result = (bool)GetPrivateMethod("TryResolveDataRootScopedFilePath").Invoke(null, args)!;
        localPath = (string)(args[2] ?? string.Empty);
        return result;
    }

    private static bool TryCreateAllowedUri(string target, out Uri? uri)
    {
        object?[] args = { target, null };
        bool result = (bool)GetPrivateMethod("TryCreateAllowedUri").Invoke(null, args)!;
        uri = args[1] as Uri;
        return result;
    }

    // ----------------------------------------------------------------- allowed schemes

    [Theory]
    [InlineData("http://classic-repair-toolbox.dk")]
    [InlineData("https://classic-repair-toolbox.dk/data")]
    [InlineData("HTTPS://UPPER.CASE.SCHEME")]
    [InlineData("mailto:someone@example.com")]
    [InlineData("  https://leading-and-trailing-space.example  ")]
    public void Allowed_uri_schemes_are_accepted(string target)
    {
        Assert.True(TryCreateAllowedUri(target, out Uri? uri));
        Assert.NotNull(uri);
    }

    [Theory]
    [InlineData("ftp://example.com/file.zip")]
    [InlineData("file:///C:/Windows/System32/calc.exe")]
    [InlineData("javascript:alert(1)")]
    [InlineData("data:text/html;base64,PHNjcmlwdD4=")]
    [InlineData("vbscript:msgbox(1)")]
    [InlineData("ms-msdt:/id")]
    [InlineData("not a uri at all")]
    [InlineData("/relative/path.txt")]
    public void Other_uri_schemes_are_not_accepted_as_uris(string target)
    {
        Assert.False(TryCreateAllowedUri(target, out _));
    }

    // ------------------------------------------------------------- data-root scoping

    [Fact]
    public void A_relative_path_resolving_inside_the_data_root_is_accepted()
    {
        Assert.True(TryResolveDataRootScopedFilePath(
            "Commodore/C64/notes.txt", this.thisDataRoot, out string localPath));

        Assert.Equal(
            Path.Combine(this.thisDataRoot, "Commodore", "C64", "notes.txt"),
            localPath);
    }

    [Fact]
    public void Forward_slashes_are_accepted_on_every_platform()
    {
        // Board data is authored on all three OSes, so separators are normalised.
        Assert.True(TryResolveDataRootScopedFilePath(
            "Commodore/C64/notes.txt", this.thisDataRoot, out _));
    }

    [Fact]
    public void An_absolute_path_inside_the_data_root_is_accepted()
    {
        string absolute = Path.Combine(this.thisDataRoot, "Commodore", "C64", "notes.txt");

        Assert.True(TryResolveDataRootScopedFilePath(absolute, this.thisDataRoot, out string localPath));
        Assert.Equal(absolute, localPath);
    }

    [Fact]
    public void A_directory_traversal_escaping_the_data_root_is_rejected()
    {
        // The attack this boundary exists to stop.
        Assert.False(TryResolveDataRootScopedFilePath(
            "../Outside/secrets.txt", this.thisDataRoot, out _));
    }

    [Fact]
    public void A_deep_directory_traversal_escaping_the_data_root_is_rejected()
    {
        Assert.False(TryResolveDataRootScopedFilePath(
            "Commodore/C64/../../../Outside/secrets.txt", this.thisDataRoot, out _));
    }

    [Fact]
    public void An_absolute_path_outside_the_data_root_is_rejected()
    {
        string absolute = Path.Combine(this.thisOutsideRoot, "secrets.txt");

        Assert.False(TryResolveDataRootScopedFilePath(absolute, this.thisDataRoot, out _));
    }

    [Fact]
    public void The_data_root_itself_is_rejected()
    {
        Assert.False(TryResolveDataRootScopedFilePath(this.thisDataRoot, this.thisDataRoot, out _));
    }

    [Fact]
    public void A_sibling_directory_with_the_data_root_as_a_name_prefix_is_rejected()
    {
        // Guards the trailing-separator check: "<root>Evil" must not pass a StartsWith test
        // against "<root>". This is what AppendDirectorySeparator is for.
        string siblingWithPrefix = this.thisDataRoot + "Evil";
        Directory.CreateDirectory(siblingWithPrefix);
        string planted = Path.Combine(siblingWithPrefix, "secrets.txt");
        File.WriteAllText(planted, "outside");

        Assert.False(TryResolveDataRootScopedFilePath(planted, this.thisDataRoot, out _));
    }

    // ###########################################################################################
    // A LINKED DIRECTORY inside the data root, pointing outside it, must not smuggle a file
    // through the containment check.
    //
    // Path.GetFullPath is purely lexical - it never touches the filesystem, so it cannot see that
    // "<root>/link" is really "<outsideRoot>". Proven with a real link on disk rather than argued
    // from documentation, and on WHICHEVER kind of link this platform can create without special
    // privilege:
    //
    //   - Linux (the CI runner, ubuntu-latest per build-and-unittest.yml): an ordinary user can
    //     always create a symlink, so Directory.CreateSymbolicLink is the real, unconditional test
    //     there.
    //   - Windows: Directory.CreateSymbolicLink needs a privilege this account may not hold
    //     outside Developer Mode (confirmed on this project's own dev box), but a JUNCTION
    //     (mklink /J) needs none and .NET's Directory.ResolveLinkTarget follows one exactly like a
    //     symlink - RealPathResolver's own header records the reflection-based proof this test
    //     mirrors. So Windows falls back to a junction rather than skipping outright, and every
    //     platform this suite runs on gets a real, non-conditional assertion.
    // ###########################################################################################
    private static void CreateDirectoryLink(string linkPath, string targetPath)
    {
        if (OperatingSystem.IsWindows())
        {
            var mklink = System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(
                "cmd.exe", $"/c mklink /J \"{linkPath}\" \"{targetPath}\"")
            {
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true
            })!;
            mklink.WaitForExit();

            if (mklink.ExitCode != 0)
            {
                throw new IOException(
                    $"mklink /J failed ({mklink.ExitCode}): {mklink.StandardError.ReadToEnd()}");
            }

            return;
        }

        Directory.CreateSymbolicLink(linkPath, targetPath);
    }

    // ###########################################################################################
    // Removes a directory link ITSELF, never its target - the reparse point/link entry, not what
    // it points at.
    //
    // WHY THIS EXISTS: Directory.Delete(recursive: true) on a folder that CONTAINS one of these
    // links follows it and deletes the TARGET's real contents too - confirmed directly, and it is
    // not a theoretical edge case: this class's own Dispose() calls exactly that recursive delete
    // on thisDataRoot's parent, and thisOutsideRoot (holding secrets.txt, the fixture several
    // other tests in this file also use) sits right next to it. A test that plants a link and
    // then lets Dispose's recursive delete reach it would destroy that shared fixture's real
    // files. Removing the link explicitly, before Dispose ever runs, is what keeps that recursive
    // delete confined to the actual data-root tree it means to clean up.
    // ###########################################################################################
    private static void RemoveDirectoryLink(string linkPath)
    {
        try
        {
            if (OperatingSystem.IsWindows())
            {
                // rmdir with no /S removes only the reparse point, not what it targets.
                var rmdir = System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(
                    "cmd.exe", $"/c rmdir \"{linkPath}\"")
                {
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true
                })!;
                rmdir.WaitForExit();
                return;
            }

            Directory.Delete(linkPath);
        }
        catch
        {
            // Best-effort, same reasoning as Dispose's own catch: a leftover link in a temp
            // folder that is about to be abandoned anyway is not worth failing a test run over.
        }
    }

    [Fact]
    public void A_linked_directory_pointing_outside_the_root_is_rejected()
    {
        string linkPath = Path.Combine(this.thisDataRoot, "escape-link");
        CreateDirectoryLink(linkPath, this.thisOutsideRoot);

        try
        {
            Assert.False(TryResolveDataRootScopedFilePath(
                "escape-link/secrets.txt", this.thisDataRoot, out _));
        }
        finally
        {
            RemoveDirectoryLink(linkPath);
        }
    }

    // The other direction: a link whose target IS still inside the root must be accepted - the
    // fix must not reject every link, only the ones that escape.
    [Fact]
    public void A_linked_directory_pointing_back_inside_the_root_is_accepted()
    {
        string realTarget = Path.Combine(this.thisDataRoot, "Commodore", "C64");
        string linkPath = Path.Combine(this.thisDataRoot, "alias");
        CreateDirectoryLink(linkPath, realTarget);

        try
        {
            Assert.True(TryResolveDataRootScopedFilePath("alias/notes.txt", this.thisDataRoot, out _));
        }
        finally
        {
            RemoveDirectoryLink(linkPath);
        }
    }

    [Fact]
    public void A_path_inside_the_data_root_that_does_not_exist_is_rejected()
    {
        Assert.False(TryResolveDataRootScopedFilePath(
            "Commodore/C64/missing.txt", this.thisDataRoot, out _));
    }

    [Fact]
    public void A_directory_inside_the_data_root_is_rejected_because_it_is_not_a_file()
    {
        Assert.False(TryResolveDataRootScopedFilePath(
            "Commodore/C64", this.thisDataRoot, out _));
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void An_empty_target_is_rejected(string target)
    {
        Assert.False(TryResolveDataRootScopedFilePath(target, this.thisDataRoot, out _));
    }

    // ------------------------------------------------------- file-extension allowlist
    //
    // TryStart hands the resolved path to the OS shell (UseShellExecute = true), and the shell
    // RUNS executables, scripts and shortcuts rather than displaying them. Board data is synced
    // from the network and community-contributed, so a *.exe/*.bat/*.lnk landing inside the data
    // root must never become code execution just because a workbook cell references it. Only
    // document, image and data formats may be opened; anything else fails closed.
    //
    // These cases stay on the private predicate on purpose: planting "tool.exe" and calling the
    // public TryOpen would hand the file to the OS shell if the check ever regressed.

    /// <summary>Creates a harmless file inside the temp data root and returns its relative path.</summary>
    private string PlantDataRootFile(string fileName)
    {
        string relative = Path.Combine("Commodore", "C64", fileName);
        File.WriteAllText(Path.Combine(this.thisDataRoot, relative), "planted by test");
        return relative;
    }

    [Theory]
    [InlineData("tool.exe")]
    [InlineData("script.bat")]
    [InlineData("script.cmd")]
    [InlineData("legacy.com")]
    [InlineData("screensaver.scr")]
    [InlineData("script.ps1")]
    [InlineData("script.vbs")]
    [InlineData("script.js")]
    [InlineData("installer.msi")]
    [InlineData("shortcut.lnk")]
    [InlineData("library.dll")]
    [InlineData("script.sh")]
    [InlineData("launcher.desktop")]
    [InlineData("archive.jar")]
    [InlineData("TOOL.EXE")]
    public void An_executable_script_or_shortcut_inside_the_data_root_is_rejected(string fileName)
    {
        string relative = this.PlantDataRootFile(fileName);

        Assert.False(TryResolveDataRootScopedFilePath(relative, this.thisDataRoot, out _));
    }

    [Fact]
    public void A_file_without_an_extension_is_rejected()
    {
        // Fail closed: no extension means no way to know what the shell would do with it.
        string relative = this.PlantDataRootFile("README");

        Assert.False(TryResolveDataRootScopedFilePath(relative, this.thisDataRoot, out _));
    }

    [Theory]
    [InlineData("datasheet.pdf")]
    [InlineData("photo.png")]
    [InlineData("photo.jpg")]
    [InlineData("notes.txt")]
    [InlineData("reference.html")]
    [InlineData("board-data.xlsx")]
    [InlineData("pinout.csv")]
    [InlineData("capture.fsc")]
    [InlineData("calibration.json")]
    [InlineData("Issue 4.sch")]
    [InlineData("board.kicad_pcb")]
    [InlineData("DATASHEET.PDF")]
    public void A_document_image_or_data_file_inside_the_data_root_is_accepted(string fileName)
    {
        string relative = this.PlantDataRootFile(fileName);

        Assert.True(TryResolveDataRootScopedFilePath(relative, this.thisDataRoot, out _));
    }

    [Fact]
    public void An_empty_data_root_rejects_everything()
    {
        Assert.False(TryResolveDataRootScopedFilePath("notes.txt", "", out _));
    }

    // ------------------------------------------------------ end-to-end (reject only)

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void TryOpen_rejects_an_empty_target(string target)
    {
        Assert.False(ExternalTargetLauncher.TryOpen(target, this.thisDataRoot));
    }

    [Theory]
    [InlineData("../Outside/secrets.txt")]
    [InlineData("Commodore/C64/../../../Outside/secrets.txt")]
    [InlineData("Commodore/C64/missing.txt")]
    public void TryOpen_rejects_a_target_outside_or_absent_from_the_data_root(string target)
    {
        // Safe to call end-to-end: every one of these returns false before Process.Start.
        Assert.False(ExternalTargetLauncher.TryOpen(target, this.thisDataRoot));
    }

    [Fact]
    public void TryOpen_rejects_an_absolute_path_outside_the_data_root()
    {
        string absolute = Path.Combine(this.thisOutsideRoot, "secrets.txt");
        Assert.False(ExternalTargetLauncher.TryOpen(absolute, this.thisDataRoot));
    }

    // ------------------------------------------------------------- OpenableFileExtensions

    // The allowlist is exposed so attach-time validation can share it, but it must be a COPY:
    // IReadOnlyCollection is only a compile-time promise, and the backing store is a HashSet, so
    // handing back the live instance would let any caller in the assembly cast it and Add(".exe"),
    // permanently defeating the executable rejection this class exists to enforce.
    [Fact]
    public void The_exposed_extension_list_cannot_be_mutated_through_a_cast()
    {
        var exposed = ExternalTargetLauncher.OpenableFileExtensions;

        if (exposed is HashSet<string> writable)
        {
            writable.Add(".exe");
        }
        else if (exposed is ICollection<string> collection && !collection.IsReadOnly)
        {
            collection.Add(".exe");
        }

        // Whatever the caller managed to do to their copy, the launcher still refuses executables.
        Assert.False(ExternalTargetLauncher.IsOpenableFile("installer.exe"));
        Assert.DoesNotContain(".exe", ExternalTargetLauncher.OpenableFileExtensions);
    }

    // The set still has to be usable for what it exists for - building a picker filter.
    [Fact]
    public void The_exposed_extension_list_carries_the_openable_types()
    {
        var extensions = ExternalTargetLauncher.OpenableFileExtensions;

        Assert.Contains(".pdf", extensions);
        Assert.Contains(".json", extensions);
        Assert.Contains(".png", extensions);
        Assert.DoesNotContain(".exe", extensions);
        Assert.DoesNotContain(".bat", extensions);
    }

    // IsOpenableFile decides whether a file may be ATTACHED; the private HasAllowedFileExtension
    // decides whether it may be OPENED. They are documented as the same set, so they must not
    // disagree about the same string - otherwise a file can be accepted at attach time and refused
    // (or resolved differently) at open time.
    //
    // The forms that used to split them are the Windows ones a raw Path.GetExtension cannot see
    // through: a trailing dot, and an alternate data stream suffix. Reading the extension from the
    // normalized full path collapses both.
    [Theory]
    [InlineData("payload.exe")]
    [InlineData("payload.exe.")]
    [InlineData("notes.txt:evil.exe")]
    [InlineData("script.bat")]
    [InlineData("shortcut.lnk")]
    [InlineData("noextension")]
    public void An_unopenable_file_is_refused_for_attaching_too(string name)
    {
        Assert.False(ExternalTargetLauncher.IsOpenableFile(name));
    }

    // The ordinary documents the worklog's Files section exists to carry still pass.
    [Theory]
    [InlineData("manual.pdf")]
    [InlineData("data.json")]
    [InlineData("sheet.xlsx")]
    [InlineData("photo.png")]
    public void An_openable_document_is_accepted_for_attaching(string name)
    {
        Assert.True(ExternalTargetLauncher.IsOpenableFile(name));
    }
}
