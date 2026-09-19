using System.Diagnostics;
using CRT;

namespace ClassicRepairToolbox.Tests;

// Tests for RealPathResolver - the "follow every symlink/junction along a path to its real
// target" helper ExternalTargetLauncher (and, on the webserver side, the equivalent PHP check)
// relies on to stop Path.GetFullPath's purely lexical normalization from missing a linked
// directory that redirects outside the data root. See the class's own header for the full
// reasoning and where this is currently unreachable in the shipped app.
public sealed class RealPathResolverTests : IDisposable
{
    private readonly TempWorkspace thisWorkspace = new();
    private readonly List<string> thisLinksToRemoveBeforeDisposal = new();

    // ###########################################################################################
    // Removing a link before TempWorkspace's own Dispose runs matters here for the same reason it
    // does in ExternalTargetLauncherTests: Directory.Delete(recursive: true) FOLLOWS a directory
    // link and deletes what it points at, confirmed directly against this exact shape of fixture -
    // and a resolver test plants links pointing at OTHER folders inside the very same workspace,
    // so an unremoved link here would have Dispose delete a sibling fixture out from under a test
    // that has not finished using it.
    // ###########################################################################################
    public void Dispose()
    {
        foreach (string linkPath in this.thisLinksToRemoveBeforeDisposal)
        {
            RemoveDirectoryLink(linkPath);
        }

        this.thisWorkspace.Dispose();
    }

    private static void CreateDirectoryLink(string linkPath, string targetPath)
    {
        if (OperatingSystem.IsWindows())
        {
            var mklink = Process.Start(new ProcessStartInfo(
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

    private static void RemoveDirectoryLink(string linkPath)
    {
        try
        {
            if (OperatingSystem.IsWindows())
            {
                var rmdir = Process.Start(new ProcessStartInfo("cmd.exe", $"/c rmdir \"{linkPath}\"")
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
            // Best-effort - see Dispose's own reasoning.
        }
    }

    private string LinkedDirectory(string linkName, string targetPath)
    {
        string linkPath = this.thisWorkspace.Path_(linkName);
        CreateDirectoryLink(linkPath, targetPath);
        this.thisLinksToRemoveBeforeDisposal.Add(linkPath);
        return linkPath;
    }

    // -------------------------------------------------------------- no links present

    // Note the CANONICAL separators in the expectation. TempWorkspace.WriteFile builds its path
    // with Path.Combine, which leaves the "Commodore/C64" forward slashes in the caller's string
    // as they were, and the resolver rebuilds what it walks from its own segments - so a path
    // carrying mixed separators comes back with them normalized. That is the point rather than a
    // side effect: the caller compares this result against a root built the same way, and a
    // resolved path that still carried the other platform's separator would fail a containment
    // check it should pass. The earlier single-separator split "preserved" the input only because
    // it treated the whole run as one component and never walked into it at all.
    [Fact]
    public void A_plain_path_with_no_links_resolves_successfully_to_its_canonical_form()
    {
        string file = this.thisWorkspace.WriteFile("Commodore/C64/board.png", "x");

        Assert.True(RealPathResolver.TryResolveRealPath(file, out string resolved));
        Assert.Equal(this.thisWorkspace.Path_("Commodore", "C64", "board.png"), resolved);
    }

    // A path that is not there cannot be a link, so this is a genuine, SUCCESSFUL resolution -
    // not a failure to look. The distinction is the whole point of the bool: the caller refuses
    // the open on a failure, and a not-yet-existing path must not be made unopenable by that.
    [Fact]
    public void A_path_that_does_not_exist_still_resolves_successfully_and_is_unchanged()
    {
        string missing = this.thisWorkspace.Path_("Commodore", "missing.png");

        Assert.True(RealPathResolver.TryResolveRealPath(missing, out string resolved));
        Assert.Equal(missing, resolved);
    }

    // Nothing past the first missing component can be a link either - the walk must not throw
    // trying to inspect segments that are not there, and must still return the FULL path rather
    // than truncating it at the point resolution gave up.
    [Fact]
    public void A_path_whose_parent_directories_do_not_exist_either_is_returned_whole()
    {
        string deep = this.thisWorkspace.Path_("nope", "also-nope", "x.png");

        Assert.True(RealPathResolver.TryResolveRealPath(deep, out string resolved));
        Assert.Equal(deep, resolved);
    }

    [Fact]
    public void The_workspace_root_itself_resolves_unchanged()
    {
        Assert.True(RealPathResolver.TryResolveRealPath(this.thisWorkspace.Root, out string resolved));
        Assert.Equal(this.thisWorkspace.Root, resolved);
    }

    // -------------------------------------------------------------- resolution REFUSED
    //
    // These pin the contract the caller depends on to fail closed. Before it existed, every one of
    // these cases handed back the input string with no way to tell it apart from a successful
    // "nothing along this path was linked" - so ExternalTargetLauncher, which resolves the target
    // and the data root SEPARATELY and compares the two, could compare a resolved path against an
    // unresolved one and refuse a legitimate file (or, in the other direction, trust a comparison
    // that had not actually been performed).

    [Fact]
    public void An_empty_input_reports_failure_rather_than_a_resolution()
    {
        Assert.False(RealPathResolver.TryResolveRealPath(string.Empty, out string resolved));
        Assert.Equal(string.Empty, resolved);
    }

    // A relative path cannot be walked - there is no root to start from - and the class documents
    // that it takes an already-GetFullPath'd string. Reporting success here would let a caller
    // compare a relative path against an absolute root and read the inevitable "not contained" as
    // a real verdict.
    [Fact]
    public void A_relative_path_reports_failure_and_is_left_untouched()
    {
        Assert.False(RealPathResolver.TryResolveRealPath(Path.Combine("Commodore", "C64", "board.png"), out string resolved));
        Assert.Equal(Path.Combine("Commodore", "C64", "board.png"), resolved);
    }

    // The out value on a failure is still the plain input rather than null or a half-walked path,
    // so a caller that logs it has something meaningful - it is simply not a RESOLVED path, which
    // is what the false says.
    [Fact]
    public void A_failed_resolution_leaves_the_out_value_as_the_unmodified_input()
    {
        const string malformed = "not-a-rooted-path";

        Assert.False(RealPathResolver.TryResolveRealPath(malformed, out string resolved));
        Assert.Equal(malformed, resolved);
    }

    // -------------------------------------------------------------- links present
    //
    // Creates a real link on disk (a Windows junction via mklink /J, or a symlink everywhere
    // else - see CreateDirectoryLink) rather than arguing from documentation. mklink /J needs no
    // special privilege on Windows, unlike Directory.CreateSymbolicLink outside Developer Mode,
    // and .NET's own Directory.ResolveLinkTarget follows a junction exactly like a symlink - the
    // same reasoning ExternalTargetLauncherTests uses for the same reason.

    [Fact]
    public void A_linked_directory_resolves_to_its_real_target()
    {
        string realTarget = this.thisWorkspace.Path_("RealTarget");
        Directory.CreateDirectory(realTarget);
        File.WriteAllText(Path.Combine(realTarget, "board.png"), "x");

        string linkPath = this.LinkedDirectory("Linked", realTarget);
        string viaLink = Path.Combine(linkPath, "board.png");

        Assert.True(RealPathResolver.TryResolveRealPath(viaLink, out string resolved));
        Assert.Equal(Path.Combine(realTarget, "board.png"), resolved);
    }

    // The link can sit in the MIDDLE of the path, not only at its end - the realistic shape for a
    // containment check, where the caller is resolving "root/link/file.png" as one string.
    [Fact]
    public void A_link_partway_through_a_longer_path_is_still_followed()
    {
        string realTarget = this.thisWorkspace.Path_("RealTarget");
        Directory.CreateDirectory(Path.Combine(realTarget, "sub"));
        File.WriteAllText(Path.Combine(realTarget, "sub", "board.png"), "x");

        string linkPath = this.LinkedDirectory("Linked", realTarget);
        string viaLink = Path.Combine(linkPath, "sub", "board.png");

        Assert.True(RealPathResolver.TryResolveRealPath(viaLink, out string resolved));
        Assert.Equal(Path.Combine(realTarget, "sub", "board.png"), resolved);
    }

    // A link whose target is itself a normal, unlinked path resolves the same as if the link had
    // never existed - proving the resolver does not disturb an ordinary already-real path.
    [Fact]
    public void A_link_pointing_at_an_ordinary_directory_changes_nothing_else()
    {
        string realTarget = this.thisWorkspace.Path_("Ordinary");
        Directory.CreateDirectory(realTarget);

        string linkPath = this.LinkedDirectory("Alias", realTarget);

        Assert.True(RealPathResolver.TryResolveRealPath(linkPath, out string resolved));
        Assert.Equal(realTarget, resolved);
    }

    // -------------------------------------------------------------- separator handling
    //
    // The walk splits on BOTH separators. Path.GetFullPath does NOT convert backslashes on Linux
    // or macOS, so a path carrying the other platform's separator would otherwise be walked as a
    // single giant component: Directory.Exists fails on it, the walk takes the missing-component
    // branch, and no link along it is ever followed - the containment check silently degrading to
    // the purely lexical one this class exists to fix, with no failure reported.
    [Fact]
    public void A_link_reached_through_the_alternate_separator_is_still_followed()
    {
        string realTarget = this.thisWorkspace.Path_("RealTarget");
        Directory.CreateDirectory(realTarget);
        File.WriteAllText(Path.Combine(realTarget, "board.png"), "x");

        string linkPath = this.LinkedDirectory("Linked", realTarget);

        // Built with the ALTERNATE separator between the link and the file it holds.
        string viaLink = linkPath + Path.AltDirectorySeparatorChar + "board.png";

        Assert.True(RealPathResolver.TryResolveRealPath(viaLink, out string resolved));
        Assert.Equal(Path.Combine(realTarget, "board.png"), resolved);
    }

    // The unresolved tail is JOINED, never Path.Combine'd. Combine(params string[]) restarts from
    // the last rooted element, so a remaining segment that Path.IsPathRooted considers rooted
    // would throw the accumulated prefix away and hand back a path shorter than, and unrelated to,
    // the input - about which a caller would then reach a containment verdict.
    [Fact]
    public void An_unresolved_tail_keeps_its_accumulated_prefix()
    {
        string deep = this.thisWorkspace.Path_("missing", "C", "board.png");

        Assert.True(RealPathResolver.TryResolveRealPath(deep, out string resolved));
        Assert.StartsWith(this.thisWorkspace.Root, resolved, StringComparison.Ordinal);
        Assert.Equal(deep, resolved);
    }
}
