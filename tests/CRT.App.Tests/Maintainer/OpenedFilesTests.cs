using Handlers.MaintainerHandling;

namespace ClassicRepairToolbox.Tests.Maintainer;

// ###########################################################################################
// Covers OpenedFiles' names - what a file tree and the table's hover card may hand to the
// operating system, and under which name.
// ###########################################################################################
public sealed class OpenedFilesTests
{
    // ###########################################################################################
    // What the tree opens: what a submission may carry, and a board's own files - workbook,
    // highlight file, KiCad data. A web page as text. Nothing else, and only the file's own name.
    // ###########################################################################################
    [Theory]
    [InlineData("Commodore/C128/310378/Data C128 310378 v2.0.0.xlsx", "Data C128 310378 v2.0.0.xlsx")]
    [InlineData("Commodore/C128/310378/Data C128 310378 v2.0.0.json", "Data C128 310378 v2.0.0.json")]
    [InlineData("Commodore/C128/310378/KiCad data/c128.kicad_pcb", "c128.kicad_pcb")]
    [InlineData("Commodore/C128/310378/Images/a.png", "a.png")]
    [InlineData("Commodore/C128/310378/notes.html", "notes.html.txt")]
    public void A_board_file_opens_under_its_own_name(string path, string expected)
    {
        Assert.True(OpenedFiles.TryGetOpenName(path, out string name));
        Assert.Equal(expected, name);
    }

    [Theory]
    [InlineData("Commodore/C128/310378/run.exe")]
    [InlineData("Commodore/C128/310378/macro.xlsm")]
    [InlineData("Commodore/C128/310378/")]
    public void Anything_else_is_not_opened(string path)
    {
        Assert.False(OpenedFiles.TryGetOpenName(path, out _));
    }

    // ###########################################################################################
    // *** IN THE USER'S OWN APP DATA FOLDER, NEVER THE SHARED TEMP ONE (code review, 2026-11-02 -
    // Security). *** Under a predictable folder in /tmp another account could make the folder first,
    // read a contributor's files in it, or swap one between the write and the launch. The rule the
    // scope captures and the KiCad cache keep.
    // ###########################################################################################
    [Fact]
    public void Opened_files_are_written_in_the_users_own_app_data_folder()
    {
        string expected = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            CRT.AppConfig.AppFolderName,
            "Maintainer opened files");

        Assert.Equal(expected, OpenedFiles.Root);
        Assert.False(
            Path.GetFullPath(OpenedFiles.Root).StartsWith(Path.GetFullPath(Path.GetTempPath()), StringComparison.OrdinalIgnoreCase),
            "Opened files must not lie under the shared temp folder.");
    }

    // ###########################################################################################
    // What goes where the operating system can read it, and the sweep of what earlier opens left: a
    // folder of its own per open, one older than a day deleted, a newer one left - its viewer may
    // still be reading it. Written under a test's own folder, never the user's real one.
    // ###########################################################################################
    [Fact]
    public async Task An_open_saves_into_a_folder_of_its_own_and_sweeps_what_is_a_day_old()
    {
        string root = Path.Combine(Path.GetTempPath(), "crt-opened-files-" + Guid.NewGuid().ToString("N"));
        string stale = Path.Combine(root, "stale");
        string recent = Path.Combine(root, "recent");

        Directory.CreateDirectory(stale);
        Directory.CreateDirectory(recent);
        File.WriteAllText(Path.Combine(stale, "old.pdf"), "old");
        Directory.SetLastWriteTimeUtc(stale, DateTime.UtcNow - ReviewTableFiles.OpenedFileLifetime - TimeSpan.FromHours(1));

        try
        {
            string? launched = null;

            string? said = await OpenedFiles.SaveAndLaunchInAsync(root, [1, 2, 3], "a.pdf", path =>
            {
                launched = path;
                return Task.FromResult(true);
            });

            Assert.Null(said);
            Assert.Equal("a.pdf", Path.GetFileName(launched));
            Assert.Equal([1, 2, 3], File.ReadAllBytes(launched!));
            Assert.Equal(root, Path.GetDirectoryName(Path.GetDirectoryName(launched)));

            Assert.False(Directory.Exists(stale));
            Assert.True(Directory.Exists(recent));
        }
        finally
        {
            Directory.Delete(root, recursive: true);
        }
    }

    // ###########################################################################################
    // *** WHAT AN EARLIER CRT LEFT IN THE SHARED TEMP FOLDER GOES TOO (code review, 2026-11-02 - Security). ***
    // Root moved, and contributors' files stayed under "<temp>/CRT/Maintainer" for good, readable on a
    // shared Linux machine. Only what an open would have made goes: a folder named by a GUID, old enough
    // for no viewer to hold it. That folder may be another account's, so nothing else is touched.
    // ###########################################################################################
    private static (string Temp, string Earlier) EarlierFolders()
    {
        string temp = Path.Combine(Path.GetTempPath(), "crt-opened-files-" + Guid.NewGuid().ToString("N"));
        string earlier = Path.Combine(temp, "CRT", "Maintainer");
        Directory.CreateDirectory(earlier);

        return (temp, earlier);
    }

    private static string FolderIn(string parent, string name, bool old)
    {
        string folder = Path.Combine(parent, name);
        Directory.CreateDirectory(folder);
        File.WriteAllText(Path.Combine(folder, "a.pdf"), "a contributor's file");

        if (old)
            Directory.SetLastWriteTimeUtc(folder, DateTime.UtcNow - ReviewTableFiles.OpenedFileLifetime - TimeSpan.FromHours(1));

        return folder;
    }

    [Fact]
    public void An_open_clears_what_an_earlier_CRT_left_in_the_shared_temp_folder_and_nothing_else()
    {
        (string temp, string earlier) = EarlierFolders();

        try
        {
            string left = FolderIn(earlier, Guid.NewGuid().ToString("N"), old: true);
            string recent = FolderIn(earlier, Guid.NewGuid().ToString("N"), old: false);
            string notAnOpens = FolderIn(earlier, "Somebody's", old: true);

            OpenedFiles.SweepEarlierRoot(temp, earlier, DateTimeOffset.UtcNow);

            Assert.False(Directory.Exists(left));
            Assert.True(File.Exists(Path.Combine(recent, "a.pdf")));
            Assert.True(File.Exists(Path.Combine(notAnOpens, "a.pdf")));
        }
        finally
        {
            Directory.Delete(temp, recursive: true);
        }
    }

    // Once nothing is left it goes altogether, so every later open costs one Directory.Exists.
    [Fact]
    public void An_earlier_folder_left_empty_goes_and_a_missing_one_is_nothing_to_do()
    {
        (string temp, string earlier) = EarlierFolders();

        try
        {
            FolderIn(earlier, Guid.NewGuid().ToString("N"), old: true);

            OpenedFiles.SweepEarlierRoot(temp, earlier, DateTimeOffset.UtcNow);

            Assert.False(Directory.Exists(Path.GetDirectoryName(earlier)));

            OpenedFiles.SweepEarlierRoot(temp, earlier, DateTimeOffset.UtcNow);
        }
        finally
        {
            Directory.Delete(temp, recursive: true);
        }
    }

    // A folder under <temp> can be another account's link to the user's own: nothing is deleted through it.
    [Fact]
    public void Nothing_is_deleted_through_a_temp_folder_that_links_elsewhere()
    {
        (string temp, string earlier) = EarlierFolders();
        string elsewhere = Path.Combine(temp, "The user's own");
        Directory.CreateDirectory(elsewhere);

        try
        {
            string kept = FolderIn(elsewhere, Guid.NewGuid().ToString("N"), old: true);

            Directory.Delete(earlier);

            try
            {
                Directory.CreateSymbolicLink(earlier, elsewhere);
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                Assert.Skip("This account may not create a folder link (Windows outside Developer Mode).");
            }

            OpenedFiles.SweepEarlierRoot(temp, earlier, DateTimeOffset.UtcNow);

            Assert.True(File.Exists(Path.Combine(kept, "a.pdf")));
        }
        finally
        {
            if (Directory.Exists(earlier))
                Directory.Delete(earlier);

            Directory.Delete(temp, recursive: true);
        }
    }

    // What is saved is what the launcher takes: inside the root it is handed, on its list, and real.
    // By its private predicate, as ExternalTargetLauncherTests does - TryOpen's accepting path starts
    // a process (test rule 6). A path outside the root is refused through TryOpen itself.
    [Fact]
    public async Task What_is_saved_is_what_the_launcher_takes_and_nothing_outside_its_folder_is()
    {
        string root = Path.GetFullPath(Path.Combine(Path.GetTempPath(), "crt-opened-files-" + Guid.NewGuid().ToString("N")));
        string outside = root + "-outside.pdf";

        try
        {
            File.WriteAllText(outside, "not opened");

            await OpenedFiles.SaveAndLaunchInAsync(root, [1], "a.pdf", path =>
            {
                object?[] arguments = [path, root, null, true];

                bool taken = (bool)typeof(CRT.ExternalTargetLauncher)
                    .GetMethod("TryResolveDataRootScopedFilePath", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static)!
                    .Invoke(null, arguments)!;

                Assert.True(taken, path);
                return Task.FromResult(true);
            });

            Assert.False(CRT.ExternalTargetLauncher.TryOpen(outside, root));
        }
        finally
        {
            Directory.Delete(root, recursive: true);
            File.Delete(outside);
        }
    }

    // A KiCad project file opens as what it is, JSON - alone in its folder KiCad would open an empty
    // project, and the launcher has no reason to take the type.
    [Fact]
    public void A_KiCad_project_file_opens_as_JSON()
    {
        Assert.True(OpenedFiles.TryGetOpenName("Commodore/C128/310378/KiCad data/c128.kicad_pro", out string name));
        Assert.Equal("c128.kicad_pro.json", name);
    }

    // ###########################################################################################
    // *** TWO LISTS, ONE ANSWER (code review, 2026-11-02). *** What the Maintainer tab opens goes
    // through ExternalTargetLauncher, whose own list decides - so a type TryGetOpenName lets through
    // and the launcher refuses would be a file that downloads and then "could not be opened". A KiCad
    // project file is the one that differed. Every type on either of the tab's lists is asked here.
    // ###########################################################################################
    [Fact]
    public void Every_type_the_tab_opens_is_one_the_launcher_opens()
    {
        IEnumerable<string> types = Handlers.DataHandling.SubmissionFileRules.AllowedExtensions
            .Concat([".xlsx", ".json", ".kicad_pcb", ".kicad_sch", ".kicad_pro"]);

        Assert.All(types, type =>
        {
            Assert.True(OpenedFiles.TryGetOpenName("Commodore/C64/250407/a" + type, out string name), type);
            Assert.True(CRT.ExternalTargetLauncher.IsOpenableFile(name), $"{type} is saved as {name}, which the launcher refuses.");
        });
    }
}
