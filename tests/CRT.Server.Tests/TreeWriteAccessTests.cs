using CRT.Server.Handlers.Submissions;
using Xunit;

namespace CRT.Server.Tests
{
    // ###########################################################################################
    // Covers TreeWriteAccess - "may the service write where it is about to?", asked before a
    // publish, a production promotion or a push-back writes anything (owner report, 2026-09-28:
    // production's files copied into BETA by hand as root, then an approval refused half-way
    // through the board).
    //
    // Which folders are asked is tested through an injected probe - a folder the service may not
    // write cannot be made on every OS this suite runs on. The real probe is tested where it can
    // be: yes on a writable folder everywhere, no on a read-only one off Windows.
    // ###########################################################################################
    public sealed class TreeWriteAccessTests : IDisposable
    {
        private readonly string thisRoot;

        public TreeWriteAccessTests()
        {
            this.thisRoot = Path.Combine(Path.GetTempPath(), "crt-tree-access", Guid.NewGuid().ToString("N"), "Data");
            Directory.CreateDirectory(this.thisRoot);
        }

        public void Dispose()
        {
            try
            {
                string parent = Path.GetDirectoryName(this.thisRoot)!;

                foreach (string folder in Directory.GetDirectories(parent, "*", SearchOption.AllDirectories))
                {
                    if (!OperatingSystem.IsWindows())
                        File.SetUnixFileMode(folder, UnixFileMode.UserRead | UnixFileMode.UserWrite | UnixFileMode.UserExecute);
                }

                Directory.Delete(parent, recursive: true);
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                // A leftover temp folder is harmless.
            }
        }

        private string Folder(string relative)
        {
            string full = Path.Combine(this.thisRoot, relative.Replace('/', Path.DirectorySeparatorChar));
            Directory.CreateDirectory(full);
            return full;
        }

        private string PathIn(string relative) =>
            Path.Combine(this.thisRoot, relative.Replace('/', Path.DirectorySeparatorChar));

        // Every folder asked ONCE however many files go into it, and only the refusing ones named.
        [Fact]
        public void Each_folder_is_asked_once_and_only_the_refusing_ones_are_named()
        {
            string board = this.Folder("Commodore/C128/310378");
            string shared = this.Folder("Commodore/Shared files");
            var asked = new List<string>();

            IReadOnlyList<string> refusing = TreeWriteAccess.FoldersRefusing(
                this.thisRoot,
                [
                    this.PathIn("Commodore/C128/310378/Data C128 310378 v2.0.0.xlsx"),
                    this.PathIn("Commodore/C128/310378/Data C128 310378 v2.0.0.json"),
                    this.PathIn("Commodore/Shared files/a.pdf")
                ],
                folder =>
                {
                    asked.Add(folder);
                    return folder != board;
                });

            Assert.Equal([board, shared], asked);
            Assert.Equal([board], refusing);
        }

        // A folder that is not there yet is created by the write, so the folder that must allow it
        // is the nearest one above it that exists.
        [Fact]
        public void A_folder_not_there_yet_is_judged_by_the_nearest_folder_above_it()
        {
            string board = this.Folder("Commodore/C128/310378");
            var asked = new List<string>();

            TreeWriteAccess.FoldersRefusing(
                this.thisRoot,
                [this.PathIn("Commodore/C128/310378/Images/New/sheet.png")],
                folder =>
                {
                    asked.Add(folder);
                    return true;
                });

            Assert.Equal([board], asked);
        }

        // ###########################################################################################
        // *** NOTHING OUTSIDE THE TREE IS EVER PROBED. *** A probe is a real file; walking up past
        // the root would create one in the web folder above it. A root that is not there is reported
        // as refusing without being asked, and a path outside the root is not the check's business -
        // callers resolved every path inside it already.
        // ###########################################################################################
        [Fact]
        public void Nothing_outside_the_tree_is_ever_probed()
        {
            string missingRoot = Path.Combine(this.thisRoot, "not-there");
            var asked = new List<string>();

            IReadOnlyList<string> refusing = TreeWriteAccess.FoldersRefusing(
                missingRoot,
                [Path.Combine(missingRoot, "Commodore", "C64", "a.png"), Path.Combine(this.thisRoot, "elsewhere.png")],
                folder =>
                {
                    asked.Add(folder);
                    return true;
                });

            Assert.Empty(asked);
            Assert.Equal([missingRoot], refusing);
        }

        [Fact]
        public void With_every_folder_writable_nothing_is_refused()
        {
            this.Folder("Commodore/C64/250407");

            Assert.Empty(TreeWriteAccess.FoldersRefusing(
                this.thisRoot, [this.PathIn("Commodore/C64/250407/a.png")], _ => true));
        }

        // The real probe says yes to a folder it may write, and leaves nothing behind in it.
        [Fact]
        public void The_real_probe_accepts_a_writable_folder_and_leaves_no_file_behind()
        {
            string board = this.Folder("Commodore/C64/250407");

            Assert.Empty(TreeWriteAccess.FoldersRefusing(this.thisRoot, [this.PathIn("Commodore/C64/250407/a.png")]));
            Assert.Empty(Directory.GetFileSystemEntries(board));
        }

        // ###########################################################################################
        // The real probe says NO to a folder the service may not write - a folder copied in by hand
        // as root. Made here as a folder with no write permission for its owner, which refuses a
        // non-root process the same way. Not on Windows (no such mode), and not as root (which
        // ignores it).
        // ###########################################################################################
        [Fact]
        public void The_real_probe_refuses_a_folder_that_may_not_be_written()
        {
            // The return is for the platform analyzer, which cannot see that Skip throws.
            if (OperatingSystem.IsWindows())
            {
                Assert.Skip("Unix permissions only.");
                return;
            }
            Assert.SkipWhen(Environment.UserName == "root", "root may write anywhere.");

            string board = this.Folder("Commodore/C128/310378");
            File.SetUnixFileMode(board, UnixFileMode.UserRead | UnixFileMode.UserExecute);

            Assert.Equal([board], TreeWriteAccess.FoldersRefusing(this.thisRoot, [this.PathIn("Commodore/C128/310378/a.json")]));
        }

        // The maintainer is told which folders, the way the data names them, and that nothing
        // changed - not the command, which is the administrator's and goes to the log.
        [Fact]
        public void The_refusal_names_the_folders_as_the_data_does_and_says_nothing_was_changed()
        {
            string message = TreeWriteAccess.RefusalMessage(
                "BETA", this.thisRoot, [this.PathIn("Commodore/C128/310378"), this.thisRoot]);

            Assert.Contains("[Commodore/C128/310378]", message, StringComparison.Ordinal);
            Assert.Contains("[the top folder]", message, StringComparison.Ordinal);
            Assert.Contains("in the BETA data, so nothing was changed", message, StringComparison.Ordinal);
            Assert.DoesNotContain("sudo", message, StringComparison.Ordinal);
            Assert.DoesNotContain(this.thisRoot, message, StringComparison.Ordinal);
        }

        [Fact]
        public void A_long_list_of_folders_is_cut_short()
        {
            IReadOnlyList<string> folders = Enumerable.Range(1, 8).Select(index => this.PathIn($"F{index}")).ToList();

            string message = TreeWriteAccess.RefusalMessage("production", this.thisRoot, folders);

            Assert.Contains("[F5] and 3 more", message, StringComparison.Ordinal);
            Assert.DoesNotContain("[F6]", message, StringComparison.Ordinal);
        }

        // ###########################################################################################
        // *** A REFUSAL NAMES A PATH FROM THE DATA'S ROOT, NEVER THE SERVER'S OWN (code review,
        // 2026-10-13). *** A symbolic link refusal printed the absolute path on the server to every
        // maintainer it refused. Every such name goes through DisplayName now; a path outside the
        // root (which no refusal should meet) is named by its last part only.
        // ###########################################################################################
        [Fact]
        public void A_path_is_named_from_the_datas_root_and_never_as_the_servers_own()
        {
            string link = Path.Combine(this.thisRoot, "Commodore", "C64", "250407", "Images");

            Assert.Equal("Commodore/C64/250407/Images", TreeWriteAccess.DisplayName(this.thisRoot, link));
            Assert.Equal("the top folder", TreeWriteAccess.DisplayName(this.thisRoot, this.thisRoot));

            string outside = Path.Combine(Path.GetTempPath(), "elsewhere", "secret");
            Assert.Equal("secret", TreeWriteAccess.DisplayName(this.thisRoot, outside));
        }

        [Fact]
        public void A_path_under_either_of_two_trees_is_named_from_its_own()
        {
            string stable = Path.Combine(Path.GetDirectoryName(this.thisRoot)!, "Stable");
            string inStable = Path.Combine(stable, "Commodore", "Shared files");

            Assert.Equal("Commodore/Shared files", TreeWriteAccess.DisplayName(this.thisRoot, stable, inStable));
            Assert.Equal("Amstrad", TreeWriteAccess.DisplayName(this.thisRoot, stable, Path.Combine(this.thisRoot, "Amstrad")));
        }

        // ###########################################################################################
        // The command is INSTALLING.md ("Folders and permissions") for exactly these folders - group, group write, and
        // setgid on every folder so what is written later keeps the group. Every path quoted for the
        // shell: board folders and "Shared files" have spaces, and a quote must not end the string.
        // ###########################################################################################
        [Fact]
        public void The_fix_command_is_step_3_for_each_folder_quoted_for_the_shell()
        {
            string command = TreeWriteAccess.FixCommand(["/srv/Data/Commodore/Shared files", "/srv/Data/O'Brien"]);

            Assert.Equal(
                "sudo chgrp -R crt-data '/srv/Data/Commodore/Shared files' && sudo chmod -R g+rwX '/srv/Data/Commodore/Shared files' && " +
                "sudo find '/srv/Data/Commodore/Shared files' -type d -exec chmod g+s {} + ; " +
                "sudo chgrp -R crt-data '/srv/Data/O'\\''Brien' && sudo chmod -R g+rwX '/srv/Data/O'\\''Brien' && " +
                "sudo find '/srv/Data/O'\\''Brien' -type d -exec chmod g+s {} +",
                command);
        }
    }
}
