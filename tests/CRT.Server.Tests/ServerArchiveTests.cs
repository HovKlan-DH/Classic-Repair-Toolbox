using System.Formats.Tar;
using System.IO.Compression;
using System.Text;
using System.Text.RegularExpressions;
using CRT.Server.Handlers.Updates;

namespace CRT.Server.Tests
{
    // ###########################################################################################
    // Unpacking a server build (ServerArchive, 2026-11-09): only ever one whose attestation was checked,
    // and even so unpacked as if anybody made it - nothing lands outside its folder. An archive with a link,
    // a device, a name climbing out or rooted elsewhere is refused whole.
    // ###########################################################################################
    public sealed class ServerArchiveTests : IDisposable
    {
        private readonly string thisRoot = Path.Combine(Path.GetTempPath(), "crt-archive-" + Guid.NewGuid().ToString("N"));

        public ServerArchiveTests() => Directory.CreateDirectory(this.thisRoot);

        public void Dispose()
        {
            try
            {
                Directory.Delete(this.thisRoot, recursive: true);
            }
            catch (IOException)
            {
                // A temp folder left behind is harmless.
            }
        }

        private string Archive(params TarEntry[] entries)
        {
            string path = Path.Combine(this.thisRoot, Guid.NewGuid().ToString("N") + ".tar.gz");

            using (FileStream file = File.Create(path))
            using (var gzip = new GZipStream(file, CompressionLevel.Fastest))
            using (var writer = new TarWriter(gzip, TarEntryFormat.Pax))
            {
                foreach (TarEntry entry in entries)
                    writer.WriteEntry(entry);
            }

            return path;
        }

        private static PaxTarEntry File_(string name, string content = "x") =>
            new(TarEntryType.RegularFile, name) { DataStream = new MemoryStream(Encoding.UTF8.GetBytes(content)) };

        private string Into => Path.Combine(this.thisRoot, "out");

        [Fact]
        public void A_build_as_the_workflow_packs_it_is_unpacked_whole()
        {
            string archive = Path.Combine(this.thisRoot, "build.tar.gz");
            File.WriteAllBytes(archive, ServerBuilds.Build("10.3.0"));

            ServerArchive.Extract(archive, this.Into);

            Assert.Equal("the program 10.3.0", File.ReadAllText(Path.Combine(this.Into, "CRT.Server.dll")));
            Assert.True(File.Exists(Path.Combine(this.Into, "Migrations", "0001_first.sql")));
            Assert.True(File.Exists(Path.Combine(this.Into, ServerReleaseInfo.FileName)));
            Assert.True(File.Exists(Path.Combine(this.Into, "crt-server.service")));
        }

        // ###########################################################################################
        // Each of these would write outside the folder, or put something there that is not a plain file -
        // the whole build is refused.
        // ###########################################################################################
        public static TheoryData<string, TarEntryType, string?> Refused() => new()
        {
            { "../outside.txt", TarEntryType.RegularFile, null },
            { "./../outside.txt", TarEntryType.RegularFile, null },
            { "Migrations/../../outside.txt", TarEntryType.RegularFile, null },
            { "/etc/cron.d/evil", TarEntryType.RegularFile, null },
            { "a\\..\\..\\outside.txt", TarEntryType.RegularFile, null },
            { "link", TarEntryType.SymbolicLink, "/etc/passwd" },
            { "hard", TarEntryType.HardLink, "CRT.Server.dll" },
            { "fifo", TarEntryType.Fifo, null },
        };

        [Theory]
        [MemberData(nameof(ServerArchiveTests.Refused))]
        public void An_entry_that_is_no_plain_file_inside_the_folder_refuses_the_build(string name, TarEntryType type, string? linkName)
        {
            TarEntry entry = type switch
            {
                TarEntryType.RegularFile => ServerArchiveTests.File_(name),
                TarEntryType.SymbolicLink or TarEntryType.HardLink => new PaxTarEntry(type, name) { LinkName = linkName! },
                _ => new PaxTarEntry(type, name),
            };

            string archive = this.Archive(ServerArchiveTests.File_("./CRT.Server.dll"), entry);

            Assert.Throws<InvalidDataException>(() => ServerArchive.Extract(archive, this.Into));

            Assert.False(File.Exists(Path.Combine(this.thisRoot, "outside.txt")));
        }

        [Fact]
        public void A_file_given_twice_refuses_the_build()
        {
            string archive = this.Archive(ServerArchiveTests.File_("./a.txt", "first"), ServerArchiveTests.File_("a.txt", "second"));

            Assert.Throws<InvalidDataException>(() => ServerArchive.Extract(archive, this.Into));
        }

        [Fact]
        public void An_archive_of_more_entries_than_a_build_has_is_refused()
        {
            TarEntry[] entries = Enumerable.Range(0, ServerArchive.MaxEntries + 1)
                .Select(index => (TarEntry)ServerArchiveTests.File_($"f{index}.txt"))
                .ToArray();

            Assert.Throws<InvalidDataException>(() => ServerArchive.Extract(this.Archive(entries), this.Into));
        }

        // ###########################################################################################
        // *** A FOLDER IS GIVEN ITS MODE AS IT IS MADE, NEVER AFTER (the test instance, 2026-11-09). *** Linux
        // takes the setgid bit off a folder re-moded by a user outside its group, silently, and the updater is
        // never in the service's group: re-moded folders made their files in the updater's own group, the service
        // could not read its runtimeconfig (dotnet exit 147) and 10.3.1 went back to 10.3.0. No test user stands
        // outside its own folder's group, so the source holds it: the one mode set after the fact is a file's.
        // ###########################################################################################
        [Fact]
        public void A_folder_is_given_its_mode_as_it_is_made_and_never_after()
        {
            string source = File.ReadAllText(
                Path.Combine(ServerRouteTable.RepositoryRoot(), "src", "CRT.Server", "Handlers", "Updates", "ServerArchive.cs"));

            List<string> remoded = Regex.Matches(source, @"SetUnixFileMode\(\s*([^,]+),")
                .Select(match => match.Groups[1].Value.Trim())
                .ToList();

            Assert.Equal(["file"], remoded);
        }

        // And as unpacked on Linux under a setgid folder, as the releases folder is: every folder keeps the bit
        // (so the service's group), 0750 apart from it; every file 0640. The CI runner checks it.
        [Fact]
        public void Unpacked_under_a_setgid_folder_every_folder_keeps_the_bit_and_files_are_0640()
        {
            // An explicit guard, which the platform analyzer (CA1416) understands.
            if (OperatingSystem.IsWindows())
            {
                Assert.Skip("Unix file modes - checked on the Linux CI runner.");
                return;
            }

            string releases = Path.Combine(this.thisRoot, "releases");
            Directory.CreateDirectory(releases);
            File.SetUnixFileMode(releases, ServerArchive.FolderMode | UnixFileMode.SetGroup);

            string archive = Path.Combine(this.thisRoot, "build.tar.gz");
            File.WriteAllBytes(archive, ServerBuilds.Build("10.3.0"));
            string build = Path.Combine(releases, "build");

            ServerArchive.Extract(archive, build);

            const UnixFileMode others = UnixFileMode.OtherRead | UnixFileMode.OtherWrite | UnixFileMode.OtherExecute;

            foreach (string folder in new[] { build, Path.Combine(build, "Migrations") })
            {
                UnixFileMode mode = File.GetUnixFileMode(folder);
                Assert.True(mode.HasFlag(UnixFileMode.SetGroup), $"[{folder}] lost its setgid bit: {mode}");
                Assert.True(mode.HasFlag(UnixFileMode.GroupRead | UnixFileMode.GroupExecute), $"[{folder}]: {mode}");
                Assert.Equal((UnixFileMode)0, mode & others);
            }

            Assert.Equal(ServerArchive.FileWrittenMode, File.GetUnixFileMode(Path.Combine(build, "CRT.Server.dll")));
            Assert.Equal(ServerArchive.FileWrittenMode, File.GetUnixFileMode(Path.Combine(build, "Migrations", "0001_first.sql")));
        }

        [Theory]
        [InlineData("./CRT.Server.dll", "CRT.Server.dll")]
        [InlineData("CRT.Server.dll", "CRT.Server.dll")]
        [InlineData("./Migrations/0001.sql", "Migrations/0001.sql")]
        [InlineData("./Migrations/", "Migrations")]
        [InlineData("./", null)]
        [InlineData("", null)]
        [InlineData("../x", null)]
        [InlineData("a/./b", null)]
        [InlineData("a//b", null)]
        [InlineData("/x", null)]
        [InlineData("a\\b", null)]
        [InlineData("C:x", null)]
        [InlineData("a\nb", null)]
        public void A_name_is_a_plain_path_inside_the_folder_or_none(string name, string? expected) =>
            Assert.Equal(expected?.Replace('/', Path.DirectorySeparatorChar), ServerArchive.RelativeName(name));
    }
}
