using System.Formats.Tar;
using System.IO.Compression;
using System.Text;
using System.Text.Json;
using CRT.Server.Handlers.Updates;

namespace CRT.Server.Tests
{
    // ###########################################################################################
    // Server builds as the release workflow packs them ("tar -czf ... -C publish-server ."), for the
    // updater's tests: real tar.gz bytes, "./" in front of every name, crt-server-release.json at the
    // top, the program, a migration, and the systemd and polkit files every build ships.
    // ###########################################################################################
    internal static class ServerBuilds
    {
        public const string Commit = "7f3a2c1b9d8e7f6a5b4c3d2e1f0a9b8c7d6e5f4a";

        // The systemd and polkit files of a build, as shipped - the same in every build unless a test changes one.
        public static Dictionary<string, string> SystemFiles() => new(StringComparer.Ordinal)
        {
            ["crt-server.service"] = "[Service]\nExecStart=/usr/bin/dotnet app/CRT.Server.dll\n",
            ["crt-server-update.service"] = "[Service]\nUser=crt-deploy\n",
            ["crt-server-update.timer"] = "[Timer]\nOnCalendar=*:0/15\n",
            ["crt-server-test.service"] = "[Service]\nExecStart=/usr/bin/dotnet test/CRT.Server.dll\n",
            ["crt-server-test-update.service"] = "[Service]\nUser=crt-deploy\n",
            ["crt-server-test-update.timer"] = "[Timer]\nOnCalendar=*:0/15\n",
            [ServerUpdater.PolkitRuleFile] = "polkit.addRule(function () {});\n",
        };

        public static string ReleaseInfo(string serverVersion, bool preRelease = false, string crtVersion = "4.1.0", string commit = ServerBuilds.Commit) =>
            JsonSerializer.Serialize(new
            {
                serverVersion,
                crtVersion,
                commit,
                preRelease,
            });

        // A build's files by their names inside it.
        public static Dictionary<string, string> Files(string serverVersion, bool preRelease = false, string commit = ServerBuilds.Commit)
        {
            var files = new Dictionary<string, string>(StringComparer.Ordinal)
            {
                [ServerReleaseInfo.FileName] = ServerBuilds.ReleaseInfo(serverVersion, preRelease, commit: commit),
                ["CRT.Server.dll"] = "the program " + serverVersion,
                ["Migrations/0001_first.sql"] = "CREATE TABLE x (id INT);",
            };

            foreach ((string name, string content) in ServerBuilds.SystemFiles())
                files[name] = content;

            return files;
        }

        // The tar.gz of `files`, every name "./"-prefixed as tar writes it, with the "./" folder and a
        // folder entry before each folder's files.
        public static byte[] Pack(IReadOnlyDictionary<string, string> files)
        {
            using var output = new MemoryStream();

            using (var gzip = new GZipStream(output, CompressionLevel.Fastest, leaveOpen: true))
            using (var writer = new TarWriter(gzip, TarEntryFormat.Pax, leaveOpen: true))
            {
                writer.WriteEntry(new PaxTarEntry(TarEntryType.Directory, "./"));

                var folders = new HashSet<string>(StringComparer.Ordinal);

                foreach ((string name, string content) in files)
                {
                    string? folder = name.Contains('/') ? name[..name.LastIndexOf('/')] : null;

                    if (folder != null && folders.Add(folder))
                        writer.WriteEntry(new PaxTarEntry(TarEntryType.Directory, "./" + folder + "/"));

                    writer.WriteEntry(new PaxTarEntry(TarEntryType.RegularFile, "./" + name)
                    {
                        DataStream = new MemoryStream(Encoding.UTF8.GetBytes(content)),
                    });
                }
            }

            return output.ToArray();
        }

        public static byte[] Build(string serverVersion, bool preRelease = false, string commit = ServerBuilds.Commit) =>
            ServerBuilds.Pack(ServerBuilds.Files(serverVersion, preRelease, commit));

        // Writes a build's files, unpacked, into `folder` - a build already on the box.
        public static void Unpack(string folder, IReadOnlyDictionary<string, string> files)
        {
            foreach ((string name, string content) in files)
            {
                string path = Path.Combine(folder, name.Replace('/', Path.DirectorySeparatorChar));
                Directory.CreateDirectory(Path.GetDirectoryName(path)!);
                File.WriteAllText(path, content);
            }
        }
    }
}
