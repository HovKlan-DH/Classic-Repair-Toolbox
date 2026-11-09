using System.Formats.Tar;
using System.IO.Compression;

namespace CRT.Server.Handlers.Updates
{
    // ###########################################################################################
    // UNPACKING A SERVER BUILD (2026-11-09). Only ever a build whose attestation was checked first
    // (ServerUpdater), and even so unpacked as if it came from anybody: what is written stays inside
    // the folder it is unpacked into.
    //
    //   - Regular files and folders only. A link, a device or anything else in the archive refuses the
    //     whole build - a server build has none, and a link is how an archive writes outside its folder.
    //   - A name is relative, uses "/" only, and has no "." or ".." part ("./" in front, as tar writes it
    //     for "tar -C dir .", is taken off first).
    //   - At most MaxEntries entries and MaxBytes unpacked - an archive that grows past either is refused.
    //   - Files are written readable by the service's group and nobody else (0640, folders 0750), owned by
    //     whoever unpacks it - the updater's user, never the service, which cannot rewrite its own program.
    //   - *** A FOLDER IS GIVEN ITS MODE AS IT IS MADE, NEVER AFTER. *** The releases folder is setgid to the
    //     service's group, so what is made under it is in that group though the updater is not. But Linux
    //     takes the setgid bit off a folder whose mode is changed by a user outside its group - silently - and
    //     the updater is always outside it: each folder it re-moded made its files in the updater's own group,
    //     and the service could not read its own CRT.Server.runtimeconfig.json (dotnet's exit 147, 0x80008093),
    //     so the first build the test instance's updater unpacked, 10.3.1, did not come up (2026-11-09).
    //
    // Refused, it throws InvalidDataException and leaves what it wrote for the caller to delete.
    // ###########################################################################################
    internal static class ServerArchive
    {
        public const int MaxEntries = 5_000;
        public const long MaxBytes = 600L * 1024 * 1024;

        // The mode a folder is made with - the setgid bit comes from the folder it is made in.
        internal const UnixFileMode FolderMode =
            UnixFileMode.UserRead | UnixFileMode.UserWrite | UnixFileMode.UserExecute | UnixFileMode.GroupRead | UnixFileMode.GroupExecute;

        // The mode a file is given once written.
        internal const UnixFileMode FileWrittenMode = UnixFileMode.UserRead | UnixFileMode.UserWrite | UnixFileMode.GroupRead;

        public static void Extract(string archivePath, string destinationFolder)
        {
            string root = Path.GetFullPath(destinationFolder);
            ServerArchive.CreateFolder(root);

            using FileStream file = File.OpenRead(archivePath);
            using var gzip = new GZipStream(file, CompressionMode.Decompress);
            using var reader = new TarReader(gzip);

            int entries = 0;
            long bytes = 0;

            while (reader.GetNextEntry() is TarEntry entry)
            {
                if (++entries > ServerArchive.MaxEntries)
                    throw new InvalidDataException($"The build holds more than {ServerArchive.MaxEntries} entries.");

                string? name = ServerArchive.RelativeName(entry.Name);

                if (name == null)
                {
                    // The archive's own top folder ("./").
                    if (ServerArchive.IsTopFolder(entry.Name) && entry.EntryType == TarEntryType.Directory)
                        continue;

                    throw new InvalidDataException($"The build holds an entry named [{entry.Name}], which is no name inside its folder.");
                }

                string target = Path.GetFullPath(Path.Combine(root, name));

                if (!target.StartsWith(root + Path.DirectorySeparatorChar, StringComparison.Ordinal))
                    throw new InvalidDataException($"The build's entry [{entry.Name}] would land outside its folder.");

                switch (entry.EntryType)
                {
                    case TarEntryType.Directory:
                        ServerArchive.MakeFolder(root, target);
                        break;

                    case TarEntryType.RegularFile:
                    case TarEntryType.V7RegularFile:
                    case TarEntryType.ContiguousFile:
                        bytes += entry.Length;

                        if (entry.Length < 0 || bytes > ServerArchive.MaxBytes)
                            throw new InvalidDataException($"The build unpacks to more than {ServerArchive.MaxBytes / 1024 / 1024} MB.");

                        ServerArchive.MakeFolder(root, Path.GetDirectoryName(target)!);

                        if (File.Exists(target) || Directory.Exists(target))
                            throw new InvalidDataException($"The build holds [{entry.Name}] twice.");

                        using (FileStream output = new(target, FileMode.CreateNew, FileAccess.Write))
                        {
                            entry.DataStream?.CopyTo(output);
                        }

                        ServerArchive.SetFileMode(target);
                        break;

                    default:
                        throw new InvalidDataException($"The build's entry [{entry.Name}] is a {entry.EntryType}, not a file or a folder.");
                }
            }
        }

        // ###########################################################################################
        // `name` as a path inside the folder ("Migrations/0001.sql"), or null when it is none: empty, rooted,
        // a "\", a control character, or a "." or ".." part. A leading "./" is taken off.
        // ###########################################################################################
        public static string? RelativeName(string? name)
        {
            if (string.IsNullOrEmpty(name))
                return null;

            string trimmed = name;

            while (trimmed.StartsWith("./", StringComparison.Ordinal))
                trimmed = trimmed[2..];

            trimmed = trimmed.TrimEnd('/');

            if (trimmed.Length == 0 ||
                trimmed.StartsWith('/') ||
                trimmed.Contains('\\') ||
                trimmed.Any(char.IsControl) ||
                trimmed.Contains(':'))
            {
                return null;
            }

            string[] parts = trimmed.Split('/');

            return parts.Any(part => part.Length == 0 || part == "." || part == "..")
                ? null
                : string.Join(Path.DirectorySeparatorChar, parts);
        }

        private static bool IsTopFolder(string? name) =>
            name is "." or "./" ;

        private static void MakeFolder(string root, string folder)
        {
            if (string.Equals(folder, root, StringComparison.Ordinal))
                return;

            if (File.Exists(folder))
                throw new InvalidDataException($"The build holds [{folder}] as a file and as a folder.");

            if (!Directory.Exists(folder))
            {
                ServerArchive.MakeFolder(root, Path.GetDirectoryName(folder)!);
                ServerArchive.CreateFolder(folder);
            }
        }

        // A folder made with its mode in the one mkdir - never changed after, which would cost it its setgid bit.
        private static void CreateFolder(string folder)
        {
            if (OperatingSystem.IsWindows())
                Directory.CreateDirectory(folder);
            else
                Directory.CreateDirectory(folder, ServerArchive.FolderMode);
        }

        // A file the updater wrote: 0640. A file has no setgid bit to lose; its group came from its folder.
        private static void SetFileMode(string file)
        {
            if (!OperatingSystem.IsWindows())
                File.SetUnixFileMode(file, ServerArchive.FileWrittenMode);
        }
    }
}
