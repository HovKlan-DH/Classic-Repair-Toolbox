using Handlers.DataHandling;

namespace CRT.Server.Handlers.Collections
{
    // ###########################################################################################
    // THE CAPTURES' BYTES ON DISK, under the CollectionsRoot setting: "<ab>/<sha256>.png", once each however many collections
    // show them. Named by checksum only - a name is never built from anything else a caller sent, so nothing can lead out of
    // the folder (IsSha256 is asked before any path is made).
    //
    // A file is written beside its place and MOVED in (FileReplacer's rule for served files): a reader never sees half of one,
    // and two uploads of the same picture at once both end with the one file.
    // ###########################################################################################
    public sealed class CollectionFiles
    {
        private readonly string thisRoot;

        public CollectionFiles(string root)
        {
            ArgumentException.ThrowIfNullOrWhiteSpace(root);
            this.thisRoot = Path.GetFullPath(root);
        }

        public string PathOf(string sha256)
        {
            if (!DataTreeLayout.IsSha256(sha256))
                throw new ArgumentException($"[{sha256}] is not a lower-case SHA-256.", nameof(sha256));

            return Path.Combine(this.thisRoot, sha256[..2], sha256 + ".png");
        }

        public bool Has(string sha256) => DataTreeLayout.IsSha256(sha256) && File.Exists(this.PathOf(sha256));

        public void Write(string sha256, ReadOnlySpan<byte> bytes)
        {
            string path = this.PathOf(sha256);

            if (File.Exists(path))
                return;

            Directory.CreateDirectory(Path.GetDirectoryName(path)!);

            string temporary = $"{path}.{Guid.NewGuid():N}.part";

            try
            {
                File.WriteAllBytes(temporary, bytes.ToArray());

                try
                {
                    File.Move(temporary, path, overwrite: false);
                }
                catch (IOException) when (File.Exists(path))
                {
                    // Another upload of the same picture got there first - the same bytes, by their checksum.
                }
            }
            finally
            {
                if (File.Exists(temporary))
                    File.Delete(temporary);
            }
        }

        // The picture's bytes for a page, or null when they are not there.
        public Stream? OpenRead(string sha256)
        {
            if (!this.Has(sha256))
                return null;

            try
            {
                return new FileStream(this.PathOf(sha256), FileMode.Open, FileAccess.Read, FileShare.Read | FileShare.Delete);
            }
            catch (Exception ex) when (ex is FileNotFoundException or DirectoryNotFoundException)
            {
                return null;
            }
        }

        public bool Delete(string sha256)
        {
            if (!this.Has(sha256))
                return false;

            File.Delete(this.PathOf(sha256));
            return true;
        }
    }
}
