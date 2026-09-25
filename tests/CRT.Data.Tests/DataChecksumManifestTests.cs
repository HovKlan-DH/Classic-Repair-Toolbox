using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// Building dataChecksums.json - the file every CRT client syncs against (owner request,
// 2026-09-23).
//
// *** WHY THIS IS WORTH REAL COVERAGE. *** Every installed copy of CRT reads this file, including
// versions that will never be updated. Three ways to get it wrong are all silent from the
// server's side:
//
//   - a WRONG CHECKSUM means a client never downloads a file that changed, which is precisely
//     the bug that prompted this (the board was published, the manifest still advertised the old
//     hash, and every sync completed successfully having done nothing);
//   - a WRONG URL means a download 404s, and board folders have spaces in their names;
//   - a RENAMED PROPERTY means older clients read the whole manifest as empty.
//
// Real files in a real temp folder throughout - this hashes bytes, so there is nothing to fake.
// ###########################################################################################
public sealed class DataChecksumManifestTests : IDisposable
{
    private readonly TempWorkspace thisWorkspace = new();

    private string DataRoot => Path.Combine(this.thisWorkspace.Root, "Data");

    private string ManifestPath => Path.Combine(this.thisWorkspace.Root, "dataChecksums.json");

    private const string BaseUrl = "https://classic-repair-toolbox.dk/app-data-BETA/Data";

    public void Dispose() => this.thisWorkspace.Dispose();

    private string Write(string relativePath, string content)
    {
        string full = Path.Combine(
            this.DataRoot,
            relativePath.Replace('/', Path.DirectorySeparatorChar));

        Directory.CreateDirectory(Path.GetDirectoryName(full)!);
        File.WriteAllText(full, content);

        return full;
    }

    private static string Sha256Of(string content) =>
        Convert.ToHexStringLower(SHA256.HashData(Encoding.UTF8.GetBytes(content)));

    // ------------------------------------------------------------------ the checksum

    [Fact]
    public void The_checksum_is_the_lowercase_SHA256_of_the_file()
    {
        this.Write("Commodore/C64/250407/Data.xlsx", "board bytes");

        DataChecksumManifest.Entry entry = Assert.Single(
            DataChecksumManifest.Scan(this.DataRoot, DataChecksumManifestTests.BaseUrl));

        Assert.Equal(DataChecksumManifestTests.Sha256Of("board bytes"), entry.Checksum);

        // Lowercase matters: a hash differing only in case compares unequal on the client, which
        // reads as the file being permanently out of date.
        Assert.Equal(entry.Checksum, entry.Checksum.ToLowerInvariant());
    }

    // ###########################################################################################
    // *** THE REPORTED BUG, AS A TEST. *** Publishing rewrote a board and the manifest went on
    // advertising the old checksum, so no client ever fetched it. Regenerating has to produce the
    // NEW hash for a file whose contents changed.
    // ###########################################################################################
    [Fact]
    public void A_REPUBLISHED_file_gets_a_new_checksum()
    {
        string path = this.Write("Commodore/C64/250407/Data.xlsx", "before the publish");

        string before = Assert.Single(
            DataChecksumManifest.Scan(this.DataRoot, DataChecksumManifestTests.BaseUrl)).Checksum;

        File.WriteAllText(path, "after the publish");

        string after = Assert.Single(
            DataChecksumManifest.Scan(this.DataRoot, DataChecksumManifestTests.BaseUrl)).Checksum;

        Assert.NotEqual(before, after);
        Assert.Equal(DataChecksumManifestTests.Sha256Of("after the publish"), after);
    }

    // ------------------------------------------------------------------ the path and the url

    [Fact]
    public void The_file_path_is_data_root_relative_with_forward_slashes()
    {
        this.Write("Commodore/C64/250407/Data.xlsx", "x");

        DataChecksumManifest.Entry entry = Assert.Single(
            DataChecksumManifest.Scan(this.DataRoot, DataChecksumManifestTests.BaseUrl));

        // Forward slashes even on Windows: the client matches this against its own relative paths
        // and the server serves it over HTTP.
        Assert.Equal("Commodore/C64/250407/Data.xlsx", entry.File);
        Assert.DoesNotContain('\\', entry.File);
    }

    // ###########################################################################################
    // *** EACH SEGMENT IS ESCAPED SEPARATELY, and getting this wrong breaks every download. ***
    // Board folders genuinely contain spaces ("Commodore/Shared files/..."). Escaping the whole
    // relative path in one go turns its slashes into %2F and the URL resolves to nothing.
    // ###########################################################################################
    [Fact]
    public void A_url_escapes_each_segment_but_keeps_the_slashes()
    {
        this.Write("Commodore/Shared files/6526 pinout.png", "x");

        DataChecksumManifest.Entry entry = Assert.Single(
            DataChecksumManifest.Scan(this.DataRoot, DataChecksumManifestTests.BaseUrl));

        Assert.Equal(
            DataChecksumManifestTests.BaseUrl + "/Commodore/Shared%20files/6526%20pinout.png",
            entry.Url);
    }

    [Fact]
    public void A_trailing_slash_on_the_base_url_is_not_doubled()
    {
        this.Write("a.txt", "x");

        DataChecksumManifest.Entry entry = Assert.Single(
            DataChecksumManifest.Scan(this.DataRoot, DataChecksumManifestTests.BaseUrl + "/"));

        Assert.DoesNotContain("//a.txt", entry.Url, StringComparison.Ordinal);
    }

    // ------------------------------------------------------------------ what is skipped

    [Theory]
    [InlineData("Thumbs.db")]
    [InlineData("desktop.ini")]
    [InlineData(".hidden")]
    [InlineData("Commodore/.git/config")]
    [InlineData("Commodore/C64/Data.xlsx.tmp_1758648000")]
    public void OS_junk_dot_paths_and_temp_files_are_left_out(string relativePath)
    {
        this.Write("Commodore/C64/250407/Data.xlsx", "real");
        this.Write(relativePath, "junk");

        IReadOnlyList<DataChecksumManifest.Entry> entries =
            DataChecksumManifest.Scan(this.DataRoot, DataChecksumManifestTests.BaseUrl);

        // The real file is still there, so this proves exclusion rather than a broken scan.
        DataChecksumManifest.Entry entry = Assert.Single(entries);
        Assert.Equal("Commodore/C64/250407/Data.xlsx", entry.File);
    }

    [Fact]
    public void A_dot_FOLDER_excludes_everything_beneath_it()
    {
        // Checked segment by segment, so it is the ".git" that excludes the file rather than
        // anything about the file's own name.
        Assert.False(DataChecksumManifest.IsSyncable("Commodore/.git/objects/abc"));
        Assert.True(DataChecksumManifest.IsSyncable("Commodore/C64/250407/Data.xlsx"));
    }

    // ------------------------------------------------------------------ ordering

    [Fact]
    public void Entries_are_sorted_by_path_so_two_runs_produce_the_same_file()
    {
        this.Write("zebra.txt", "z");
        this.Write("Amstrad/board.txt", "a");
        this.Write("Commodore/board.txt", "c");

        IReadOnlyList<DataChecksumManifest.Entry> entries =
            DataChecksumManifest.Scan(this.DataRoot, DataChecksumManifestTests.BaseUrl);

        Assert.Equal(
            ["Amstrad/board.txt", "Commodore/board.txt", "zebra.txt"],
            entries.Select(entry => entry.File));
    }

    // ------------------------------------------------------------------ writing

    [Fact]
    public void Writing_produces_the_three_property_names_every_client_reads()
    {
        // *** A WIRE CONTRACT WITH INSTALLED COPIES OF CRT. *** These names are lowercase "file",
        // "checksum" and "url"; renaming a C# property must not change them, which is why they
        // carry JsonPropertyName. A rename here makes older clients read the manifest as empty.
        this.Write("Commodore/C64/250407/Data.xlsx", "x");

        Assert.Equal(1, DataChecksumManifest.Write(
            this.DataRoot, DataChecksumManifestTests.BaseUrl, this.ManifestPath));

        using JsonDocument document = JsonDocument.Parse(File.ReadAllText(this.ManifestPath));
        JsonElement row = document.RootElement[0];

        Assert.Equal("Commodore/C64/250407/Data.xlsx", row.GetProperty("file").GetString());
        Assert.Equal(DataChecksumManifestTests.Sha256Of("x"), row.GetProperty("checksum").GetString());
        Assert.Contains("Data.xlsx", row.GetProperty("url").GetString()!, StringComparison.Ordinal);
    }

    [Fact]
    public void The_manifest_is_written_BESIDE_the_data_root_not_inside_it()
    {
        // The project owner's own words: one folder up from "Data". It is configured separately
        // precisely because it cannot be derived from the data root.
        this.Write("a.txt", "x");

        DataChecksumManifest.Write(
            this.DataRoot, DataChecksumManifestTests.BaseUrl, this.ManifestPath);

        Assert.True(File.Exists(this.ManifestPath));
        Assert.False(File.Exists(Path.Combine(this.DataRoot, "dataChecksums.json")));
    }

    // ###########################################################################################
    // *** REFUSING TO WRITE AN EMPTY MANIFEST IS THE IMPORTANT GUARD HERE. ***
    //
    // An empty scan means the data root was wrong, unreadable or momentarily unmounted - never
    // that the tree is genuinely empty. Writing it would tell every client that all data had been
    // deleted. Keeping the previous manifest is always the better failure.
    // ###########################################################################################
    [Fact]
    public void An_EMPTY_scan_leaves_the_previous_manifest_untouched()
    {
        File.WriteAllText(this.ManifestPath, "[PREVIOUS]");

        Assert.Equal(-1, DataChecksumManifest.Write(
            Path.Combine(this.thisWorkspace.Root, "no-such-data-root"),
            DataChecksumManifestTests.BaseUrl,
            this.ManifestPath));

        Assert.Equal("[PREVIOUS]", File.ReadAllText(this.ManifestPath));
    }

    [Fact]
    public void No_manifest_path_writes_nothing_rather_than_throwing()
    {
        this.Write("a.txt", "x");

        // Runs after a publish that has already succeeded, so a misconfiguration must be reported
        // by the caller rather than thrown at a maintainer.
        Assert.Equal(-1, DataChecksumManifest.Write(
            this.DataRoot, DataChecksumManifestTests.BaseUrl, string.Empty));
    }

    [Fact]
    public void No_temp_file_is_left_behind_after_a_successful_write()
    {
        this.Write("a.txt", "x");

        DataChecksumManifest.Write(
            this.DataRoot, DataChecksumManifestTests.BaseUrl, this.ManifestPath);

        Assert.Empty(Directory.GetFiles(this.thisWorkspace.Root, "*.tmp_*"));
    }

    [Fact]
    public void An_existing_manifest_is_REPLACED_rather_than_appended_to()
    {
        File.WriteAllText(this.ManifestPath, "[\"stale\"]");
        this.Write("a.txt", "x");

        DataChecksumManifest.Write(
            this.DataRoot, DataChecksumManifestTests.BaseUrl, this.ManifestPath);

        string json = File.ReadAllText(this.ManifestPath);

        Assert.DoesNotContain("stale", json, StringComparison.Ordinal);
        Assert.Single(JsonDocument.Parse(json).RootElement.EnumerateArray());
    }
}
