using System.Reflection;
using CRT;
using Handlers.DataHandling;
using Handlers.OnlineHandling;

namespace ClassicRepairToolbox.Tests;

// Characterisation tests for the manifest-validation predicates inside OnlineServices - the
// checks every entry of the online checksum manifest must pass before the sync writes a single
// byte to disk.
//
// This is a trust boundary, not a parser. The manifest is fetched over the network, and its
// entries decide WHERE a file is written and WHERE it is downloaded from: a hole in
// TryResolveValidatedLocalPath is a write outside the data root, and a hole in
// TryCreateTrustedDownloadUri is a download from somebody else's server. That puts these in the
// same category as ExternalTargetLauncher, and they are tested here for the same reason.
//
// IMPORTANT - why this file uses reflection:
// CLAUDE.md rules OnlineServices out of scope as an I/O boundary class, and that is right for
// FetchManifestAsync / SyncFilesAsync / DownloadFileAsync - those need a real server. The four
// predicates covered here are not I/O; they are string, Uri and Path logic that happens to live
// in that class, and they run BEFORE any file is touched. They are private, so - exactly as in
// ExternalTargetLauncherTests - they are reached by reflection rather than by making a real sync
// testable. If these lookups ever fail, the fix is to make the predicates `internal` (the app
// project already grants InternalsVisibleTo to this project), not to drive a sync.
//
// Nothing here touches the filesystem: the predicates resolve path STRINGS only, so the data root
// below is a synthetic absolute path that is never created. That is part of what is being pinned
// down - containment must not depend on the file already existing.
//
// TryCreateTrustedDownloadUri reads AppConfig.GetChecksumsUrl(), which reads
// UserSettings.DownloadDataFromTestSource, so this class joins the "UserSettings" collection per
// the CLAUDE.md rule about test classes that touch that static.
[Collection("UserSettings")]
public class OnlineServicesTests
{
    // Never created on disk - see the header. GetFullPath normalises it for the current platform
    // so the same assertions hold on Windows and on the Linux CI runner.
    private static readonly string DataRoot =
        Path.GetFullPath(Path.Combine(Path.GetTempPath(), "crt-online-tests-root"));

    // 64 hex characters - the shape SHA-256 hex is required to have.
    private const string ValidChecksum =
        "0123456789abcdef0123456789abcdef0123456789abcdef0123456789abcdef";

    // Read from configuration rather than hardcoded, so that changing the domain does not quietly
    // turn every "trusted host" assertion below into a test of a string that no longer matters.
    private static string TrustedAuthority =>
        new Uri(AppConfig.GetChecksumsUrl(), UriKind.Absolute).Authority;

    // ------------------------------------------------------------- private predicates

    private static MethodInfo GetPrivateMethod(string name)
    {
        MethodInfo? method = typeof(OnlineServices)
            .GetMethod(name, BindingFlags.NonPublic | BindingFlags.Static);

        Assert.True(method is not null,
            $"OnlineServices.{name} not found - if it was renamed, update these tests.");

        return method!;
    }

    private static bool TryResolveValidatedLocalPath(
        string? dataRoot, string? manifestFile, out string localPath, out string failureReason)
    {
        object?[] args = { dataRoot, manifestFile, null, null };
        bool result = (bool)GetPrivateMethod("TryResolveValidatedLocalPath").Invoke(null, args)!;
        localPath = (string)(args[2] ?? string.Empty);
        failureReason = (string)(args[3] ?? string.Empty);
        return result;
    }

    private static bool TryNormalizeManifestChecksum(string? checksum, out string normalizedChecksum)
    {
        object?[] args = { checksum, null };
        bool result = (bool)GetPrivateMethod("TryNormalizeManifestChecksum").Invoke(null, args)!;
        normalizedChecksum = (string)(args[1] ?? string.Empty);
        return result;
    }

    private static bool TryCreateTrustedDownloadUri(
        string? url, out Uri? downloadUri, out string failureReason)
    {
        object?[] args = { url, null, null };
        bool result = (bool)GetPrivateMethod("TryCreateTrustedDownloadUri").Invoke(null, args)!;
        downloadUri = args[1] as Uri;
        failureReason = (string)(args[2] ?? string.Empty);
        return result;
    }

    private static bool TryValidateManifestEntry(
        DataFileEntry entry,
        string dataRoot,
        out string localPath,
        out Uri? downloadUri,
        out string expectedChecksum,
        out string failureReason)
    {
        object?[] args = { entry, dataRoot, null, null, null, null };
        bool result = (bool)GetPrivateMethod("TryValidateManifestEntry").Invoke(null, args)!;
        localPath = (string)(args[2] ?? string.Empty);
        downloadUri = args[3] as Uri;
        expectedChecksum = (string)(args[4] ?? string.Empty);
        failureReason = (string)(args[5] ?? string.Empty);
        return result;
    }

    // --------------------------------------------- local path: what is allowed inside

    [Fact]
    public void A_relative_manifest_path_resolves_inside_the_data_root()
    {
        // The manifest is generated server-side and always uses forward slashes, so the separator
        // conversion is not cosmetic - without it nothing would resolve on Windows.
        Assert.True(TryResolveValidatedLocalPath(
            DataRoot, "Commodore/C64/250407/250407.xlsx", out string localPath, out string failureReason));

        Assert.Equal(Path.Combine(DataRoot, "Commodore", "C64", "250407", "250407.xlsx"), localPath);
        Assert.Empty(failureReason);
    }

    [Fact]
    public void Traversal_that_stays_inside_the_data_root_is_allowed()
    {
        // ".." is not banned outright - only escaping is. A generated path that doubles back is
        // odd rather than hostile, and it still lands somewhere legitimate.
        Assert.True(TryResolveValidatedLocalPath(
            DataRoot, "Commodore/../Amstrad/CPC464/board.xlsx", out string localPath, out _));

        Assert.Equal(Path.Combine(DataRoot, "Amstrad", "CPC464", "board.xlsx"), localPath);
    }

    [Fact]
    public void Surrounding_whitespace_in_a_manifest_path_is_trimmed()
    {
        Assert.True(TryResolveValidatedLocalPath(
            DataRoot, "  Commodore/C64/notes.txt  ", out string localPath, out _));

        Assert.Equal(Path.Combine(DataRoot, "Commodore", "C64", "notes.txt"), localPath);
    }

    [Fact]
    public void Validation_does_not_require_the_file_to_exist()
    {
        // Phase 1 of SyncFilesAsync validates every entry before it looks at the disk at all, and
        // the common case is a file that is not there yet. A containment check that only worked
        // for existing files would pass in testing and wave new downloads through in the field.
        Assert.False(Directory.Exists(DataRoot));

        Assert.True(TryResolveValidatedLocalPath(
            DataRoot, "Commodore/C64/never-downloaded.xlsx", out _, out _));
    }

    // ---------------------------------------------- local path: what must be rejected

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("   ")]
    public void An_empty_data_root_is_rejected(string? dataRoot)
    {
        Assert.False(TryResolveValidatedLocalPath(dataRoot, "file.xlsx", out _, out string failureReason));
        Assert.Equal("data root is empty", failureReason);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("   ")]
    public void An_empty_manifest_path_is_rejected(string? manifestFile)
    {
        // What a manifest entry missing its "file" key deserialises to.
        Assert.False(TryResolveValidatedLocalPath(DataRoot, manifestFile, out _, out string failureReason));
        Assert.Equal("manifest file path is empty", failureReason);
    }

    [Theory]
    [InlineData("/etc/passwd")]
    [InlineData("/")]
    [InlineData("  /var/tmp/evil.so  ")]
    public void An_absolute_manifest_path_is_rejected(string manifestFile)
    {
        // The manifest may only ever name a path RELATIVE to the data root. An absolute one would
        // make Path.Combine discard the root entirely and write wherever the server said.
        Assert.False(TryResolveValidatedLocalPath(DataRoot, manifestFile, out _, out string failureReason));
        Assert.Equal("manifest file path must be relative", failureReason);
    }

    [Fact]
    public void Drive_qualified_and_unc_manifest_paths_are_rejected_on_windows()
    {
        // Path.IsPathRooted is platform-dependent, so this rule is asserted where it applies
        // rather than pretended to be cross-platform: on Linux "C:\Windows\evil.dll" contains no
        // separator at all and is simply a very odd FILE NAME sitting inside the data root.
        if (!OperatingSystem.IsWindows())
            return;

        Assert.False(TryResolveValidatedLocalPath(
            DataRoot, @"C:\Windows\System32\evil.dll", out _, out string driveFailure));
        Assert.Equal("manifest file path must be relative", driveFailure);

        Assert.False(TryResolveValidatedLocalPath(
            DataRoot, @"\\attacker\share\evil.dll", out _, out string uncFailure));
        Assert.Equal("manifest file path must be relative", uncFailure);
    }

    [Theory]
    [InlineData("../outside.xlsx")]
    [InlineData("Commodore/../../outside.xlsx")]
    [InlineData("Commodore/C64/../../../../../../etc/passwd")]
    public void A_manifest_path_that_escapes_the_data_root_is_rejected(string manifestFile)
    {
        Assert.False(TryResolveValidatedLocalPath(DataRoot, manifestFile, out _, out string failureReason));
        Assert.Equal("manifest file path escapes data root", failureReason);
    }

    [Fact]
    public void A_sibling_directory_whose_name_merely_starts_with_the_data_root_is_rejected()
    {
        // This is why the comparison is against the data root WITH a trailing separator appended.
        // "<root>-evil/x" starts with "<root>" as a plain string, so a naive StartsWith check
        // would happily accept a write into the directory next door.
        string escape = "../" + Path.GetFileName(DataRoot) + "-evil/x.xlsx";

        Assert.False(TryResolveValidatedLocalPath(DataRoot, escape, out _, out string failureReason));
        Assert.Equal("manifest file path escapes data root", failureReason);
    }

    [Theory]
    [InlineData(".")]
    [InlineData("Commodore/..")]
    public void A_manifest_path_that_resolves_to_the_data_root_itself_is_rejected(string manifestFile)
    {
        // Contained, but not a file - it would have the sync try to overwrite the data root
        // directory with a download, which is why it gets its own check and its own message.
        Assert.False(TryResolveValidatedLocalPath(DataRoot, manifestFile, out _, out string failureReason));
        Assert.Equal("manifest file path resolves to the data root", failureReason);
    }

    [Fact]
    public void A_manifest_path_that_cannot_be_resolved_is_reported_rather_than_thrown()
    {
        // An embedded null makes Path.GetFullPath throw. The manifest is remote input, so one
        // malformed entry has to be skipped and logged - never take the whole sync down with it.
        Assert.False(TryResolveValidatedLocalPath(
            DataRoot, "Commodore/ba\0d.xlsx", out _, out string failureReason));

        Assert.StartsWith("manifest file path is invalid:", failureReason);
    }

    // ------------------------------------------------------------------------ checksum

    [Fact]
    public void A_well_formed_checksum_is_accepted_trimmed_and_lowercased()
    {
        // Downloaded bytes are hashed with Convert.ToHexString(...).ToLowerInvariant(), and the
        // two are then compared with StringComparison.Ordinal - so normalising here is what stops
        // an uppercase manifest checksum failing every single verification.
        Assert.True(TryNormalizeManifestChecksum(
            "  " + ValidChecksum.ToUpperInvariant() + "  ", out string normalized));

        Assert.Equal(ValidChecksum, normalized);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("   ")]
    public void A_missing_checksum_is_rejected(string? checksum)
    {
        Assert.False(TryNormalizeManifestChecksum(checksum, out _));
    }

    [Theory]
    // 63 characters - one short.
    [InlineData("0123456789abcdef0123456789abcdef0123456789abcdef0123456789abcde")]
    // 65 characters - one long.
    [InlineData("0123456789abcdef0123456789abcdef0123456789abcdef0123456789abcdef0")]
    // Right length, but 'g' is not hex.
    [InlineData("g123456789abcdef0123456789abcdef0123456789abcdef0123456789abcdef")]
    // Right length, but a separator has crept in.
    [InlineData("0123456789abcdef-123456789abcdef0123456789abcdef0123456789abcdef")]
    public void A_checksum_of_the_wrong_length_or_alphabet_is_rejected(string checksum)
    {
        Assert.False(TryNormalizeManifestChecksum(checksum, out _));
    }

    [Fact]
    public void A_rejected_checksum_leaves_the_normalised_value_empty()
    {
        // The caller passes this straight on to the download verification, so it must not be left
        // holding a half-normalised value from a failed parse.
        Assert.False(TryNormalizeManifestChecksum("not-a-checksum", out string normalized));
        Assert.Empty(normalized);
    }

    // --------------------------------------------------------------------- download URL

    [Fact]
    public void A_download_url_on_the_configured_authority_is_accepted()
    {
        Assert.True(TryCreateTrustedDownloadUri(
            $"https://{TrustedAuthority}/app-data/Commodore/C64/250407.xlsx",
            out Uri? downloadUri,
            out string failureReason));

        Assert.NotNull(downloadUri);
        Assert.Equal(TrustedAuthority, downloadUri!.Authority);
        Assert.Empty(failureReason);
    }

    [Fact]
    public void The_scheme_and_host_are_matched_case_insensitively_after_trimming()
    {
        Assert.True(TryCreateTrustedDownloadUri(
            $"  HTTPS://{TrustedAuthority.ToUpperInvariant()}/app-data/x.xlsx  ", out Uri? downloadUri, out _));

        Assert.NotNull(downloadUri);
    }

    [Fact]
    public void An_explicit_default_https_port_is_accepted_because_Uri_elides_it()
    {
        // Characterisation, not a rule anybody wrote: Uri.Authority drops :443 for https, so an
        // explicitly-defaulted port still compares equal to the trusted authority.
        Assert.True(TryCreateTrustedDownloadUri(
            $"https://{TrustedAuthority}:443/app-data/x.xlsx", out _, out _));
    }

    [Fact]
    public void A_non_default_port_on_the_trusted_host_is_rejected()
    {
        // The other half of the above: :8443 survives into Authority and therefore does not match.
        Assert.False(TryCreateTrustedDownloadUri(
            $"https://{TrustedAuthority}:8443/app-data/x.xlsx", out Uri? downloadUri, out string failureReason));

        Assert.Null(downloadUri);
        Assert.Contains("is not trusted", failureReason);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("   ")]
    public void An_empty_download_url_is_rejected(string? url)
    {
        Assert.False(TryCreateTrustedDownloadUri(url, out _, out string failureReason));
        Assert.Equal("download URL is empty", failureReason);
    }

    [Theory]
    [InlineData("app-data/Commodore/C64/250407.xlsx")]
    [InlineData("not a url at all")]
    public void A_download_url_that_is_not_absolute_is_rejected(string url)
    {
        Assert.False(TryCreateTrustedDownloadUri(url, out _, out string failureReason));
        Assert.Equal("download URL is not a valid absolute URI", failureReason);
    }

    [Theory]
    // Protocol-relative - a browser would resolve it, but it carries no authority to check.
    [InlineData("//classic-repair-toolbox.dk/app-data/x.xlsx")]
    // A bare local path, which is what a manifest generator bug would most plausibly emit.
    [InlineData("/etc/passwd")]
    public void A_slash_prefixed_download_url_is_rejected_whichever_way_Uri_reads_it(string url)
    {
        // Deliberately asserts the OUTCOME and not the failure reason: Uri's implicit-file-path
        // handling is platform-dependent, so "//host/share" parses on Windows as an absolute UNC
        // file:// URI (rejected a check later, for not being HTTPS) while elsewhere it may not
        // parse as absolute at all (rejected immediately). Either route is correct and the message
        // differs, so pinning one of them here would just be a test that fails on the Linux CI
        // runner or on the maintainer's machine depending on who wrote it. What must hold on every
        // platform is that nothing slash-prefixed ever becomes a download URI.
        Assert.False(TryCreateTrustedDownloadUri(url, out Uri? downloadUri, out _));
        Assert.Null(downloadUri);
    }

    [Theory]
    [InlineData("http")]
    [InlineData("ftp")]
    [InlineData("file")]
    public void A_download_url_that_is_not_https_is_rejected(string scheme)
    {
        // Checked before the authority, so plain http to the RIGHT host is still refused: the
        // checksum verification that follows cannot tell a tampered download from a real one.
        Assert.False(TryCreateTrustedDownloadUri(
            $"{scheme}://{TrustedAuthority}/app-data/x.xlsx", out _, out string failureReason));

        Assert.Equal("download URL must use HTTPS", failureReason);
    }

    [Fact]
    public void A_download_url_on_another_host_is_rejected_and_the_reason_names_it()
    {
        Assert.False(TryCreateTrustedDownloadUri(
            "https://evil.example/app-data/x.xlsx", out Uri? downloadUri, out string failureReason));

        Assert.Null(downloadUri);
        Assert.Contains("evil.example", failureReason);
    }

    [Fact]
    public void A_host_that_only_looks_like_the_trusted_one_is_rejected()
    {
        // The three shapes a sloppier check would let through: the trusted name as a subdomain of
        // somebody else's domain, an extra label bolted onto the front of it, and the trusted name
        // pushed into the userinfo so that the URL merely READS like the right host. Uri.Authority
        // excludes userinfo, which is what makes the third one fail.
        Assert.False(TryCreateTrustedDownloadUri($"https://{TrustedAuthority}.evil.example/x.xlsx", out _, out _));
        Assert.False(TryCreateTrustedDownloadUri($"https://cdn.{TrustedAuthority}/x.xlsx", out _, out _));
        Assert.False(TryCreateTrustedDownloadUri($"https://{TrustedAuthority}@evil.example/x.xlsx", out _, out _));
    }

    [Fact]
    public void Switching_to_the_beta_data_source_cannot_widen_what_is_trusted()
    {
        // GetChecksumsUrl() returns the BETA manifest when UserSettings.DownloadDataFromTestSource
        // is on, and TryCreateTrustedDownloadUri derives the trusted authority from whichever URL
        // it returns. Both are on the same host today, so that setting changes the PATH the
        // manifest is read from and nothing about which server downloads may come from. If that
        // ever stops being true the beta switch becomes a trust decision, and this is the test
        // that will say so.
        Assert.Equal(
            new Uri(AppConfig.ChecksumsUrl, UriKind.Absolute).Authority,
            new Uri(AppConfig.ChecksumsUrl_test, UriKind.Absolute).Authority);
    }

    // The trusted authority is MEMOISED against the URL it came from, not computed once for the
    // life of the process. It used to be a static readonly filled in by the type initializer, which
    // is wrong now that GetChecksumsUrl() reads a user setting: whichever source happened to be
    // selected the first time this class was touched would be frozen in for the whole session.
    //
    // Both sources are on the same host today, so this cannot assert two DIFFERENT authorities - it
    // asserts the thing that would actually break, which is that toggling the setting either way
    // still validates against the authority the currently-configured manifest names. A frozen value
    // passes this only while the two hosts agree; the test above is what fails when they stop.
    [Fact]
    public void The_trusted_authority_is_re_derived_when_the_data_source_setting_changes()
    {
        bool originalSetting = UserSettings.DownloadDataFromTestSource;

        try
        {
            foreach (bool useTestSource in new[] { false, true, false })
            {
                UserSettings.DownloadDataFromTestSource = useTestSource;

                string expectedAuthority =
                    new Uri(AppConfig.GetChecksumsUrl(), UriKind.Absolute).Authority;

                Assert.True(TryCreateTrustedDownloadUri(
                    $"https://{expectedAuthority}/app-data/x.xlsx", out Uri? uri, out string reason));

                Assert.Equal(string.Empty, reason);
                Assert.Equal(expectedAuthority, uri!.Authority);

                // The memo must not have widened what is accepted: an unrelated host is still
                // refused after the switch, which is the failure mode a stale cache would cause if
                // the two sources ever moved apart.
                Assert.False(TryCreateTrustedDownloadUri(
                    "https://evil.example/app-data/x.xlsx", out _, out _));
            }
        }
        finally
        {
            UserSettings.DownloadDataFromTestSource = originalSetting;
        }
    }

    // Repeated validation of the same source is the hot path - a sync validates one manifest entry
    // per file, thousands of them - and it must keep returning the same answer while the memo is
    // being reused, not just on the first call that fills it.
    [Fact]
    public void Repeated_validation_against_the_same_source_stays_consistent()
    {
        string authority = TrustedAuthority;

        for (int i = 0; i < 50; i++)
        {
            Assert.True(TryCreateTrustedDownloadUri($"https://{authority}/app-data/{i}.xlsx", out Uri? uri, out _));
            Assert.Equal(authority, uri!.Authority);
            Assert.False(TryCreateTrustedDownloadUri($"https://other.example/app-data/{i}.xlsx", out _, out _));
        }
    }

    // ------------------------------------------------------- the whole entry, in order

    [Fact]
    public void A_valid_entry_yields_a_local_path_a_download_uri_and_a_lowercased_checksum()
    {
        DataFileEntry entry = new()
        {
            File = "Commodore/C64/250407/250407.xlsx",
            Checksum = ValidChecksum.ToUpperInvariant(),
            Url = $"https://{TrustedAuthority}/app-data/Commodore/C64/250407/250407.xlsx",
        };

        Assert.True(TryValidateManifestEntry(
            entry, DataRoot, out string localPath, out Uri? downloadUri, out string expectedChecksum, out string failureReason));

        Assert.Equal(Path.Combine(DataRoot, "Commodore", "C64", "250407", "250407.xlsx"), localPath);
        Assert.NotNull(downloadUri);
        Assert.Equal(ValidChecksum, expectedChecksum);
        Assert.Empty(failureReason);
    }

    [Fact]
    public void The_path_check_runs_before_the_checksum_and_url_checks()
    {
        // Order decides what gets logged. An entry that is wrong in three ways should be reported
        // as the containment failure, because that is the one worth somebody noticing.
        DataFileEntry entry = new()
        {
            File = "../outside.xlsx",
            Checksum = "nonsense",
            Url = "http://evil.example/x.xlsx",
        };

        Assert.False(TryValidateManifestEntry(entry, DataRoot, out _, out _, out _, out string failureReason));
        Assert.Equal("manifest file path escapes data root", failureReason);
    }

    [Fact]
    public void The_checksum_check_runs_before_the_url_check()
    {
        DataFileEntry entry = new()
        {
            File = "Commodore/C64/250407/250407.xlsx",
            Checksum = "nonsense",
            Url = "http://evil.example/x.xlsx",
        };

        Assert.False(TryValidateManifestEntry(entry, DataRoot, out _, out _, out _, out string failureReason));
        Assert.Equal("manifest checksum is missing or invalid", failureReason);
    }

    [Fact]
    public void An_entry_whose_only_fault_is_an_untrusted_url_is_still_rejected()
    {
        // A correct path and a correct checksum do not buy an entry a download from anywhere else.
        DataFileEntry entry = new()
        {
            File = "Commodore/C64/250407/250407.xlsx",
            Checksum = ValidChecksum,
            Url = "https://evil.example/app-data/250407.xlsx",
        };

        Assert.False(TryValidateManifestEntry(
            entry, DataRoot, out _, out Uri? downloadUri, out _, out string failureReason));

        Assert.Null(downloadUri);
        Assert.Contains("evil.example", failureReason);
    }

    [Fact]
    public void An_entry_with_every_field_defaulted_is_rejected()
    {
        // DataFileEntry's properties all default to string.Empty, which is what an entry missing
        // its keys deserialises to. It must not sail through as "the data root".
        Assert.False(TryValidateManifestEntry(
            new DataFileEntry(), DataRoot, out _, out _, out _, out string failureReason));

        Assert.Equal("manifest file path is empty", failureReason);
    }
}
