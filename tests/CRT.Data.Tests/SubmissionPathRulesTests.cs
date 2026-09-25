using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Handlers.DataHandling;
using Xunit;

namespace CRT.Data.Tests
{
    // ###########################################################################################
    // Covers the rule that decides whether a submitted file path may be written.
    //
    // THIS IS THE MOST SECURITY-SENSITIVE LOGIC IN PHASE 4. Every path here arrives over the
    // network from anyone with an account and names a location the server will create a file at.
    // A hole means writing outside the system folder - into another system's data, into the
    // service's own directory, or over a published file.
    //
    // The traversal tests are written as a LIST OF KNOWN TRICKS rather than one representative
    // case, because that is how this class of bug actually ships: the obvious "../" is caught and
    // a variant is not. Each one below defeats at least one naive implementation.
    //
    // Paths are built with Path.Combine and Path.GetTempPath rather than written as literals, so
    // these pass on the Linux CI runner as well as on Windows.
    // ###########################################################################################
    public class SubmissionPathRulesTests
    {
        private static string SystemFolder()
        {
            return Path.Combine(Path.GetTempPath(), "crt-test", "Commodore", "C64", "250407");
        }

        private static bool Allows(string relativePath)
        {
            return SubmissionPathRules.TryResolve(
                SubmissionPathRulesTests.SystemFolder(), relativePath, out _, out _);
        }

        // -----------------------------------------------------------------------------------
        // What must be allowed. A rule that rejects legitimate data is a bug too - it silently
        // blocks contributions and nobody reports it, they just give up.
        // -----------------------------------------------------------------------------------

        [Theory]
        [InlineData("schematic.png")]
        [InlineData("Images/U8.jpg")]
        [InlineData("Images/Oscilloscope/U8-pin3.png")]
        [InlineData("Datasheets/74LS08.pdf")]
        [InlineData("Datasheets/74LS08.PDF")]
        [InlineData("file with spaces.png")]
        [InlineData("Ævre-Æsel.png")]
        [InlineData("250407.xlsx")]
        public void An_ordinary_relative_path_is_allowed(string relativePath)
        {
            Assert.True(
                SubmissionPathRulesTests.Allows(relativePath),
                $"Expected [{relativePath}] to be allowed.");
        }

        [Fact]
        public void Paths_differing_only_by_case_are_BOTH_allowed_individually()
        {
            // Case is preserved, never normalised: mixed case is real in the shipped tree (27 .PDF
            // alongside 85 .pdf), and the server's filesystem is case-sensitive. Whether the two
            // may coexist in ONE submission is a separate question - see the collision test below.
            Assert.True(SubmissionPathRulesTests.Allows("Datasheets/74LS08.pdf"));
            Assert.True(SubmissionPathRulesTests.Allows("Datasheets/74LS08.PDF"));
        }

        [Fact]
        public void The_resolved_path_preserves_the_submitted_capitalisation()
        {
            // If this normalised case, a row naming "foo.PDF" would find a file written as
            // "foo.pdf" on Windows and fail on the Linux server - the exact failure the strategy
            // document warns about, appearing only in production.
            SubmissionPathRules.TryResolve(
                SubmissionPathRulesTests.SystemFolder(), "Datasheets/74LS08.PDF", out string resolved, out _);

            Assert.EndsWith("74LS08.PDF", resolved, StringComparison.Ordinal);
        }

        // -----------------------------------------------------------------------------------
        // Traversal. Each of these defeats at least one naive implementation.
        // -----------------------------------------------------------------------------------

        [Theory]
        [InlineData("../evil.png")]
        [InlineData("../../evil.png")]
        [InlineData("Images/../../evil.png")]          // traverses back out after going in
        [InlineData("Images/./../../evil.png")]        // a "." in the middle defeats a simple split
        [InlineData("..")]
        [InlineData("../")]
        [InlineData("a/../../../../../../etc/passwd")] // more levels than the path is deep
        [InlineData("./../evil.png")]
        public void A_traversing_path_is_refused(string relativePath)
        {
            Assert.False(
                SubmissionPathRulesTests.Allows(relativePath),
                $"Expected [{relativePath}] to be REFUSED - it escapes the system folder.");
        }

        [Fact]
        public void A_traversal_that_lands_back_INSIDE_the_folder_is_still_refused()
        {
            // "Images/../schematic.png" resolves inside the folder, so a containment check alone
            // would allow it. It is refused anyway because a legitimate client never produces it,
            // and accepting paths that need normalising means the stored path and the submitted
            // path can differ - which is how two manifest rows end up pointing at one file.
            Assert.False(SubmissionPathRulesTests.Allows("Images/../schematic.png"));
        }

        // -----------------------------------------------------------------------------------
        // Absolute paths, in the several forms they take across platforms.
        // -----------------------------------------------------------------------------------

        [Theory]
        [InlineData("/etc/passwd")]
        [InlineData("/tmp/evil.png")]
        [InlineData("C:/Windows/System32/evil.dll")]
        [InlineData("C:\\Windows\\evil.dll")] // windows-path-literal: a hostile input, refused on every OS.
        [InlineData("//server/share/evil.png")]
        [InlineData("\\\\server\\share\\evil.png")]
        public void An_absolute_path_is_refused(string relativePath)
        {
            // Several of these are only "rooted" on Windows - on Linux a colon is an ordinary
            // character - so they are checked explicitly rather than left to Path.IsPathRooted.
            // Otherwise the answer would depend on which platform the server runs, and the server
            // runs Linux.
            Assert.False(
                SubmissionPathRulesTests.Allows(relativePath),
                $"Expected [{relativePath}] to be REFUSED as absolute.");
        }

        // ###########################################################################################
        // IsSafelyShaped is TryResolve's checks without a folder (code review, 2026-09-25), so the
        // file rules can skip a path the path rules already refused. It must agree with TryResolve
        // on every shape - a path one accepts and the other refuses would slip between them.
        // ###########################################################################################
        [Theory]
        [InlineData("Commodore/C64/250407/sheet1.png", true)]
        [InlineData("Generic shared files/Datasheets/7805.pdf", true)]
        [InlineData("../outside.png", false)]
        [InlineData("Commodore/../../x.png", false)]
        [InlineData("Commodore/./x.png", false)]
        [InlineData("Commodore//x.png", false)]
        [InlineData("/etc/passwd", false)]
        [InlineData("C:/Windows/x.png", false)]
        [InlineData(@"Commodore\x.png", false)]
        [InlineData("Commodore/CON.png", false)]
        [InlineData("Commodore/x.png ", false)]
        [InlineData("", false)]
        public void The_shape_check_agrees_with_resolving(string relativePath, bool safe)
        {
            Assert.Equal(safe, SubmissionPathRules.IsSafelyShaped(relativePath, out string why));
            Assert.Equal(safe, SubmissionPathRulesTests.Allows(relativePath));
            Assert.Equal(safe, why.Length == 0);
        }

        [Fact]
        public void A_backslash_is_refused_rather_than_translated()
        {
            // Backslash separates directories on Windows and is a legal FILE NAME character on
            // Linux, so "a\..\b" is a traversal on one and a single odd file name on the other.
            // Translating it would make the same submission produce different trees on different
            // servers; refusing it keeps one answer.
            Assert.False(SubmissionPathRulesTests.Allows("Images\\U8.png"));
        }

        // -----------------------------------------------------------------------------------
        // Characters and names that break something downstream.
        // -----------------------------------------------------------------------------------

        [Fact]
        public void A_path_containing_a_NUL_byte_is_refused()
        {
            // A NUL truncates the path in any C library underneath, so this can pass a managed
            // containment check and then be written somewhere else entirely.
            Assert.False(SubmissionPathRulesTests.Allows("safe.png\0../../evil.png"));
        }

        [Theory]
        [InlineData("file\nname.png")]
        [InlineData("file\rname.png")]
        [InlineData("file\tname.png")]
        public void A_path_containing_a_control_character_is_refused(string relativePath)
        {
            Assert.False(SubmissionPathRulesTests.Allows(relativePath));
        }

        [Theory]
        [InlineData("CON")]
        [InlineData("CON.txt")]
        [InlineData("nul.png")]
        [InlineData("Images/COM1.jpg")]
        [InlineData("LPT9.pdf")]
        public void A_reserved_Windows_device_name_is_refused(string relativePath)
        {
            // The server runs Linux and would create these happily. A Windows client then cannot
            // sync the file, and the failure appears on a machine that never saw the submission.
            Assert.False(
                SubmissionPathRulesTests.Allows(relativePath),
                $"Expected [{relativePath}] to be REFUSED as a reserved device name.");
        }

        [Theory]
        [InlineData("file.png ")]
        [InlineData(" file.png")]
        [InlineData("folder /file.png")]
        [InlineData("file.")]
        [InlineData("Images./file.png")]
        public void A_segment_with_edge_whitespace_or_a_trailing_dot_is_refused(string relativePath)
        {
            // Windows silently STRIPS a trailing dot or space when creating a file, so two entries
            // differing only in that would collide into one - and the row referencing the other
            // would name a file that does not exist.
            Assert.False(
                SubmissionPathRulesTests.Allows(relativePath),
                $"Expected [{relativePath}] to be REFUSED.");
        }

        [Theory]
        [InlineData("")]
        [InlineData("   ")]
        [InlineData("Images//U8.png")]
        [InlineData("/")]
        public void An_empty_or_malformed_path_is_refused(string relativePath)
        {
            Assert.False(SubmissionPathRulesTests.Allows(relativePath));
        }

        [Fact]
        public void An_over_long_path_is_refused()
        {
            string tooLong = new string('a', SubmissionPathRules.MaximumPathLength + 1) + ".png";

            Assert.False(SubmissionPathRulesTests.Allows(tooLong));
        }

        [Fact]
        public void A_refusal_always_explains_itself()
        {
            // The reason reaches the contributor, who has to fix it. A bare "invalid path"
            // produces a resubmission of the same thing.
            SubmissionPathRules.TryResolve(
                SubmissionPathRulesTests.SystemFolder(), "../evil.png", out _, out string reason);

            Assert.False(string.IsNullOrWhiteSpace(reason));
            Assert.Contains("evil.png", reason, StringComparison.Ordinal);
        }

        [Fact]
        public void A_hostile_path_is_TRUNCATED_in_the_failure_message()
        {
            // A 200-character path should not flood a log line or a dialog, but enough of it must
            // survive for the contributor to recognise which entry is at fault.
            string longPath = "../" + new string('x', 150) + ".png";

            SubmissionPathRules.TryResolve(
                SubmissionPathRulesTests.SystemFolder(), longPath, out _, out string reason);

            Assert.Contains("...", reason, StringComparison.Ordinal);
            Assert.True(reason.Length < longPath.Length + 60);
        }

        // -----------------------------------------------------------------------------------
        // Hashes.
        // -----------------------------------------------------------------------------------

        [Fact]
        public void A_well_formed_hash_is_accepted()
        {
            Assert.True(SubmissionPathRules.IsValidHash(new string('a', 64)));
            Assert.True(SubmissionPathRules.IsValidHash("0123456789abcdef0123456789abcdef0123456789abcdef0123456789abcdef"));
        }

        [Theory]
        [InlineData(null)]
        [InlineData("")]
        [InlineData("abc")]
        [InlineData("ABCDEF0123456789ABCDEF0123456789ABCDEF0123456789ABCDEF0123456789")] // uppercase
        [InlineData("0123456789abcdef0123456789abcdef0123456789abcdef0123456789abcdeg")] // 'g'
        public void A_malformed_hash_is_rejected(string? hash)
        {
            // Strict rather than forgiving because this is a STORE KEY: a hash accepted in two
            // spellings stores the same blob twice, and lets two manifest entries disagree about
            // which one they meant.
            Assert.False(SubmissionPathRules.IsValidHash(hash));
        }

        // -----------------------------------------------------------------------------------
        // Whole-manifest validation.
        // -----------------------------------------------------------------------------------

        private static SubmissionFile File(string path, long size = 1024)
        {
            return new SubmissionFile { Path = path, Sha256 = new string('a', 64), SizeBytes = size };
        }

        [Fact]
        public void A_clean_manifest_produces_no_findings()
        {
            var manifest = new SubmissionManifest
            {
                Files =
                {
                    SubmissionPathRulesTests.File("schematic.png"),
                    SubmissionPathRulesTests.File("Images/U8.jpg")
                }
            };

            Assert.Empty(SubmissionPathRules.ValidateManifestPaths(
                manifest, SubmissionPathRulesTests.SystemFolder()));
        }

        [Fact]
        public void EVERY_bad_path_is_reported_not_just_the_first()
        {
            // A client that produced one bad path probably produced several, and fixing them one
            // round trip at a time is intolerable over a slow upload.
            var manifest = new SubmissionManifest
            {
                Files =
                {
                    SubmissionPathRulesTests.File("../evil.png"),
                    SubmissionPathRulesTests.File("/etc/passwd"),
                    SubmissionPathRulesTests.File("CON.txt")
                }
            };

            IReadOnlyList<ValidationFinding> findings = SubmissionPathRules.ValidateManifestPaths(
                manifest, SubmissionPathRulesTests.SystemFolder());

            Assert.Equal(3, findings.Count);
            Assert.All(findings, finding => Assert.Equal(ValidationSeverity.Error, finding.Severity));
        }

        [Fact]
        public void A_duplicate_path_is_an_error()
        {
            var manifest = new SubmissionManifest
            {
                Files =
                {
                    SubmissionPathRulesTests.File("Images/U8.png"),
                    SubmissionPathRulesTests.File("Images/U8.png")
                }
            };

            Assert.Contains(
                SubmissionPathRules.ValidateManifestPaths(manifest, SubmissionPathRulesTests.SystemFolder()),
                finding => finding.Code == "path.duplicate");
        }

        [Fact]
        public void Two_paths_differing_only_by_CASE_are_reported_as_a_collision()
        {
            // Not a security problem - a portability one. Both files can exist on the Linux server
            // and only one can exist on a contributor's Windows machine, so the tree would be
            // un-syncable for a large part of the audience.
            var manifest = new SubmissionManifest
            {
                Files =
                {
                    SubmissionPathRulesTests.File("Images/U8.png"),
                    SubmissionPathRulesTests.File("Images/u8.PNG")
                }
            };

            ValidationFinding finding = Assert.Single(
                SubmissionPathRules.ValidateManifestPaths(manifest, SubmissionPathRulesTests.SystemFolder()),
                finding => finding.Code == "path.case_collision");

            // The message must name BOTH paths, or the contributor cannot tell which pair collided.
            Assert.Contains("U8.png", finding.Message, StringComparison.Ordinal);
        }

        [Fact]
        public void An_exact_duplicate_does_not_hide_a_LATER_case_collision()
        {
            // The exact duplicate is reported and skipped, but skipping it must not also skip
            // RECORDING the path - otherwise the third entry, which differs from the first only by
            // capitalisation, is compared against an empty set and passes silently.
            //
            // That is the portability bug this check exists to catch, reaching the published tree
            // because an unrelated mistake earlier in the same manifest masked it.
            var manifest = new SubmissionManifest
            {
                Files =
                {
                    SubmissionPathRulesTests.File("Images/U8.png"),
                    SubmissionPathRulesTests.File("Images/U8.png"),
                    SubmissionPathRulesTests.File("Images/u8.PNG")
                }
            };

            IReadOnlyList<ValidationFinding> findings =
                SubmissionPathRules.ValidateManifestPaths(manifest, SubmissionPathRulesTests.SystemFolder());

            Assert.Contains(findings, finding => finding.Code == "path.duplicate");
            Assert.Contains(findings, finding => finding.Code == "path.case_collision");
        }

        [Fact]
        public void The_case_collision_check_sees_EVERY_accepted_path_that_precedes_it()
        {
            // Regression guard for the shape of bug the loop's `continue` statements invite: a
            // finding that skips the rest of an iteration must not also skip RECORDING the path,
            // or the entry after it is compared against an incomplete set.
            //
            // Here the second file carries a bad hash - a finding that deliberately does NOT
            // continue - and the third differs from it only by capitalisation. If the hash finding
            // ever grew a `continue` (the obvious "tidy-up" to make it match its neighbours), the
            // collision would stop being reported and an un-syncable tree would publish.
            var manifest = new SubmissionManifest
            {
                Files =
                {
                    SubmissionPathRulesTests.File("schematic.png"),
                    new SubmissionFile { Path = "Images/U8.png", Sha256 = "nope", SizeBytes = 10 },
                    SubmissionPathRulesTests.File("Images/u8.PNG")
                }
            };

            IReadOnlyList<ValidationFinding> findings =
                SubmissionPathRules.ValidateManifestPaths(manifest, SubmissionPathRulesTests.SystemFolder());

            Assert.Contains(findings, finding => finding.Code == "hash.malformed");
            Assert.Contains(findings, finding => finding.Code == "path.case_collision");
        }

        [Fact]
        public void A_malformed_hash_in_the_manifest_is_an_error()
        {
            var manifest = new SubmissionManifest
            {
                Files = { new SubmissionFile { Path = "a.png", Sha256 = "nope", SizeBytes = 10 } }
            };

            Assert.Contains(
                SubmissionPathRules.ValidateManifestPaths(manifest, SubmissionPathRulesTests.SystemFolder()),
                finding => finding.Code == "hash.malformed");
        }

        [Fact]
        public void An_over_large_file_is_an_error()
        {
            var manifest = new SubmissionManifest
            {
                Files = { SubmissionPathRulesTests.File("big.png", SubmissionFormat.MaximumBlobBytes + 1) }
            };

            Assert.Contains(
                SubmissionPathRules.ValidateManifestPaths(manifest, SubmissionPathRulesTests.SystemFolder()),
                finding => finding.Code == "file.too_large");
        }

        [Fact]
        public void A_negative_file_size_is_an_error()
        {
            // Nothing legitimate produces this; it means a broken client or a hand-edited payload,
            // and an unchecked negative would reach an allocation somewhere downstream.
            var manifest = new SubmissionManifest
            {
                Files = { SubmissionPathRulesTests.File("a.png", -1) }
            };

            Assert.Contains(
                SubmissionPathRules.ValidateManifestPaths(manifest, SubmissionPathRulesTests.SystemFolder()),
                finding => finding.Code == "file.too_large");
        }

        [Fact]
        public void Every_finding_names_the_path_it_is_about()
        {
            // A contributor with 400 files needs to know WHICH one is at fault.
            var manifest = new SubmissionManifest
            {
                Files =
                {
                    SubmissionPathRulesTests.File("../evil.png"),
                    new SubmissionFile { Path = "bad-hash.png", Sha256 = "x", SizeBytes = 1 }
                }
            };

            IReadOnlyList<ValidationFinding> findings = SubmissionPathRules.ValidateManifestPaths(
                manifest, SubmissionPathRulesTests.SystemFolder());

            Assert.All(findings, finding => Assert.False(string.IsNullOrWhiteSpace(finding.Subject)));
        }
    }
}
