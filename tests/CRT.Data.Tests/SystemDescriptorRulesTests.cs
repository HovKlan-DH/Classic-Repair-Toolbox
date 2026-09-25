using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Threading;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// The derivations behind system.json (NewContributeStrategy.md Phase 4, task 7).
//
// THE ID IS "Manufacturer/Hardware/Board" - the key the `systems` table, the data tree, the sync
// manifest and every draft folder already use. A random surrogate id was built here first and
// removed: it existed to survive a rename, and nothing in CRT can rename a system (a folder path
// IS a system's identity everywhere, and SubmissionManifest.Renames covers rows INSIDE a board,
// never the board itself). See SystemDescriptorRules.BuildSystemId's own header.
//
// ContentHash is the other half, and it is the kind of code that fails silently: a hash that
// disagrees between two machines makes every client believe it is permanently out of date, with
// nothing visibly broken to point at.
// ###########################################################################################
public sealed class SystemDescriptorRulesTests
{
    private static SystemContentEntry File_(string path, string hash) => new(path, hash);

    // ------------------------------------------------------------------ BuildSystemId

    [Fact]
    public void A_system_id_is_the_three_name_parts_joined_by_slashes()
    {
        Assert.Equal(
            "Commodore/C64/250407",
            SystemDescriptorRules.BuildSystemId("Commodore", "C64", "250407"));
    }

    // It is the key a DATABASE and a FOLDER LOOKUP use, so it carries no file name - unlike
    // ExcelDataFile, whose trailing ".xlsx" names a file that, for a draft-only system, does not
    // even exist.
    [Fact]
    public void A_system_id_carries_no_file_name()
    {
        string id = SystemDescriptorRules.BuildSystemId("Commodore", "C64", "250407");

        Assert.DoesNotContain(".xlsx", id);
        Assert.Equal(3, id.Split('/').Length);
    }

    [Fact]
    public void Whitespace_around_a_name_part_is_normalised_out_of_the_id()
    {
        Assert.Equal(
            "Commodore/C64/250407",
            SystemDescriptorRules.BuildSystemId("  Commodore ", "C64  ", " 250407"));
    }

    // A half-formed id would be a database key naming a system that cannot exist, so a blank part
    // yields no id at all rather than "Commodore//250407".
    [Theory]
    [InlineData("", "C64", "250407")]
    [InlineData("Commodore", "", "250407")]
    [InlineData("Commodore", "C64", "")]
    [InlineData("   ", "C64", "250407")]
    [InlineData(null, "C64", "250407")]
    public void A_blank_name_part_yields_no_id_at_all(string? manufacturer, string? hardware, string? board)
    {
        Assert.Equal(string.Empty, SystemDescriptorRules.BuildSystemId(manufacturer, hardware, board));
    }

    // ------------------------------------------------------------------ SystemIdFromExcelDataFile

    // The app carries a system as its ExcelDataFile everywhere, so this is the conversion every
    // client-side caller needs - and having it here is what stops each call site taking the string
    // apart by hand, which is where the two forms would drift.
    [Fact]
    public void An_excel_data_file_converts_to_the_system_id_by_dropping_its_file_name()
    {
        Assert.Equal(
            "Commodore/C64/250407",
            SystemDescriptorRules.SystemIdFromExcelDataFile("Commodore/C64/250407/Data C64 250407 v2.0.0.xlsx"));
    }

    [Theory]
    [InlineData("")]
    [InlineData(null)]
    [InlineData("nofolders.xlsx")]
    [InlineData("Commodore/Data.xlsx")]
    [InlineData("Commodore/C64/Data.xlsx")]
    public void An_excel_data_file_of_the_wrong_shape_yields_no_id(string? excelDataFile)
    {
        Assert.Equal(string.Empty, SystemDescriptorRules.SystemIdFromExcelDataFile(excelDataFile));
    }

    // The two ways to arrive at an id must agree, or a system submitted from a draft would carry a
    // different key from the same system published by the maintainer.
    [Fact]
    public void Both_ways_of_building_an_id_agree()
    {
        string excelDataFile = NewSystemIdentity.BuildExcelDataFile("Commodore", "C64", "250407");

        Assert.Equal(
            SystemDescriptorRules.BuildSystemId("Commodore", "C64", "250407"),
            SystemDescriptorRules.SystemIdFromExcelDataFile(excelDataFile));
    }

    // ------------------------------------------------------------------ IsValidSystemId

    [Fact]
    public void An_ordinary_id_is_valid()
    {
        Assert.True(SystemDescriptorRules.IsValidSystemId("Commodore/C64/250407"));
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("   ")]
    [InlineData("Commodore")]                       // one segment
    [InlineData("Commodore/C64")]                   // two
    [InlineData("Commodore/C64/250407/Data.xlsx")]  // four - the ExcelDataFile, not the id
    public void An_id_without_exactly_three_segments_is_refused(string? candidate)
    {
        Assert.False(SystemDescriptorRules.IsValidSystemId(candidate));
    }

    // ###########################################################################################
    // AN EMPTY SEGMENT IS CAUGHT rather than collapsing into a valid id.
    //
    // Splitting with RemoveEmptyEntries would turn "Commodore//250407" into two segments and then
    // refuse it for the wrong reason - and "Commodore/C64/250407/" into three and ACCEPT it, which
    // is a different key for the same system.
    // ###########################################################################################
    [Theory]
    [InlineData("Commodore//250407")]
    [InlineData("/C64/250407")]
    [InlineData("Commodore/C64/")]
    [InlineData("Commodore/C64/250407/")]
    public void An_id_with_an_empty_segment_is_refused(string candidate)
    {
        Assert.False(SystemDescriptorRules.IsValidSystemId(candidate));
    }

    // ###########################################################################################
    // THE SEGMENT RULE IS NewSystemIdentity'S, DELIBERATELY.
    //
    // That rule already refuses traversal, reserved device names, trailing dots, control
    // characters and every character ANY platform reserves. A second opinion here about what a
    // folder name may be would eventually disagree with it - at which point a system could be
    // created locally and refused by the server, or worse, accepted by the server and impossible
    // to sync on Windows.
    //
    // This is the test that fails if someone re-implements the check locally.
    // ###########################################################################################
    [Theory]
    [InlineData("Commodore/C64/../etc")]
    [InlineData("Commodore/C64/250407.")]
    [InlineData("Commodore/CON/250407")]
    [InlineData("Commodore/C64/250*407")]
    [InlineData("Commodore/C6:4/250407")]
    public void An_id_whose_segment_could_not_be_a_folder_is_refused(string candidate)
    {
        Assert.False(SystemDescriptorRules.IsValidSystemId(candidate));
    }

    // The database column is VARCHAR(255). Refusing here rather than letting MariaDB truncate,
    // which would map two different systems onto one key.
    [Fact]
    public void An_over_length_id_is_refused()
    {
        string segment = new('x', NewSystemIdentity.MaxSegmentLength);
        string atLimit = $"{segment}/{segment}/{segment}";

        Assert.Equal(SystemDescriptorRules.MaximumSystemIdLength, atLimit.Length);
        Assert.True(SystemDescriptorRules.IsValidSystemId(atLimit));

        Assert.False(SystemDescriptorRules.IsValidSystemId(atLimit + "x"));
    }

    // The limit is derived from the segment rule rather than guessed, so raising one raises the
    // other - and it stays inside the column.
    [Fact]
    public void The_id_limit_stays_within_the_database_column()
    {
        Assert.True(SystemDescriptorRules.MaximumSystemIdLength <= 255);
    }

    // Anything this builds must be accepted by the check, or a system could be created and then
    // refused on submission.
    [Fact]
    public void A_built_id_is_always_a_valid_id()
    {
        Assert.True(SystemDescriptorRules.IsValidSystemId(
            SystemDescriptorRules.BuildSystemId("Commodore", "C64", "250407")));

        Assert.True(SystemDescriptorRules.IsValidSystemId(
            SystemDescriptorRules.BuildSystemId("Sinclair", "ZX Spectrum", "Issue 6A")));
    }

    // ------------------------------------------------------------------ ComputeContentHash

    [Fact]
    public void The_same_content_hashes_the_same_way_twice()
    {
        var files = new[] { File_("a.png", "aaa"), File_("b.png", "bbb") };

        Assert.Equal(
            SystemDescriptorRules.ComputeContentHash("2026-09-21", files),
            SystemDescriptorRules.ComputeContentHash("2026-09-21", files));
    }

    // ###########################################################################################
    // ORDER OF THE INPUT MUST NOT MATTER - a directory walk returns files in whatever order the
    // filesystem feels like, and two machines hashing the same published tree have to agree.
    // ###########################################################################################
    [Fact]
    public void The_hash_does_not_depend_on_the_order_the_files_arrive_in()
    {
        var forwards = new[] { File_("a.png", "aaa"), File_("b.png", "bbb"), File_("c.png", "ccc") };
        var backwards = forwards.Reverse().ToArray();

        Assert.Equal(
            SystemDescriptorRules.ComputeContentHash("r1", forwards),
            SystemDescriptorRules.ComputeContentHash("r1", backwards));
    }

    // ###########################################################################################
    // THE SORT IS ORDINAL, NOT CULTURE-AWARE.
    //
    // A culture-aware sort reorders these differently under a Turkish or Swedish collation, so
    // the same published tree would hash differently depending on the SERVER'S LOCALE - a failure
    // that appears only on a machine nobody thought to test, and that looks like every client
    // being permanently out of date.
    //
    // Driven by actually switching culture, because a comment asserting "ordinal" proves nothing.
    // ###########################################################################################
    [Fact]
    public void The_hash_is_the_same_under_a_culture_that_sorts_differently()
    {
        var files = new[]
        {
            File_("zebra.png", "111"),
            File_("Illustration.png", "222"),
            File_("angstrom.png", "333"),
        };

        string invariantHash = SystemDescriptorRules.ComputeContentHash("r1", files);

        CultureInfo original = Thread.CurrentThread.CurrentCulture;

        try
        {
            foreach (string culture in new[] { "tr-TR", "sv-SE" })
            {
                Thread.CurrentThread.CurrentCulture = new CultureInfo(culture);

                Assert.Equal(invariantHash, SystemDescriptorRules.ComputeContentHash("r1", files));
            }
        }
        finally
        {
            Thread.CurrentThread.CurrentCulture = original;
        }
    }

    [Fact]
    public void Changing_one_files_bytes_changes_the_hash()
    {
        Assert.NotEqual(
            SystemDescriptorRules.ComputeContentHash("r1", new[] { File_("a.png", "aaa"), File_("b.png", "bbb") }),
            SystemDescriptorRules.ComputeContentHash("r1", new[] { File_("a.png", "aaa"), File_("b.png", "CHANGED") }));
    }

    [Fact]
    public void Adding_a_file_changes_the_hash()
    {
        Assert.NotEqual(
            SystemDescriptorRules.ComputeContentHash("r1", new[] { File_("a.png", "aaa") }),
            SystemDescriptorRules.ComputeContentHash("r1", new[] { File_("a.png", "aaa"), File_("b.png", "bbb") }));
    }

    // Renaming a file changes the hash even though the bytes are identical - a rename IS a change
    // to the published tree, and a client holding the old name does not hold the current state.
    [Fact]
    public void Renaming_a_file_changes_the_hash_even_though_its_bytes_did_not()
    {
        Assert.NotEqual(
            SystemDescriptorRules.ComputeContentHash("r1", new[] { File_("a.png", "aaa") }),
            SystemDescriptorRules.ComputeContentHash("r1", new[] { File_("b.png", "aaa") }));
    }

    // A metadata-only republish (same files, corrected revision date) must still move the hash,
    // or clients would never pick the correction up.
    [Fact]
    public void Changing_only_the_revision_changes_the_hash()
    {
        var files = new[] { File_("a.png", "aaa") };

        Assert.NotEqual(
            SystemDescriptorRules.ComputeContentHash("2026-09-01", files),
            SystemDescriptorRules.ComputeContentHash("2026-09-21", files));
    }

    // ###########################################################################################
    // THE DELIMITER PREVENTS A CONSTRUCTIBLE COLLISION.
    //
    // A naive implementation concatenates path + hash + path + hash, under which ("ab","cd") and
    // ("a","bcd") produce the same byte stream. That is not a theoretical hazard: the file names
    // come from a contribution, so an attacker chooses them. These two file sets are genuinely
    // different published trees and must hash differently.
    // ###########################################################################################
    [Fact]
    public void Two_different_file_sets_that_would_concatenate_identically_hash_differently()
    {
        Assert.NotEqual(
            SystemDescriptorRules.ComputeContentHash("r", new[] { File_("ab", "cd") }),
            SystemDescriptorRules.ComputeContentHash("r", new[] { File_("a", "bcd") }));
    }

    // A newline in a path would make the delimited stream ambiguous again. It is refused rather
    // than escaped, because SubmissionPathRules already refuses control characters - so such a
    // path cannot legitimately reach a published tree, and quietly accepting one here would mean
    // the two rules disagree about what is publishable.
    [Fact]
    public void A_path_containing_a_newline_is_refused_rather_than_hashed_ambiguously()
    {
        Assert.Throws<ArgumentException>(() =>
            SystemDescriptorRules.ComputeContentHash("r", new[] { File_("a\nb.png", "aaa") }));
    }

    // Case-sensitive, like every other path comparison in this project - the Linux server is the
    // filesystem that matters, and these are two files there.
    [Fact]
    public void Paths_differing_only_in_case_are_two_different_files()
    {
        Assert.NotEqual(
            SystemDescriptorRules.ComputeContentHash("r", new[] { File_("U8.png", "aaa") }),
            SystemDescriptorRules.ComputeContentHash("r", new[] { File_("u8.png", "aaa") }));
    }

    [Fact]
    public void A_system_with_no_files_still_hashes()
    {
        string hash = SystemDescriptorRules.ComputeContentHash("r1", []);

        Assert.False(string.IsNullOrWhiteSpace(hash));
        Assert.Equal(64, hash.Length);
    }

    // ------------------------------------------------------------------ Build

    [Fact]
    public void A_built_descriptor_carries_every_field_it_was_given()
    {
        var published = new DateTimeOffset(2026, 9, 21, 10, 0, 0, TimeSpan.Zero);

        SystemDescriptor descriptor = SystemDescriptorRules.Build(
            "Commodore", "C64", "250407", "2026-09-21", published,
            new[] { "Dennis" }, SystemDescriptorRules.SystemOrigin.Contributed,
            new[] { File_("a.png", "aaa") });

        Assert.Equal("Commodore/C64/250407", descriptor.SystemId);
        Assert.Equal("Commodore", descriptor.Manufacturer);
        Assert.Equal("C64", descriptor.Hardware);
        Assert.Equal("250407", descriptor.Board);
        Assert.Equal("2026-09-21", descriptor.Revision);
        Assert.Equal(published, descriptor.PublishedUtc);
        Assert.Equal(new[] { "Dennis" }, descriptor.Maintainers);
        Assert.Equal("contributed", descriptor.Origin);

        Assert.Equal(
            SystemDescriptorRules.ComputeContentHash("2026-09-21", new[] { File_("a.png", "aaa") }),
            descriptor.ContentHash);
    }

    // ###########################################################################################
    // The id is DERIVED from the three name parts rather than passed in, so a descriptor's id and
    // the parts it names can never disagree - which a hand-assembled one could, and which would
    // show as a system filed under a key that does not match what is printed on it.
    // ###########################################################################################
    [Fact]
    public void A_built_descriptors_id_always_matches_the_parts_it_names()
    {
        SystemDescriptor descriptor = SystemDescriptorRules.Build(
            "Sinclair", "ZX Spectrum", "Issue 6A", "r1", DateTimeOffset.UtcNow,
            null, SystemDescriptorRules.SystemOrigin.Shipped, []);

        Assert.Equal(
            SystemDescriptorRules.BuildSystemId(
                descriptor.Manufacturer, descriptor.Hardware, descriptor.Board),
            descriptor.SystemId);
    }

    // Republishing the same system keeps the same id, because the id is what the system IS rather
    // than something allocated to it. This is the property the surrogate id was supposed to give
    // and that the path form gives for free, as long as the names do not change.
    [Fact]
    public void Republishing_the_same_system_produces_the_same_id()
    {
        SystemDescriptor first = SystemDescriptorRules.Build(
            "Commodore", "C64", "250407", "r1", DateTimeOffset.UtcNow,
            null, SystemDescriptorRules.SystemOrigin.Shipped, []);

        SystemDescriptor second = SystemDescriptorRules.Build(
            "Commodore", "C64", "250407", "r2", DateTimeOffset.UtcNow,
            new[] { "Dennis" }, SystemDescriptorRules.SystemOrigin.Shipped,
            new[] { File_("a.png", "aaa") });

        Assert.Equal(first.SystemId, second.SystemId);
    }

    // Maintainers are de-duplicated and ordered, so republishing an unchanged system does not
    // produce a file differing only in the order a database query happened to return rows.
    [Fact]
    public void Maintainers_are_deduplicated_and_ordered()
    {
        SystemDescriptor descriptor = SystemDescriptorRules.Build(
            "Commodore", "C64", "250407", "r", DateTimeOffset.UtcNow,
            new[] { "Zoe", "Dennis", "  Dennis  ", "", "   ", "Anders" },
            SystemDescriptorRules.SystemOrigin.Contributed, []);

        Assert.Equal(new[] { "Anders", "Dennis", "Zoe" }, descriptor.Maintainers);
    }

    [Fact]
    public void A_system_with_no_maintainers_gets_an_empty_list_rather_than_null()
    {
        SystemDescriptor descriptor = SystemDescriptorRules.Build(
            "Commodore", "C64", "250407", "r", DateTimeOffset.UtcNow,
            null, SystemDescriptorRules.SystemOrigin.Shipped, []);

        Assert.NotNull(descriptor.Maintainers);
        Assert.Empty(descriptor.Maintainers);
    }

    // ------------------------------------------------------------------ SystemOrigin

    // The two values the `systems` table's own origin column was created to hold.
    [Fact]
    public void The_two_origins_are_the_ones_the_schema_names()
    {
        Assert.Equal("shipped", SystemDescriptorRules.SystemOrigin.Shipped);
        Assert.Equal("contributed", SystemDescriptorRules.SystemOrigin.Contributed);

        Assert.True(SystemDescriptorRules.SystemOrigin.IsKnown("shipped"));
        Assert.True(SystemDescriptorRules.SystemOrigin.IsKnown("contributed"));
    }

    // An unrecognised origin is not an error - a newer server may know more of them - but a caller
    // choosing a label needs to be able to tell it is looking at something it cannot name.
    [Theory]
    [InlineData("mirrored")]
    [InlineData("")]
    [InlineData(null)]
    public void An_origin_this_build_does_not_know_is_reported_as_unknown(string? origin)
    {
        Assert.False(SystemDescriptorRules.SystemOrigin.IsKnown(origin));
    }
}
