using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// Tests for NewSystemIdentity - the identity a system created through "Add a new system" gets
// (NewContributeStrategy.md Phase 2, session 2c, task 9).
//
// Two things make this worth covering thoroughly despite being plain string work. First, the
// ExcelDataFile it produces is the system's identity EVERYWHERE (the BoardDataReader cache key,
// DraftManager.GetSystemFolder's input, the schematic image root), and three separate consumers
// read meaning out of its segment COUNT - a shape change here breaks them silently. Second, each
// validation rule exists because its failure mode produces a system that looks created but is
// broken in a way the contributor cannot diagnose; the trailing-dot case is the sharpest, since
// Windows strips it silently and the folder then no longer matches the key stored in draft.json.
public sealed class NewSystemIdentityTests
{
    // --------------------------------------------------------------- BuildExcelDataFile

    [Fact]
    public void The_identity_is_three_folder_segments_plus_a_data_file_name()
    {
        string result = NewSystemIdentity.BuildExcelDataFile("Commodore", "C64", "250407");

        Assert.Equal("Commodore/C64/250407/Data C64 250407.xlsx", result);
    }

    // Always "/", never the platform separator - ExcelDataFile follows the sync manifest's own
    // convention, and a "\" would not match the split in DraftManager.GetSystemFolder on Windows.
    [Fact]
    public void The_identity_always_uses_forward_slashes()
    {
        string result = NewSystemIdentity.BuildExcelDataFile("Commodore", "C64", "250407");

        Assert.DoesNotContain('\\', result);
    }

    // DraftManager.GetSystemFolder needs at least two segments and
    // HardwareBoardEntry.ShortHardwareBoardLabel needs three, so the built key must always carry
    // four parts (three folders plus the file name). This is the contract those two depend on.
    [Fact]
    public void The_identity_carries_four_slash_separated_parts_so_both_consumers_can_split_it()
    {
        string result = NewSystemIdentity.BuildExcelDataFile("Commodore", "C64", "250407");

        Assert.Equal(4, result.Split('/').Length);
    }

    // No version suffix, unlike a real board's "Data C64 250407 v2.0.0.xlsx" - keeping a version in
    // step with the resolved main workbook by hand is one of the burdens task 9 removes.
    [Fact]
    public void The_identity_carries_no_version_suffix()
    {
        string result = NewSystemIdentity.BuildExcelDataFile("Commodore", "C64", "250407");

        Assert.DoesNotContain(" v", result);
    }

    [Theory]
    [InlineData("", "C64", "250407")]
    [InlineData("Commodore", "", "250407")]
    [InlineData("Commodore", "C64", "")]
    [InlineData("  ", "C64", "250407")]
    public void A_blank_segment_yields_no_identity_at_all_rather_than_a_half_formed_one(
        string manufacturer, string hardware, string board)
    {
        Assert.Equal(string.Empty, NewSystemIdentity.BuildExcelDataFile(manufacturer, hardware, board));
    }

    [Fact]
    public void Surrounding_whitespace_is_trimmed_out_of_the_identity()
    {
        string result = NewSystemIdentity.BuildExcelDataFile("  Commodore ", " C64", "250407  ");

        Assert.Equal("Commodore/C64/250407/Data C64 250407.xlsx", result);
    }

    // ------------------------------------------------------------------ SanitizePathSegment

    // Invisible to the user, but "C 64" and "C  64" would otherwise be two different systems with
    // two different folders.
    [Fact]
    public void Inner_whitespace_runs_collapse_to_a_single_space()
    {
        Assert.Equal("Commodore 64", NewSystemIdentity.SanitizePathSegment("Commodore   64"));
    }

    [Fact]
    public void Sanitizing_a_blank_value_gives_an_empty_string()
    {
        Assert.Equal(string.Empty, NewSystemIdentity.SanitizePathSegment("   "));
        Assert.Equal(string.Empty, NewSystemIdentity.SanitizePathSegment(null));
    }

    // Deliberately NOT a character-stripping pass: an illegal character is reported and corrected
    // by the user, because silently turning "C64:NTSC" into "C64NTSC" would create a folder they
    // did not ask for and cannot find again.
    [Fact]
    public void Sanitizing_does_not_strip_illegal_characters_it_only_normalizes_whitespace()
    {
        Assert.Equal("C64:NTSC", NewSystemIdentity.SanitizePathSegment("C64:NTSC"));
    }

    // ------------------------------------------------------------------- IsValidPathSegment

    [Theory]
    [InlineData("Commodore")]
    [InlineData("C64")]
    [InlineData("250407")]
    [InlineData("250407 rev B")]
    [InlineData("A500+")]
    public void An_ordinary_name_is_valid(string value)
    {
        Assert.True(NewSystemIdentity.IsValidPathSegment(value, out string reason), reason);
        Assert.Equal(string.Empty, reason);
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    [InlineData(null)]
    public void A_blank_name_is_rejected(string? value)
    {
        Assert.False(NewSystemIdentity.IsValidPathSegment(value, out string reason));
        Assert.Contains("blank", reason);
    }

    // "/" and "\" would break the "/"-separated key this segment ends up inside even on a platform
    // whose filesystem would have accepted them, so both are rejected everywhere.
    [Theory]
    [InlineData("C64/NTSC")]
    [InlineData("C64\\NTSC")]
    [InlineData("C64:NTSC")]
    [InlineData("C64*")]
    [InlineData("C64?")]
    [InlineData("C64\"x\"")]
    [InlineData("C64<x>")]
    [InlineData("C64|x")]
    public void A_name_with_a_path_breaking_character_is_rejected_and_names_the_character(string value)
    {
        Assert.False(NewSystemIdentity.IsValidPathSegment(value, out string reason));
        Assert.Contains("cannot contain", reason);
    }

    // ###########################################################################################
    // THE RULE IS THE STRICTEST PLATFORM'S, NOT THE HOST'S.
    //
    // Path.GetInvalidFileNameChars returns the full Windows set on Windows and only "/" and NUL on
    // Linux. An implementation that trusted it would accept "Commo*dore" when run on Linux - and
    // that name becomes a FOLDER in a tree every platform syncs, so the system would be impossible
    // to sync for every Windows user who downloaded it, with the failure appearing on a machine
    // that never saw the name being typed.
    //
    // This test asserts the whole Windows-reserved set is refused REGARDLESS of where the suite
    // runs, which is what makes it meaningful on the Linux CI runner. It fails there against an
    // implementation that only unions in the host's own set.
    //
    // WorkbookExportModel.SanitizeForFileName documents the same reasoning for exported file
    // names; this is that lesson applied to a directory segment.
    // ###########################################################################################
    [Theory]
    [InlineData('<')]
    [InlineData('>')]
    [InlineData(':')]
    [InlineData('"')]
    [InlineData('/')]
    [InlineData('\\')]
    [InlineData('|')]
    [InlineData('?')]
    [InlineData('*')]
    public void Every_character_ANY_platform_reserves_is_refused_on_EVERY_platform(char reserved)
    {
        Assert.False(
            NewSystemIdentity.IsValidPathSegment($"C64{reserved}NTSC", out string reason),
            $"[{reserved}] must be refused on every platform, not only where the host rejects it.");

        Assert.Contains("cannot contain", reason);
    }

    [Fact]
    public void A_pasted_newline_or_tab_becomes_a_SPACE_rather_than_a_refusal()
    {
        // WHITESPACE CONTROL CHARACTERS ARE NORMALISED, NOT REJECTED, and that is deliberate:
        // SanitizePathSegment splits on all whitespace and rejoins with single spaces, so a name
        // pasted out of a spreadsheet cell arrives as "C64 NTSC" rather than being refused for
        // something the user cannot see. Only NON-whitespace control characters reach the
        // invalid-character check.
        //
        // Pinned because it reads like a hole - a control character surviving validation - and the
        // next person to notice would be right to check. The answer is that it does not survive;
        // it is gone before the check runs.
        Assert.True(NewSystemIdentity.IsValidPathSegment("C64\nNTSC", out _));
        Assert.Equal("C64 NTSC", NewSystemIdentity.SanitizePathSegment("C64\nNTSC"));
        Assert.Equal("C64 NTSC", NewSystemIdentity.SanitizePathSegment("C64\tNTSC"));
    }

    [Fact]
    public void A_non_whitespace_control_character_is_refused()
    {
        // These are not normalised away, and would break a log line or a manifest entry even where
        // the filesystem tolerated them.
        Assert.False(NewSystemIdentity.IsValidPathSegment("C64\u0001NTSC", out string reason));
        Assert.Contains("cannot contain", reason);
    }

    [Theory]
    [InlineData(".")]
    [InlineData("..")]
    public void A_relative_path_name_is_rejected(string value)
    {
        Assert.False(NewSystemIdentity.IsValidPathSegment(value, out string reason));
        Assert.Contains("[.]", reason);
    }

    // Windows refuses these outright as folder names, with or without an extension - a registered
    // system whose folder cannot be created would point at nothing.
    [Theory]
    [InlineData("CON")]
    [InlineData("con")]
    [InlineData("PRN")]
    [InlineData("NUL")]
    [InlineData("COM1")]
    [InlineData("LPT9")]
    [InlineData("CON.txt")]
    public void A_reserved_device_name_is_rejected(string value)
    {
        Assert.False(NewSystemIdentity.IsValidPathSegment(value, out string reason));
        Assert.Contains("reserved", reason);
    }

    // The sharpest of these rules: Windows SILENTLY STRIPS a trailing dot, so the folder created on
    // disk would no longer match the ExcelDataFile key recorded in draft.json, and the draft would
    // become unfindable by the very lookup that is supposed to locate it.
    [Fact]
    public void A_name_ending_in_a_dot_is_rejected()
    {
        Assert.False(NewSystemIdentity.IsValidPathSegment("C64.", out string reason));
        Assert.Contains("dot", reason);
    }

    // A trailing SPACE is stripped by the sanitize pass before the check, so it is simply accepted
    // as the trimmed name rather than rejected - the user gets "C64", which is what they meant.
    [Fact]
    public void A_name_with_a_trailing_space_is_accepted_as_its_trimmed_form()
    {
        Assert.True(NewSystemIdentity.IsValidPathSegment("C64 ", out _));
        Assert.Equal("C64", NewSystemIdentity.SanitizePathSegment("C64 "));
    }

    [Fact]
    public void An_over_length_name_is_rejected()
    {
        string tooLong = new('x', NewSystemIdentity.MaxSegmentLength + 1);

        Assert.False(NewSystemIdentity.IsValidPathSegment(tooLong, out string reason));
        Assert.Contains("too long", reason);
    }

    [Fact]
    public void A_name_at_exactly_the_length_limit_is_accepted()
    {
        string atLimit = new('x', NewSystemIdentity.MaxSegmentLength);

        Assert.True(NewSystemIdentity.IsValidPathSegment(atLimit, out _));
    }

    // -------------------------------------------------------------------- ExtractManufacturer

    [Fact]
    public void The_manufacturer_is_the_first_folder_segment_of_an_existing_identity()
    {
        Assert.Equal(
            "Commodore",
            NewSystemIdentity.ExtractManufacturer("Commodore/C64/250407/Data C64 250407 v2.0.0.xlsx"));
    }

    [Theory]
    [InlineData("")]
    [InlineData(null)]
    [InlineData("nofolders.xlsx")]
    public void A_key_with_no_folder_segments_yields_no_manufacturer(string? value)
    {
        Assert.Equal(string.Empty, NewSystemIdentity.ExtractManufacturer(value));
    }

    // -------------------------------------------------------------- ExtractHardware / ExtractBoard

    [Fact]
    public void The_hardware_and_board_are_the_second_and_third_folder_segments()
    {
        const string Key = "Commodore/C64/250407/Data C64 250407 v2.0.0.xlsx";

        Assert.Equal("C64", NewSystemIdentity.ExtractHardware(Key));
        Assert.Equal("250407", NewSystemIdentity.ExtractBoard(Key));
    }

    // ###########################################################################################
    // *** THE POINT OF THESE TWO: the three parts must REBUILD the system id. ***
    //
    // SubmissionValidator rebuilds the id from Manufacturer/Hardware/Board and refuses the
    // submission when it does not match the SystemId the client sent. TabDrafts used to take the
    // id from the path but the hardware and board from the master workbook's DISPLAY names
    // ("Commodore 64", "250407 (long board)"), so every submission for a published board was
    // refused with 400 "identity.system_id_mismatch" (reported 2026-09-23).
    //
    // This test is the one that fails if either side moves on its own.
    // ###########################################################################################
    [Fact]
    public void The_extracted_parts_rebuild_the_system_id_the_validator_expects()
    {
        const string Key = "Commodore/C64/250407/Data C64 250407 v2.0.0.xlsx";

        string rebuilt = SystemDescriptorRules.BuildSystemId(
            NewSystemIdentity.ExtractManufacturer(Key),
            NewSystemIdentity.ExtractHardware(Key),
            NewSystemIdentity.ExtractBoard(Key));

        Assert.Equal(SystemDescriptorRules.SystemIdFromExcelDataFile(Key), rebuilt);
    }

    // A display name is exactly what must NOT be used, so this pins the failure it caused: the
    // master workbook calls this board "Commodore 64" / "250407 (long board)".
    [Fact]
    public void Display_names_do_NOT_rebuild_the_system_id()
    {
        const string Key = "Commodore/C64/250407/Data C64 250407 v2.0.0.xlsx";

        string fromDisplayNames = SystemDescriptorRules.BuildSystemId(
            "Commodore",
            "Commodore 64",
            "250407 (long board)");

        Assert.NotEqual(SystemDescriptorRules.SystemIdFromExcelDataFile(Key), fromDisplayNames);
    }

    [Theory]
    [InlineData("")]
    [InlineData(null)]
    [InlineData("nofolders.xlsx")]
    [InlineData("Commodore/Data.xlsx")]
    public void A_key_without_enough_folder_segments_yields_no_hardware_or_board(string? value)
    {
        Assert.Equal(string.Empty, NewSystemIdentity.ExtractHardware(value));
        Assert.Equal(string.Empty, NewSystemIdentity.ExtractBoard(value));
    }
}
