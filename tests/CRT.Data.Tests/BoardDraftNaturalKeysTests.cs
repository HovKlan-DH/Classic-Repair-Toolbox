using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// Characterisation tests for BoardDraftNaturalKeys - the row-identity rule BoardDraftApplier
// merges by, and (per NewContributeStrategy.md's "Retiring the UUIDs") what a future server-side
// diff will pair rows on too. The one thing worth pinning here is that keys are built from
// TRIMMED, case-preserved values with a separator that cannot appear in ordinary board data, so
// two rows that a human would call "the same row" always produce the same key even with stray
// whitespace, while genuinely different values never collide by accident.
public sealed class BoardDraftNaturalKeysTests
{
    [Fact]
    public void ForComponent_trims_surrounding_whitespace()
    {
        Assert.Equal(BoardDraftNaturalKeys.ForComponent("U8"), BoardDraftNaturalKeys.ForComponent("  U8  "));
    }

    [Fact]
    public void ForComponent_treats_null_as_empty()
    {
        Assert.Equal(BoardDraftNaturalKeys.ForComponent(string.Empty), BoardDraftNaturalKeys.ForComponent(null!));
    }

    [Fact]
    public void ForComponentImage_is_sensitive_to_every_one_of_its_four_parts()
    {
        string baseline = BoardDraftNaturalKeys.ForComponentImage("U8", "CPU", "1", "Clock");

        Assert.NotEqual(baseline, BoardDraftNaturalKeys.ForComponentImage("U9", "CPU", "1", "Clock"));
        Assert.NotEqual(baseline, BoardDraftNaturalKeys.ForComponentImage("U8", "PLA", "1", "Clock"));
        Assert.NotEqual(baseline, BoardDraftNaturalKeys.ForComponentImage("U8", "CPU", "2", "Clock"));
        Assert.NotEqual(baseline, BoardDraftNaturalKeys.ForComponentImage("U8", "CPU", "1", "Reset"));
    }

    [Fact]
    public void ForComponentImage_does_not_let_field_boundaries_shift_and_collide()
    {
        // Without a separator that cannot appear in the fields themselves, ("U8", "CPU1") and
        // ("U8C", "PU1") could concatenate to the same string. The separator is what stops that.
        string a = BoardDraftNaturalKeys.ForComponentImage("U8", "CPU1", string.Empty, string.Empty);
        string b = BoardDraftNaturalKeys.ForComponentImage("U8C", "PU1", string.Empty, string.Empty);

        Assert.NotEqual(a, b);
    }

    [Fact]
    public void ForComponentHighlight_distinguishes_the_same_label_on_different_schematics()
    {
        string mainBoard = BoardDraftNaturalKeys.ForComponentHighlight("Main board", "J1");
        string daughterBoard = BoardDraftNaturalKeys.ForComponentHighlight("Daughterboard", "J1");

        Assert.NotEqual(mainBoard, daughterBoard);
    }

    [Fact]
    public void ForCredit_distinguishes_two_different_people_in_the_same_category_and_subcategory()
    {
        string alex = BoardDraftNaturalKeys.ForCredit("Schematics", "Tracing", "Alex");
        string sam = BoardDraftNaturalKeys.ForCredit("Schematics", "Tracing", "Sam");

        Assert.NotEqual(alex, sam);
    }

    // ###########################################################################################
    // A credit's main contribution is part of its identity (2026-11-15) - the same person credited twice for the same work
    // is two rows - and a credit without one keys exactly as it always did. The key also NAMES the row (Describe), so the
    // contribution stands in it as one line of its words, its formatting gone and a long one cut short.
    // ###########################################################################################
    [Fact]
    public void ForCredit_holds_the_main_contribution_as_one_short_line()
    {
        Assert.Equal(BoardDraftNaturalKeys.ForCredit("Board data", "", "Bo"), BoardDraftNaturalKeys.ForCredit("Board data", "", "Bo", "  "));
        Assert.NotEqual(BoardDraftNaturalKeys.ForCredit("Board data", "", "Bo", "Fixed U8"), BoardDraftNaturalKeys.ForCredit("Board data", "", "Bo", "Added the pinouts"));

        Assert.Equal(
            "Board data / Bo / Added the pinouts of every chip • U1 to U8",
            BoardDraftNaturalKeys.Describe(BoardDraftNaturalKeys.ForCredit("Board data", "", "Bo", "Added the **pinouts** of every chip\n- U1 to U8")));

        string described = BoardDraftNaturalKeys.Describe(BoardDraftNaturalKeys.ForCredit("Board data", "", "Bo", new string('x', 200)));
        Assert.Matches(@"x\.\.\. #[0-9a-f]{6}$", described);
        Assert.Equal("Board data / Bo / ".Length + BoardDraftNaturalKeys.ContributionInKey, described.Length);
    }

    // ###########################################################################################
    // *** Two long contributions that begin alike are two keys (code review, 2026-11-16). *** Cut short alone, the same person
    // credited twice for the same kind of work - "... for these components on PAL: U1, U2" and "...: U7, U8" - was one key,
    // and every comparison left the second out. In other capitals they are still the one key, as keys are compared.
    // ###########################################################################################
    [Fact]
    public void ForCredit_tells_apart_long_contributions_that_begin_alike()
    {
        const string Start = "Measured the oscilloscope baselines for these components on PAL: ";

        string first = BoardDraftNaturalKeys.ForCredit("Board data", "", "Unknown", Start + "U1, U2");
        string second = BoardDraftNaturalKeys.ForCredit("Board data", "", "Unknown", Start + "U7, U8");

        Assert.NotEqual(first, second, StringComparer.OrdinalIgnoreCase);
        Assert.Equal(first, BoardDraftNaturalKeys.ForCredit("Board data", "", "Unknown", (Start + "U1, U2").ToUpperInvariant()), StringComparer.OrdinalIgnoreCase);
    }

    // A long contribution is never cut between the two halves of one character (code review, 2026-11-16): an emoji standing
    // where it is cut goes whole, rather than leaving half of it in the name of the row.
    [Fact]
    public void ForCredit_cuts_a_long_contribution_between_characters_never_inside_one()
    {
        int cut = BoardDraftNaturalKeys.ContributionInKey - 3 - " #123456".Length;
        string contribution = new string('x', cut - 1) + "\U0001F527" + new string('y', 50);

        string key = BoardDraftNaturalKeys.ForCredit("Board data", "", "Bo", contribution);

        Assert.DoesNotContain(key, char.IsHighSurrogate);
        Assert.Matches(new string('x', cut - 1) + @"\.\.\. #[0-9a-f]{6}$", key);
    }

    // ------------------------------------------------------------------ Component region (2026-09-24)

    [Fact]
    public void A_component_WITHOUT_a_region_keys_exactly_as_it_always_did()
    {
        // Nearly every component has no region; their keys must not move.
        Assert.Equal("U8", BoardDraftNaturalKeys.ForComponent("U8"));
        Assert.Equal("U8", BoardDraftNaturalKeys.ForComponent(" U8 ", "  "));
        Assert.Equal("U8", BoardDraftNaturalKeys.ForRow(new ComponentEntry { BoardLabel = "U8" }));
    }

    [Fact]
    public void The_same_label_in_two_REGIONS_is_two_rows()
    {
        // A regionalised component is one row per region. Keyed on the label alone they were one
        // row to every comparison, and an added NTSC variant was invisible.
        string pal = BoardDraftNaturalKeys.ForRow(new ComponentEntry { BoardLabel = "U8", Region = "PAL" });
        string ntsc = BoardDraftNaturalKeys.ForRow(new ComponentEntry { BoardLabel = "U8", Region = "NTSC" });

        Assert.NotEqual(pal, ntsc);
        Assert.NotEqual("U8", pal);
        Assert.Equal("U8", pal.Split(BoardDraftNaturalKeys.Separator)[0]);
    }

    // ###########################################################################################
    // ColumnsOf names the columns each sheet's key is built from - for the duplicate warning's words
    // (2026-10-03: "the same Board label, Region, Pin and Name"). It must name EXACTLY what ForRow
    // reads, or the warning tells a contributor two rows share something that is not what makes
    // them collide. So every column of every sheet is changed alone: a named one must move the
    // key, any other must not.
    // ###########################################################################################
    [Fact]
    public void ColumnsOf_names_exactly_the_columns_each_sheets_key_is_built_from()
    {
        foreach (BoardWorkbookSchema.SheetDefinition sheet in BoardWorkbookSchema.AllSheets)
        {
            Dictionary<string, string> row = sheet.ColumnOrder.ToDictionary(column => column, column => $"value of {column}", StringComparer.OrdinalIgnoreCase);
            string key = KeyOf(sheet, row);
            IReadOnlyList<string> named = BoardDraftNaturalKeys.ColumnsOf(sheet.SheetName);

            foreach (string column in sheet.ColumnOrder)
            {
                var changed = new Dictionary<string, string>(row, StringComparer.OrdinalIgnoreCase) { [column] = "something else" };
                bool moved = KeyOf(sheet, changed) != key;

                Assert.True(moved == named.Contains(column), $"[{sheet.SheetName}] / [{column}]: {(moved ? "moves the key but is not named" : "is named but does not move the key")}");
            }
        }
    }

    // ###########################################################################################
    // PropertiesOf is ColumnsOf's twin on the entry side, and decides which cells a row may change
    // and still be the same row (BoardDataDiffer.PairRenamedRows, 2026-10-04). Named one too many,
    // two rows differing in that cell pair as a "rename"; one too few, a key change counts as
    // "something else changed" and never pairs. So every property of every row type is changed
    // alone: a named one must move the key, any other must not.
    // ###########################################################################################
    [Theory]
    [InlineData(typeof(BoardSchematicEntry))]
    [InlineData(typeof(ComponentEntry))]
    [InlineData(typeof(ComponentImageEntry))]
    [InlineData(typeof(ComponentHighlightEntry))]
    [InlineData(typeof(ComponentLocalFileEntry))]
    [InlineData(typeof(ComponentLinkEntry))]
    [InlineData(typeof(BoardLocalFileEntry))]
    [InlineData(typeof(BoardLinkEntry))]
    [InlineData(typeof(CreditEntry))]
    [InlineData(typeof(KiCadImportantSignalEntry))]
    [InlineData(typeof(KiCadCalibrationEntry))]
    public void PropertiesOf_names_exactly_the_properties_each_rows_key_is_built_from(System.Type rowType)
    {
        IReadOnlyList<string> named = BoardDraftNaturalKeys.PropertiesOf(rowType);
        Assert.NotEmpty(named);

        string key = BoardDraftNaturalKeys.ForRow(EveryPropertySet(rowType, changed: null));

        foreach (System.Reflection.PropertyInfo property in rowType.GetProperties())
        {
            bool moved = BoardDraftNaturalKeys.ForRow(EveryPropertySet(rowType, changed: property.Name)) != key;

            Assert.True(moved == named.Contains(property.Name), $"[{rowType.Name}] / [{property.Name}]: {(moved ? "moves the key but is not named" : "is named but does not move the key")}");
        }
    }

    // A row type with no key rule has no key properties - ForRow throws for it, so nothing pairs it.
    [Fact]
    public void A_type_with_no_key_rule_has_no_key_properties()
    {
        Assert.Empty(BoardDraftNaturalKeys.PropertiesOf(typeof(string)));
    }

    // A row with every writable property set to a value of its own, and `changed` set to another.
    private static object EveryPropertySet(System.Type rowType, string? changed)
    {
        object row = System.Activator.CreateInstance(rowType)!;

        foreach (System.Reflection.PropertyInfo property in rowType.GetProperties().Where(property => property.CanWrite))
        {
            bool other = property.Name == changed;

            object value = property.PropertyType == typeof(string) ? $"{(other ? "other" : "value")} of {property.Name}"
                : property.PropertyType == typeof(double) ? (other ? 2.5 : 1.5)
                : property.PropertyType == typeof(bool) ? other
                : throw new System.NotSupportedException(property.PropertyType.Name);

            property.SetValue(row, value);
        }

        return row;
    }

    // A key as words, never with the separator, which renders as a box.
    [Fact]
    public void A_key_is_described_with_its_parts_and_without_the_separator()
    {
        Assert.Equal("R307 / Pinout (secondary)", BoardDraftNaturalKeys.Describe(BoardDraftNaturalKeys.ForComponentImage("R307", "", "", "Pinout (secondary)")));
    }

    private static string KeyOf(BoardWorkbookSchema.SheetDefinition sheet, IReadOnlyDictionary<string, string> row) =>
        BoardDraftNaturalKeys.ForRow(BoardWorkbookSchema.MapRows(sheet, [row]).Single());
}
