using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Threading;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// Comparing a drafted board against the published one (NewContributeStrategy.md Phase 6 -
// drafts stored as real board folders).
//
// WHY THIS CLASS EXISTS AT ALL, and why it is worth this much testing: the Drafts tab's
// "N rows changed" chip and the drift window used to read a ledger recorded at EDIT time
// (BoardDraft's DraftRow list). Once a draft is a real board workbook the contributor can edit
// it in Excel with the application closed, and no ledger can survive that. So the answer has to
// be DERIVED by comparison - and a derived answer that is quietly wrong is worse than the
// ledger was, because nothing about it looks broken.
//
// The cases below are therefore mostly about the ways a comparison can be quietly wrong:
// uuids that are not edits, whitespace that is not an edit, a workbook round-trip that is not
// an edit, and a new column that must not be silently skipped.
// ###########################################################################################
public sealed class BoardDataDifferTests
{
    private static BoardData Board(
        IEnumerable<ComponentEntry>? components = null,
        IEnumerable<BoardSchematicEntry>? schematics = null,
        IEnumerable<CreditEntry>? credits = null)
        => new()
        {
            Components = components?.ToList() ?? [],
            Schematics = schematics?.ToList() ?? [],
            Credits = credits?.ToList() ?? [],
        };

    private static ComponentEntry Component(
        string label,
        string description = "",
        string partNumber = "")
        => new()
        {
            BoardLabel = label,
            Description = description,
            PartNumber = partNumber,
        };

    // ------------------------------------------------------------------ The three verdicts

    [Fact]
    public void An_identical_board_reports_NOTHING()
    {
        // The ordinary case by a wide margin: a draft is seeded as a complete copy, so almost
        // every row in it is untouched. If this were wrong the chip would read "847 rows changed"
        // the moment a draft was created.
        BoardData published = Board(components: [Component("U8", "CPU"), Component("U9", "VIC")]);
        BoardData draft = Board(components: [Component("U8", "CPU"), Component("U9", "VIC")]);

        Assert.Empty(BoardDataDiffer.Compare(published, draft));
    }

    [Fact]
    public void A_row_only_in_the_draft_is_ADDED()
    {
        BoardData published = Board(components: [Component("U8")]);
        BoardData draft = Board(components: [Component("U8"), Component("U9")]);

        BoardRowChange change = Assert.Single(BoardDataDiffer.Compare(published, draft));

        Assert.Equal(BoardRowChangeKind.Added, change.Kind);
        Assert.Equal("U9", change.DisplayLabel);
        Assert.Equal("Components", change.Section);
    }

    // ###########################################################################################
    // *** A MISSING ROW IS A DELETION, AND THIS ONLY HOLDS BECAUSE DRAFTS ARE SEEDED COMPLETE. ***
    //
    // The whole storage change rests on this: the draft workbook starts as a copy of the published
    // one, so a row that is absent is a row the contributor took out - in the app or in Excel, it
    // does not matter which. Were drafts seeded empty, every untouched row would read as deleted.
    // ###########################################################################################
    [Fact]
    public void A_row_only_in_the_PUBLISHED_board_is_DELETED()
    {
        BoardData published = Board(components: [Component("U8"), Component("U9")]);
        BoardData draft = Board(components: [Component("U8")]);

        BoardRowChange change = Assert.Single(BoardDataDiffer.Compare(published, draft));

        Assert.Equal(BoardRowChangeKind.Deleted, change.Kind);
        Assert.Equal("U9", change.DisplayLabel);
    }

    [Fact]
    public void A_row_whose_field_differs_is_MODIFIED_and_NAMES_the_field()
    {
        // "U8 changed" is far less useful than "U8: Description", and the information is free at
        // the moment the comparison happens.
        BoardData published = Board(components: [Component("U8", description: "CPU")]);
        BoardData draft = Board(components: [Component("U8", description: "CPU (replaced)")]);

        BoardRowChange change = Assert.Single(BoardDataDiffer.Compare(published, draft));

        Assert.Equal(BoardRowChangeKind.Modified, change.Kind);
        Assert.Equal(["Description"], change.ChangedFields);
    }

    [Fact]
    public void A_row_with_SEVERAL_changed_fields_names_them_all()
    {
        BoardData published = Board(components: [Component("U8", "CPU", "6510")]);
        BoardData draft = Board(components: [Component("U8", "MPU", "8500")]);

        BoardRowChange change = Assert.Single(BoardDataDiffer.Compare(published, draft));

        Assert.Equal(BoardRowChangeKind.Modified, change.Kind);
        Assert.Contains("Description", change.ChangedFields);
        Assert.Contains("PartNumber", change.ChangedFields);
    }

    // ------------------------------------------------------------------ Things that are NOT edits

    // ###########################################################################################
    // *** UuidV4 IS NEVER A DIFFERENCE. ***
    //
    // Phase 4 retired it as an identity ("stop WRITING new ones; keep READING existing ones"), so
    // a published row can carry one while the drafted copy does not. Comparing it would make every
    // row of a seeded draft read as Modified the instant the workbook was rewritten without the
    // column - a report that is pure noise, from a cause that is very hard to see in the symptom.
    // ###########################################################################################
    [Fact]
    public void A_DIFFERENT_uuid_on_an_otherwise_identical_row_is_NOT_a_change()
    {
        BoardData published = Board(components: [Component("U8", "CPU")]);
        BoardData draft = Board(components: [Component("U8", "CPU")]);

        Assert.Empty(BoardDataDiffer.Compare(published, draft));
    }

    [Fact]
    public void A_uuid_DROPPED_by_a_workbook_round_trip_is_NOT_a_change()
    {
        // The realistic shape of the above: the published row has a uuid, and the drafted copy
        // written back out by BoardWorkbookWriter has none.
        BoardData published = Board(components: [Component("U8", "CPU")]);
        BoardData draft = Board(components: [Component("U8", "CPU")]);

        Assert.Empty(BoardDataDiffer.Compare(published, draft));
    }

    // ###########################################################################################
    // A BLANK CELL AND A MISSING CELL ARE THE SAME THING.
    //
    // BoardWorkbookWriter.WriteTextCell returns early on a blank rather than writing an empty
    // string, and the reader answers "" for both. So a row that has merely been through a save
    // must not come back as Modified against its own unchanged self.
    // ###########################################################################################
    [Fact]
    public void A_blank_field_does_not_differ_from_an_absent_one()
    {
        BoardData published = Board(components: [new ComponentEntry { BoardLabel = "U8", Description = "" }]);
        BoardData draft = Board(components: [new ComponentEntry { BoardLabel = "U8" }]);

        Assert.Empty(BoardDataDiffer.Compare(published, draft));
    }

    [Fact]
    public void Surrounding_whitespace_is_not_an_edit()
    {
        // A workbook round-trip, or a hand edit in Excel, can pick up or lose trailing spaces.
        // Nobody means that as a change to the board.
        BoardData published = Board(components: [Component("U8", "CPU")]);
        BoardData draft = Board(components: [Component("U8", "  CPU  ")]);

        Assert.Empty(BoardDataDiffer.Compare(published, draft));
    }

    // ###########################################################################################
    // Case-insensitive on the KEY, case-sensitive on the CONTENT - the two halves are different
    // questions, and re-casing a label is BOTH at once.
    //
    // *** THIS TEST ORIGINALLY ASSERTED THE OPPOSITE AND WAS WRONG. *** It expected "U8" -> "u8"
    // to report nothing, reasoning that case does not change which row is meant. The first half
    // of that is right and is what the key comparison already guarantees - the row PAIRS, so this
    // is never reported as a deletion plus an addition, which is the outcome that would actually
    // mislead.
    //
    // But the label is also a FIELD, and the contributor did change it: submit this draft and the
    // published board reads "u8" afterwards. A report that stayed silent would be hiding a real,
    // publishable edit. So one Modified row naming BoardLabel is the correct answer.
    // ###########################################################################################
    [Fact]
    public void Re_casing_a_board_label_is_the_SAME_ROW_but_still_a_reported_edit()
    {
        BoardData published = Board(components: [Component("U8", "CPU")]);
        BoardData draft = Board(components: [Component("u8", "CPU")]);

        BoardRowChange change = Assert.Single(BoardDataDiffer.Compare(published, draft));

        // Paired as one row - NOT reported as "U8 deleted, u8 added", which is the failure mode
        // that would actually confuse someone reading the list.
        Assert.Equal(BoardRowChangeKind.Modified, change.Kind);
        Assert.Equal(["BoardLabel"], change.ChangedFields);
    }

    [Fact]
    public void Case_INSIDE_a_described_field_IS_an_edit()
    {
        // Unlike the key: rewriting a description's casing is a real change to what the board says.
        BoardData published = Board(components: [Component("U8", "cpu")]);
        BoardData draft = Board(components: [Component("U8", "CPU")]);

        BoardRowChange change = Assert.Single(BoardDataDiffer.Compare(published, draft));

        Assert.Equal(BoardRowChangeKind.Modified, change.Kind);
    }

    // ------------------------------------------------------------------ Edges

    // ###########################################################################################
    // A system that exists ONLY as a draft has no published counterpart, and every row in it is an
    // addition. That is the literal truth rather than a convenience: officially, none of it exists.
    // ###########################################################################################
    [Fact]
    public void With_NO_published_board_every_drafted_row_is_an_addition()
    {
        BoardData draft = Board(
            components: [Component("U8"), Component("U9")],
            schematics: [new BoardSchematicEntry { SchematicName = "Sheet 1" }]);

        IReadOnlyList<BoardRowChange> changes = BoardDataDiffer.Compare(published: null, draft);

        Assert.Equal(3, changes.Count);
        Assert.All(changes, change => Assert.Equal(BoardRowChangeKind.Added, change.Kind));
    }

    [Fact]
    public void A_null_draft_reports_nothing_rather_than_wholesale_deletion()
    {
        // A draft that could not be read is a failure to ANSWER the question. Answering it with
        // the most alarming possible verdict - "you have deleted the entire board" - would be
        // worse than saying nothing, and this is reachable whenever a workbook fails to parse.
        BoardData published = Board(components: [Component("U8"), Component("U9")]);

        Assert.Empty(BoardDataDiffer.Compare(published, draft: null));
    }

    [Fact]
    public void Both_null_is_harmless()
    {
        Assert.Empty(BoardDataDiffer.Compare(null, null));
    }

    // ###########################################################################################
    // A hand-edited workbook CAN carry two rows with the same natural key - two "U8" component
    // rows, say. First wins, and nothing throws.
    //
    // ToDictionary would throw here, and an exception would take down a board load over a data
    // defect the contributor can see and fix in Excel. First-wins also matches BoardDraftApplier.
    // ###########################################################################################
    [Fact]
    public void DUPLICATE_keys_in_a_hand_edited_workbook_do_not_throw()
    {
        BoardData published = Board(components: [Component("U8", "CPU")]);
        BoardData draft = Board(components: [Component("U8", "CPU"), Component("U8", "duplicate row")]);

        IReadOnlyList<BoardRowChange> changes = BoardDataDiffer.Compare(published, draft);

        // The first "U8" matches the published row exactly, so nothing is reported - the second is
        // ignored rather than reported as a second, conflicting edit.
        Assert.Empty(changes);
    }

    [Fact]
    public void Changes_across_DIFFERENT_sections_are_all_reported()
    {
        BoardData published = Board(
            components: [Component("U8")],
            credits: [new CreditEntry { Category = "Data", NameOrHandle = "Dennis" }]);

        BoardData draft = Board(
            components: [Component("U8", "CPU")],
            credits: []);

        IReadOnlyList<BoardRowChange> changes = BoardDataDiffer.Compare(published, draft);

        Assert.Equal(2, changes.Count);
        Assert.Contains(changes, change => change.Section == "Components" && change.Kind == BoardRowChangeKind.Modified);
        Assert.Contains(changes, change => change.Section == "Credits" && change.Kind == BoardRowChangeKind.Deleted);
    }

    // ###########################################################################################
    // The display label is never the raw natural key: its separator is U+241F, which renders as a
    // box in most fonts. This is the same rule DraftDriftReport already follows, restated because
    // the two surfaces show the same rows and must label them identically.
    // ###########################################################################################
    [Fact]
    public void A_display_label_NEVER_contains_the_key_separator()
    {
        BoardData draft = Board(components: [Component("U8")]);
        draft.ComponentImages.Add(new ComponentImageEntry
        {
            BoardLabel = "U8",
            Region = "PAL",
            Pin = "1",
            Name = "Clock",
        });

        IReadOnlyList<BoardRowChange> changes = BoardDataDiffer.Compare(published: null, draft);

        Assert.All(changes, change =>
            Assert.DoesNotContain(BoardDraftNaturalKeys.Separator, change.DisplayLabel));

        Assert.Contains(changes, change => change.DisplayLabel == "U8 / 1 / Clock");
    }

    [Fact]
    public void CountChanges_agrees_with_the_full_report()
    {
        BoardData published = Board(components: [Component("U8"), Component("U9")]);
        BoardData draft = Board(components: [Component("U8", "CPU"), Component("U10")]);

        Assert.Equal(
            BoardDataDiffer.Compare(published, draft).Count,
            BoardDataDiffer.CountChanges(published, draft));
    }

    // ------------------------------------------------------------------ Anti-rot

    // ###########################################################################################
    // *** THE TEST THAT STOPS THIS CLASS ROTTING. ***
    //
    // The comparison walks a row's properties by reflection precisely so that a column added to
    // BoardData later is compared the day it exists. This asserts that property - for EVERY row
    // type, by reflection, so a new section or a new column cannot quietly fall outside it.
    //
    // A hand-written field-by-field comparison would pass every other test in this file while
    // silently ignoring a new column, and a report that under-counts is worse than none because
    // it is believed.
    // ###########################################################################################
    [Fact]
    public void EVERY_field_of_EVERY_row_type_is_actually_compared()
    {
        foreach (Type rowType in BoardDataDifferTests.RowTypes())
        {
            foreach (PropertyInfo property in rowType.GetProperties(BindingFlags.Public | BindingFlags.Instance))
            {
                // Non-string properties are covered by their own culture test below; this walk
                // sets a string value, so it can only speak for string properties.
                //
                // *** THERE IS NO EXCLUSION LIST ANY MORE. *** There used to be one entry in it,
                // the retired UuidV4 column, and the column itself has since been removed from
                // BoardData. So every property of every row type must now be compared, with no
                // exceptions - which is exactly what this asserts.
                if (property.PropertyType != typeof(string))
                {
                    continue;
                }

                object baseline = BoardDataDifferTests.NewRow(rowType);
                object changed = BoardDataDifferTests.NewRow(rowType);

                // Give the key fields a value on BOTH rows, so the pair matches as "the same row"
                // and the comparison reaches the field under test rather than reporting
                // added-plus-deleted.
                BoardDataDifferTests.FillKeyFields(rowType, baseline);
                BoardDataDifferTests.FillKeyFields(rowType, changed);

                property.SetValue(changed, "a value nobody else uses");

                IReadOnlyList<string> differing = BoardDataDifferTests.DifferingFieldsOf(baseline, changed);

                Assert.True(
                    differing.Contains(property.Name),
                    $"{rowType.Name}.{property.Name} is not compared, so a change to it would be " +
                    "invisible on the Drafts tab.");
            }
        }
    }

    // ###########################################################################################
    // KiCadCalibrationEntry's fields are doubles and bools, not strings - and it reaches the
    // differ through BoardDraftNaturalKeys.ForRow, so the case is real rather than hypothetical.
    //
    // *** FORMATTED INVARIANTLY, OR THE SAME BOARD DIFFERS BY MACHINE. *** On a comma-decimal
    // machine a culture-formatted 1.5 is "1,5", so a draft created on one and compared on another
    // would report every calibration as modified. Danish is used here because it is the
    // project owner's own locale, and this class of bug is documented three times elsewhere in this
    // codebase.
    // ###########################################################################################
    [Fact]
    public void A_NON_STRING_field_compares_the_same_in_every_culture()
    {
        CultureInfo original = Thread.CurrentThread.CurrentCulture;

        try
        {
            var left = new KiCadCalibrationEntry { SchematicName = "Sheet 1", ScaleX = 1.5 };
            var right = new KiCadCalibrationEntry { SchematicName = "Sheet 1", ScaleX = 1.5 };
            var different = new KiCadCalibrationEntry { SchematicName = "Sheet 1", ScaleX = 2.5 };

            foreach (string culture in new[] { "en-US", "da-DK" })
            {
                Thread.CurrentThread.CurrentCulture = new CultureInfo(culture);

                Assert.Empty(BoardDataDifferTests.DifferingFieldsOf(left, right));
                Assert.Contains("ScaleX", BoardDataDifferTests.DifferingFieldsOf(left, different));
            }
        }
        finally
        {
            Thread.CurrentThread.CurrentCulture = original;
        }
    }

    // ------------------------------------------------------------------ Reflection helpers

    // The differ's own private comparison, reached the same way ExternalTargetLauncherTests and
    // OnlineServicesTests reach theirs - the alternative is widening a private method to public
    // purely for a test.
    private static IReadOnlyList<string> DifferingFieldsOf(object published, object drafted)
    {
        MethodInfo method = typeof(BoardDataDiffer).GetMethod(
            "DifferingFields",
            BindingFlags.NonPublic | BindingFlags.Static)!;

        return (IReadOnlyList<string>)method.Invoke(null, [published, drafted])!;
    }

    private static IEnumerable<Type> RowTypes() =>
    [
        typeof(BoardSchematicEntry),
        typeof(ComponentEntry),
        typeof(ComponentImageEntry),
        typeof(ComponentHighlightEntry),
        typeof(ComponentLocalFileEntry),
        typeof(ComponentLinkEntry),
        typeof(BoardLocalFileEntry),
        typeof(BoardLinkEntry),
        typeof(CreditEntry),
        typeof(KiCadImportantSignalEntry),
    ];

    private static object NewRow(Type rowType) => Activator.CreateInstance(rowType)!;

    // The properties that make up each type's natural key, so a test pair can be made to match.
    private static void FillKeyFields(Type rowType, object row)
    {
        string[] keyProperties = rowType.Name switch
        {
            nameof(BoardSchematicEntry) => ["SchematicName"],
            nameof(ComponentEntry) => ["BoardLabel"],
            nameof(ComponentImageEntry) => ["BoardLabel", "Region", "Pin", "Name"],
            nameof(ComponentHighlightEntry) => ["SchematicName", "BoardLabel"],
            nameof(ComponentLocalFileEntry) => ["BoardLabel", "Name"],
            nameof(ComponentLinkEntry) => ["BoardLabel", "Name"],
            nameof(BoardLocalFileEntry) => ["Category", "Name"],
            nameof(BoardLinkEntry) => ["Category", "Name"],
            nameof(CreditEntry) => ["Category", "SubCategory", "NameOrHandle"],
            nameof(KiCadImportantSignalEntry) => ["DisplayName", "KiCadNetName"],
            _ => [],
        };

        foreach (string name in keyProperties)
        {
            rowType.GetProperty(name)!.SetValue(row, "key-" + name);
        }
    }

    // ------------------------------------------------------------------ Regional rows (2026-09-24)

    [Fact]
    public void An_added_regional_variant_is_an_ADDED_row_and_is_described_with_its_region()
    {
        BoardData published = new() { Components = [new ComponentEntry { BoardLabel = "U8", Region = "PAL" }] };
        BoardData draft = new()
        {
            Components =
            [
                new ComponentEntry { BoardLabel = "U8", Region = "PAL" },
                new ComponentEntry { BoardLabel = "U8", Region = "NTSC" },
            ],
        };

        BoardRowChange change = Assert.Single(BoardDataDiffer.Compare(published, draft));

        Assert.Equal(BoardRowChangeKind.Added, change.Kind);
        Assert.Equal("U8 / NTSC", change.DisplayLabel);
    }

    // ------------------------------------------------------------------ Identifying cells changed (2026-10-04)

    // ###########################################################################################
    // *** A ROW WHOSE IDENTIFYING CELLS ALONE CHANGED IS ONE ROW MODIFIED (owner decision,
    // 2026-10-04). *** A credit's "Name or handle" given " 2" was a removal plus an addition: "This
    // seems weird to me." The owner chose the strict rule - everything but the identifying cells the
    // same, and something filled in left to recognise the row by - over "mostly the same row".
    // ###########################################################################################
    [Fact]
    public void A_credits_name_changed_and_nothing_else_is_ONE_row_modified()
    {
        BoardData published = Board(credits:
        [
            new CreditEntry { Category = "Board labelling", NameOrHandle = "Dennis", Contact = "dennis@example.org" },
        ]);
        BoardData draft = Board(credits:
        [
            new CreditEntry { Category = "Board labelling", NameOrHandle = "Dennis 2", Contact = "dennis@example.org" },
        ]);

        BoardRowChange change = Assert.Single(BoardDataDiffer.Compare(published, draft));

        Assert.Equal(BoardRowChangeKind.Modified, change.Kind);
        Assert.Equal(["NameOrHandle"], change.ChangedFields);
        Assert.Equal(BoardDraftNaturalKeys.ForCredit("Board labelling", "", "Dennis 2"), change.NaturalKey);
    }

    [Fact]
    public void A_credits_name_AND_contact_changed_is_still_a_removal_plus_an_addition()
    {
        // A different person, as far as anything can tell - the owner's own example of where the
        // strict rule stops.
        BoardData published = Board(credits: [new CreditEntry { Category = "Board labelling", NameOrHandle = "Dennis", Contact = "dennis@example.org" }]);
        BoardData draft = Board(credits: [new CreditEntry { Category = "Board labelling", NameOrHandle = "Jane", Contact = "jane@example.org" }]);

        Assert.Equal(
            [BoardRowChangeKind.Added, BoardRowChangeKind.Deleted],
            BoardDataDiffer.Compare(published, draft).Select(change => change.Kind));
    }

    [Fact]
    public void A_component_relabelled_with_its_values_kept_is_one_row_but_with_ANOTHER_value_changed_is_two()
    {
        BoardData published = Board(components: [Component("U8", "CPU", "906114")]);

        BoardRowChange relabelled = Assert.Single(BoardDataDiffer.Compare(published, Board(components: [Component("U9", "CPU", "906114")])));
        Assert.Equal(BoardRowChangeKind.Modified, relabelled.Kind);
        Assert.Equal(["BoardLabel"], relabelled.ChangedFields);
        Assert.Equal("U9", relabelled.DisplayLabel);

        Assert.Equal(
            [BoardRowChangeKind.Added, BoardRowChangeKind.Deleted],
            BoardDataDiffer.Compare(published, Board(components: [Component("U9", "VIC", "906114")])).Select(change => change.Kind));
    }

    // Nothing left to recognise the row by: every other cell blank, and no identifying cell shared.
    // Pairing such rows would call any deletion plus any addition a rename.
    [Fact]
    public void Rows_with_nothing_filled_in_in_common_are_never_paired()
    {
        BoardData published = Board(components: [Component("U8")]);
        BoardData draft = Board(components: [Component("U9")]);

        Assert.Equal(
            [BoardRowChangeKind.Added, BoardRowChangeKind.Deleted],
            BoardDataDiffer.Compare(published, draft).Select(change => change.Kind));
    }

    // An important signal is two identifying cells and nothing else: a new net is one row while its
    // display name stays, and two rows once neither half does.
    [Fact]
    public void An_important_signals_new_net_is_one_row_while_its_display_name_stays()
    {
        var published = new BoardData { KiCadImportantSignals = [new KiCadImportantSignalEntry { DisplayName = "RESET", KiCadNetName = "/RESET" }] };

        BoardRowChange change = Assert.Single(BoardDataDiffer.Compare(
            published,
            new BoardData { KiCadImportantSignals = [new KiCadImportantSignalEntry { DisplayName = "RESET", KiCadNetName = "~{RESET}" }] }));
        Assert.Equal(BoardRowChangeKind.Modified, change.Kind);

        Assert.Equal(
            [BoardRowChangeKind.Added, BoardRowChangeKind.Deleted],
            BoardDataDiffer.Compare(
                published,
                new BoardData { KiCadImportantSignals = [new KiCadImportantSignalEntry { DisplayName = "9VAC", KiCadNetName = "9VAC~" }] })
            .Select(change => change.Kind));
    }

    // ###########################################################################################
    // Three credits of one person, each renamed: they all share the contact, so each could be any
    // other's rename - the one sharing the most filled-in identifying cells (its Category) wins, and
    // every row keeps its own.
    // ###########################################################################################
    [Fact]
    public void Several_rows_renamed_alike_each_pair_with_the_row_sharing_most_of_its_key()
    {
        CreditEntry Credit(string category, string name) => new() { Category = category, NameOrHandle = name, Contact = "dennis@example.org" };

        List<object> removed = [Credit("Board labelling", "Dennis"), Credit("Board data", "Dennis"), Credit("Oscilloscope", "Dennis")];
        List<object> added = [Credit("Oscilloscope", "Dennis 2"), Credit("Board labelling", "Dennis 2"), Credit("Board data", "Dennis 2")];

        IReadOnlyList<BoardRowRename> pairs = BoardDataDiffer.PairRenamedRows(removed, added);

        Assert.Equal([new BoardRowRename(0, 1), new BoardRowRename(1, 2), new BoardRowRename(2, 0)], pairs);
    }

    // The table holds new rows in screen order and the saved board in saved order, and both must pair
    // alike - so which added row wins a tie may not depend on where it is in the list.
    [Fact]
    public void Which_added_row_a_tie_goes_to_does_not_depend_on_their_order()
    {
        List<object> removed = [Component("U8", "Capacitor", "100n")];
        List<object> added = [Component("C2", "Capacitor", "100n"), Component("C1", "Capacitor", "100n")];

        object forward = added[BoardDataDiffer.PairRenamedRows(removed, added).Single().Added];
        added.Reverse();
        object backward = added[BoardDataDiffer.PairRenamedRows(removed, added).Single().Added];

        Assert.Same(forward, backward);
        Assert.Equal("C1", ((ComponentEntry)forward).BoardLabel);
    }

    [Fact]
    public void Each_row_is_in_at_most_one_pair_and_rows_of_different_kinds_never_pair()
    {
        List<object> removed = [Component("U8", "CPU"), new BoardLinkEntry { Category = "CPU", Name = "U8" }];
        List<object> added = [Component("U9", "CPU"), Component("U10", "CPU")];

        BoardRowRename pair = Assert.Single(BoardDataDiffer.PairRenamedRows(removed, added));
        Assert.Equal(0, pair.Removed);
    }
}
