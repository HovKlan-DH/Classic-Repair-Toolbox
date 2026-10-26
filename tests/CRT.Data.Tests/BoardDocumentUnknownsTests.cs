using System.Collections;
using System.Reflection;
using System.Text.Json;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// BoardDocumentUnknowns - what a board file holds that this CRT does not know survives the file
// being written again from BoardData (code review, 2026-10-25). Only the top-level fields did; a
// field a later CRT adds to a row was dropped from every row by the first edit an installed CRT
// 4.0.0 made, and that cannot be fixed in a CRT already installed.
// ###########################################################################################
public sealed class BoardDocumentUnknownsTests
{
    private static Dictionary<string, JsonElement> Later(object value) =>
        new() { ["laterField"] = JsonSerializer.SerializeToElement(value) };

    private static BoardDocument.ComponentRow Component(string label, string description = "", string region = "") =>
        new() { BoardLabel = label, FriendlyName = "CPU", Description = description, Region = region };

    // ###########################################################################################
    // EVERY list of a board file is carried - checked by reflection over BoardDocument, so a list
    // added later without a line in CarryOver fails here instead of losing a later CRT's fields.
    // ###########################################################################################
    [Fact]
    public void Every_row_of_every_list_keeps_its_unknown_fields_when_the_file_is_built_again()
    {
        var previous = new BoardDocument { Unknown = Later("top") };
        List<PropertyInfo> lists = typeof(BoardDocument).GetProperties()
            .Where(property => property.PropertyType.IsGenericType && property.PropertyType.GetGenericTypeDefinition() == typeof(List<>))
            .ToList();

        foreach (PropertyInfo list in lists)
        {
            var row = (IKeepsUnknownFields)Activator.CreateInstance(list.PropertyType.GetGenericArguments()[0])!;
            row.Unknown = Later(list.Name);
            ((IList)list.GetValue(previous)!).Add(row);
        }

        // The same rows as this CRT builds them: everything it knows, nothing it does not.
        BoardDocument next = DataTreeJson.ReadBoard(DataTreeJson.Write(previous));
        next.Unknown = null;

        foreach (PropertyInfo list in lists)
        {
            foreach (IKeepsUnknownFields row in (IList)list.GetValue(next)!)
                row.Unknown = null;
        }

        BoardDocumentUnknowns.CarryOver(previous, next);

        Assert.Equal("top", next.Unknown!["laterField"].GetString());

        foreach (PropertyInfo list in lists)
        {
            IKeepsUnknownFields row = Assert.Single(((IList)list.GetValue(next)!).Cast<IKeepsUnknownFields>());
            Assert.Equal(list.Name, row.Unknown?["laterField"].GetString());
        }
    }

    // A component whose description was edited is still that component (its label and region).
    [Fact]
    public void An_edited_row_keeps_the_unknown_fields_of_the_row_it_was()
    {
        BoardDocument.ComponentRow before = Component("U8", "old text");
        before.Unknown = Later(40);
        var previous = new BoardDocument { Components = [before] };
        var next = new BoardDocument { Components = [Component("u8", "new text")] };

        BoardDocumentUnknowns.CarryOver(previous, next);

        Assert.Equal(40, next.Components[0].Unknown!["laterField"].GetInt32());
    }

    // Relabelled, or in another region, it is another component - nothing is guessed onto it.
    [Fact]
    public void A_row_whose_identity_changed_and_a_new_row_take_nothing()
    {
        BoardDocument.ComponentRow before = Component("U8");
        before.Unknown = Later(40);
        var previous = new BoardDocument { Components = [before] };
        var next = new BoardDocument { Components = [Component("U9"), Component("U8", region: "PAL")] };

        BoardDocumentUnknowns.CarryOver(previous, next);

        Assert.All(next.Components, row => Assert.Null(row.Unknown));
    }

    // Of two rows alike in everything this CRT knows, each keeps its own - not the first's twice.
    [Fact]
    public void Two_rows_alike_in_every_known_value_each_keep_their_own()
    {
        BoardDocument.ComponentRow first = Component("U8");
        BoardDocument.ComponentRow second = Component("U8");
        first.Unknown = Later("first");
        second.Unknown = Later("second");
        var previous = new BoardDocument { Components = [first, second] };
        var next = new BoardDocument { Components = [Component("U8"), Component("U8")] };

        BoardDocumentUnknowns.CarryOver(previous, next);

        Assert.Equal(["first", "second"], next.Components.Select(row => row.Unknown!["laterField"].GetString()));
    }

    // The exact row wins over a row that only shares the identity.
    [Fact]
    public void The_row_with_the_same_values_is_matched_before_one_with_the_same_identity()
    {
        BoardDocument.ComponentRow edited = Component("U8", "was edited");
        BoardDocument.ComponentRow untouched = Component("U8", "untouched");
        edited.Unknown = Later("edited");
        untouched.Unknown = Later("untouched");
        var previous = new BoardDocument { Components = [edited, untouched] };
        var next = new BoardDocument { Components = [Component("U8", "edited again"), Component("U8", "untouched")] };

        BoardDocumentUnknowns.CarryOver(previous, next);

        Assert.Equal(["edited", "untouched"], next.Components.Select(row => row.Unknown!["laterField"].GetString()));
    }

    [Fact]
    public void A_field_the_new_file_already_has_is_never_overwritten()
    {
        BoardDocument.ComponentRow before = Component("U8");
        before.Unknown = Later("old");
        BoardDocument.ComponentRow after = Component("U8");
        after.Unknown = Later("new");

        var next = new BoardDocument { Components = [after], Unknown = Later("new top") };

        BoardDocumentUnknowns.CarryOver(new BoardDocument { Components = [before], Unknown = Later("old top") }, next);

        Assert.Equal("new", next.Components[0].Unknown!["laterField"].GetString());
        Assert.Equal("new top", next.Unknown!["laterField"].GetString());
    }

    // A file's name is compared without regard to capitals, like every name of a board's files.
    [Fact]
    public void A_listed_file_keeps_its_unknown_fields_by_its_name()
    {
        var before = new BoardDocument.StoredFileRow { Name = "Images/Top.png", Sha256 = "old", Unknown = Later("kept") };
        var after = new BoardDocument.StoredFileRow { Name = "images/top.png", Sha256 = "new" };

        BoardDocumentUnknowns.CarryOver(new BoardDocument { Files = [before] }, new BoardDocument { Files = [after] });

        Assert.Equal("kept", after.Unknown!["laterField"].GetString());
    }

    [Fact]
    public void With_no_file_before_there_is_nothing_to_carry()
    {
        var next = new BoardDocument { Components = [Component("U8")] };

        BoardDocumentUnknowns.CarryOver(null, next);

        Assert.Null(next.Unknown);
        Assert.Null(next.Components[0].Unknown);
    }
}
