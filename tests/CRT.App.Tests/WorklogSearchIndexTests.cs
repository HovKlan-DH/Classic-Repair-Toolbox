using System;
using System.Collections.Generic;
using System.Linq;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// Tests for WorklogSearchIndex - which fields the "Find a previous repair" box searches.
//
// The rule these pin down is the one that is easy to break by accident: user-typed TEXT is
// searchable, numbers are not. A search for "2" that matched every entry through its id, hours or
// cost would make the box useless on exactly the short queries people type first.
public class WorklogSearchIndexTests
{
    [Fact]
    public void A_workbooks_searchable_text_is_its_title_and_note()
    {
        var workbook = new WorkbookRecord
        {
            Id = 7,
            BoardKey = "Commodore 64|250469",
            Title = "Mr Jensens C64",
            Note = "collected tuesday",
            Status = "Open",
            WorklogCount = 12,
        };

        var fields = WorklogSearchIndex.ForWorkbook(workbook).ToList();

        Assert.Contains("Mr Jensens C64", fields);
        Assert.Contains("collected tuesday", fields);

        // The id, board key and entry count are not searchable text - see the class header.
        Assert.DoesNotContain("7", fields);
        Assert.DoesNotContain("12", fields);
        Assert.DoesNotContain("Commodore 64|250469", fields);
    }

    [Fact]
    public void An_entrys_searchable_text_covers_its_own_fields()
    {
        var entry = new WorklogEntryRecord
        {
            Id = 3,
            Title = "No picture",
            Description = "black screen on boot",
            Category = "Issue",
            State = "Closed",
            SchematicName = "Video circuit",
        };

        var fields = WorklogSearchIndex.ForEntry(entry).ToList();

        Assert.Contains("No picture", fields);
        Assert.Contains("black screen on boot", fields);
        Assert.Contains("Video circuit", fields);

        // Category is a fixed vocabulary but a descriptive one, and none of its three values turns
        // up incidentally in repair notes - so searching "Issue" is meaningful.
        Assert.Contains("Issue", fields);
    }

    // Status/State are excluded on purpose. Every record carries one of two values, and "open" is a
    // word this domain uses constantly ("open circuit", "opened the case"), so including them made
    // that search match almost everything - see the class header.
    [Fact]
    public void Status_and_state_are_not_searchable_so_the_word_open_stays_useful()
    {
        var workbook = new WorkbookRecord { Id = 1, Title = "Mr Jensens C64", Status = "Open" };
        var entry = new WorklogEntryRecord { Title = "No picture", Description = "dead VIC", State = "Open" };

        Assert.DoesNotContain("Open", WorklogSearchIndex.ForWorkbook(workbook));
        Assert.DoesNotContain("Open", WorklogSearchIndex.ForEntry(entry));

        // So a search for "open" finds only records that actually SAY it.
        var query = WorklogSearchQuery.Parse("open");

        Assert.False(query.Matches(WorklogSearchIndex.ForWorkbook(workbook)));
        Assert.False(query.Matches(WorklogSearchIndex.ForEntry(entry)));

        var openCircuit = new WorklogEntryRecord { Title = "Open circuit at CN2", State = "Closed" };
        Assert.True(query.Matches(WorklogSearchIndex.ForEntry(openCircuit)));
    }

    [Fact]
    public void An_entrys_searchable_text_reaches_into_every_sub_list()
    {
        // The point of searching a worklog is to find the comment somebody left months ago, not
        // just the headline - so every list the editor holds has to be reachable.
        var entry = new WorklogEntryRecord
        {
            Title = "Recap",
            ComponentLabels = new List<string> { "C64", "U18" },
            Links = new List<WorklogLinkRecord>
            {
                new() { Headline = "datasheet", Url = "https://example.com/6510" },
            },
            Comments = new List<WorklogCommentRecord>
            {
                new() { Text = "smells of burnt tantalum", Date = DateTime.Now },
            },
            WorkDoneItems = new List<WorklogWorkDoneRecord>
            {
                new() { Text = "replaced every electrolytic", HoursSpent = 2.5, Cost = 140 },
            },
            Photos = new List<WorklogAttachmentRecord>
            {
                new() { FileName = "psu-rail.jpg", Comment = "before cleaning" },
            },
            Files = new List<WorklogAttachmentRecord>
            {
                new() { FileName = "scope-trace.csv", Comment = "dot clock" },
            },
        };

        var fields = WorklogSearchIndex.ForEntry(entry).ToList();

        Assert.Contains("U18", fields);
        Assert.Contains("datasheet", fields);
        Assert.Contains("https://example.com/6510", fields);
        Assert.Contains("smells of burnt tantalum", fields);
        Assert.Contains("replaced every electrolytic", fields);

        // File names count as searchable: the user chose the file, and "the photo called psu-rail"
        // is a realistic way to find an entry again.
        Assert.Contains("psu-rail.jpg", fields);
        Assert.Contains("before cleaning", fields);
        Assert.Contains("scope-trace.csv", fields);
        Assert.Contains("dot clock", fields);
    }

    // Searching for a component reference finds the worklogs that have it in scope. This is worth
    // its own test rather than resting on the sub-list one above, because it is the example the
    // Wiki teaches ("u8" -> every repair that touched U8) and the whole reason it works is that
    // ComponentLabels is in the index at all - a field someone could reasonably think of as
    // structured board data rather than as user-typed text, and drop.
    //
    // It is genuinely user-chosen: the labels are whatever the user ticked in the editor's
    // "components in scope" checklist.
    [Fact]
    public void A_component_in_scope_is_findable_by_its_reference()
    {
        var entry = new WorklogEntryRecord
        {
            Title = "Replaced the PLA",
            ComponentLabels = new List<string> { "U8", "U17" },
        };

        var query = WorklogSearchQuery.Parse("u8");

        // Case-insensitive, like every other term.
        Assert.True(query.Matches(WorklogSearchIndex.ForEntry(entry)));

        // And an entry that never scoped it is not dragged in.
        var other = new WorklogEntryRecord
        {
            Title = "Recapped the PSU",
            ComponentLabels = new List<string> { "C38" },
        };

        Assert.False(query.Matches(WorklogSearchIndex.ForEntry(other)));
    }

    // The caveat that goes with the example above: matching is substring, so a short reference is
    // a PREFIX of the longer ones and pulls them in too. "U8" finds U8, but also U80 and U81.
    //
    // This is not a defect to fix - narrowing it to whole words would break "cap" finding
    // "capacitor", which is the far more common way the box is used - but it is the reason the Wiki
    // tells the reader to add another word rather than to trust a two-character search.
    [Fact]
    public void A_short_component_reference_also_matches_longer_ones()
    {
        var longerLabel = new WorklogEntryRecord
        {
            Title = "Address decoding",
            ComponentLabels = new List<string> { "U80" },
        };

        Assert.True(WorklogSearchQuery.Parse("u8").Matches(WorklogSearchIndex.ForEntry(longerLabel)));

        // Adding a second term is what narrows it, which is what the documentation recommends.
        Assert.False(
            WorklogSearchQuery.Parse("u8 pla").Matches(WorklogSearchIndex.ForEntry(longerLabel)));
    }

    [Fact]
    public void Numbers_are_not_searchable_text()
    {
        var entry = new WorklogEntryRecord
        {
            Id = 2,
            Title = "Recap",
            AreaX = 12,
            AreaY = 34,
            WorkDoneItems = new List<WorklogWorkDoneRecord>
            {
                new() { Text = "replaced C16", HoursSpent = 2, Cost = 250 },
            },
            Photos = new List<WorklogAttachmentRecord>
            {
                new() { FileName = "a.jpg", Comment = "", DisplayOrder = 2 },
            },
        };

        var fields = WorklogSearchIndex.ForEntry(entry)
            .Where(f => !string.IsNullOrEmpty(f))
            .ToList();

        // Hours (2), cost (250), the entry id (2), the area coordinates and the display order must
        // not leak in as searchable strings - a search for "2" would otherwise match this entry
        // four times over without the word "2" appearing anywhere the user can see.
        Assert.DoesNotContain("2", fields);
        Assert.DoesNotContain("250", fields);
        Assert.DoesNotContain("12", fields);
        Assert.DoesNotContain("34", fields);

        // The work-done NOTE is still searchable - only its numbers were dropped.
        Assert.Contains("replaced C16", fields);
    }

    [Fact]
    public void A_null_record_yields_no_fields_rather_than_throwing()
    {
        Assert.Empty(WorklogSearchIndex.ForWorkbook(null!));
        Assert.Empty(WorklogSearchIndex.ForEntry(null!));
    }

    // The two halves compose: a query is run against a workbook's own text plus its entries' text,
    // which is how "workbook matches, or any of its entries does" is expressed at the call site.
    [Fact]
    public void A_query_matches_a_workbook_through_one_of_its_entries()
    {
        var workbook = new WorkbookRecord { Id = 1, Title = "Mr Jensens C64", Status = "Open" };
        var entry = new WorklogEntryRecord { Title = "No picture", Description = "dead VIC" };

        var query = WorklogSearchQuery.Parse("vic");

        Assert.False(query.Matches(WorklogSearchIndex.ForWorkbook(workbook)));
        Assert.True(query.Matches(
            WorklogSearchIndex.ForWorkbook(workbook).Concat(WorklogSearchIndex.ForEntry(entry))));
    }
}
