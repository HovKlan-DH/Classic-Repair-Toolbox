using System.Collections.Generic;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // Which text a search query is matched against, for each kind of worklog record.
    //
    // ONLY fields a user typed are included. Numbers are deliberately left out - ids, hours spent,
    // cost, display order - because a search for "2" would otherwise match half the database
    // through entry ids and costs nobody was searching for. Dates are out for the same reason
    // (they are rendered, not typed) and because a date's searchable form depends on a format the
    // user never chose.
    //
    // Category IS included - "Note"/"Cosmetic"/"Issue" are descriptive words the user picked and
    // reads back on screen, and none of them is a word that turns up incidentally in repair notes.
    //
    // Status/State ("Open"/"Closed") are deliberately NOT. They are two-valued and every record
    // carries one, so including them made "open" - a word that appears constantly in this domain
    // ("open circuit", "opened the case", "open trace on CN2") - match nearly the whole database,
    // and because terms are ANDed across a record, "open trace" then matched any Open workbook
    // mentioning "trace" anywhere. Both values already have their own always-visible pill, so they
    // filter by eye far better than they ever did by substring.
    //
    // File names are included: the user chose the file, and remembering "the photo called
    // psu-rail" is a realistic way to find an entry again.
    // ###########################################################################################
    public static class WorklogSearchIndex
    {
        // ###########################################################################################
        // The workbook's own text - what the left-hand card shows. Does NOT include its entries;
        // callers that want "workbook matches, or any of its entries does" combine the two, so that
        // each can also be reported separately (the workbook card highlights its own fields, an
        // entry highlights its own).
        // ###########################################################################################
        public static IEnumerable<string?> ForWorkbook(WorkbookRecord workbook)
        {
            if (workbook == null)
                yield break;

            yield return workbook.Title;
            yield return workbook.Note;
        }

        // ###########################################################################################
        // One entry's full text, INCLUDING every sub-list the editor holds (links, comments, work
        // done, photos, files) and the component labels it has in scope. Searching the worklog is
        // expected to reach the comment somebody left three months ago, not just the headline.
        // ###########################################################################################
        public static IEnumerable<string?> ForEntry(WorklogEntryRecord entry)
        {
            if (entry == null)
                yield break;

            yield return entry.Title;
            yield return entry.Description;
            yield return entry.Category;
            yield return entry.SchematicName;

            foreach (var label in entry.ComponentLabels)
                yield return label;

            foreach (var link in entry.Links)
            {
                yield return link.Headline;
                yield return link.Url;
            }

            foreach (var comment in entry.Comments)
                yield return comment.Text;

            // The note only - HoursSpent and Cost are numbers, see the class header.
            foreach (var workDone in entry.WorkDoneItems)
                yield return workDone.Text;

            foreach (var photo in entry.Photos)
            {
                yield return photo.FileName;
                yield return photo.Comment;
            }

            foreach (var file in entry.Files)
            {
                yield return file.FileName;
                yield return file.Comment;
            }
        }
    }
}
