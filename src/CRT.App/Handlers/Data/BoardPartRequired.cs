using System;
using System.Collections.Generic;
using System.Linq;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // WHAT A ROW OF THE BOARD'S OWN LISTS MUST HAVE FILLED IN, AS THE COMPONENT PANE NAMES IT (owner
    // request, 2026-10-30: "make sure that as minimum then 'What' and 'Name or handle' must be set ...
    // In 'Links' then all 3 fields must be filled ... Same for 'Files' ... In 'Important signals' the
    // 'Important signal name' must not be empty"):
    //
    //   - a FILE: its category, its name and its file;
    //   - a LINK: its category, its name and its address;
    //   - a CREDIT: what it is for ("What") and the name or handle - which part of it and the contact
    //     may stay empty;
    //   - an IMPORTANT SIGNAL: its KiCad net and its name.
    //
    // The rule is CRT.Data's BoardRowRequired - an error on the cell everywhere, the server's refusal
    // too; this names each field by its heading in the pane. Until 2026-11-01 the pane kept such a row
    // out of the draft with a line under it; now the row goes in as it is, the empty box marked
    // "Required" (owner request: "highlight the problematic cell with a helper what is missing ...
    // These errors must be flagged in 'Board status' and then it should not be possible to submit").
    // ###########################################################################################
    internal static class BoardPartRequired
    {
        // The headings the pane shows over each list - the names a line says.
        public const string Category = "Category";
        public const string Name = "Name";
        public const string File = "File";
        public const string Address = "Address";
        public const string What = "What";
        public const string NameOrHandle = "Name or handle";
        public const string KiCadNet = "KiCad net";
        public const string SignalName = "Important signal name";

        // A board's FIRST file added gets this category - "Add file" makes a row with every field filled
        // in, or its copied file would lie in the draft cited by nothing. The live boards' most used, and
        // the category box's placeholder.
        public const string FirstFileCategory = "Resources";

        // The category, the name and the file of a file added to the board as `name` from `source` - in
        // the last row's category, so it joins a group (FirstFileCategory when there is none), named
        // after the file picked.
        public static BoardLocalFileEntry NewFile(string source, string name, string? lastCategory)
        {
            ArgumentNullException.ThrowIfNull(source);
            ArgumentNullException.ThrowIfNull(name);

            string fileName = System.IO.Path.GetFileName(source);
            string shown = System.IO.Path.GetFileNameWithoutExtension(source);

            return new BoardLocalFileEntry
            {
                Category = string.IsNullOrWhiteSpace(lastCategory) ? FirstFileCategory : lastCategory.Trim(),
                Name = string.IsNullOrWhiteSpace(shown) ? fileName : shown,
                File = name,
            };
        }

        // The required fields of `row` that are empty, in the order the pane shows them - BoardRowRequired's
        // columns by the pane's headings.
        public static IReadOnlyList<string> EmptyOf(BoardLocalFileEntry row) =>
            BoardPartRequired.Named(BoardRowRequired.EmptyColumns(row), (BoardWorkbookSchema.ColCategory, Category), (BoardWorkbookSchema.ColName, Name), (BoardWorkbookSchema.ColFile, File));

        public static IReadOnlyList<string> EmptyOf(BoardLinkEntry row) =>
            BoardPartRequired.Named(BoardRowRequired.EmptyColumns(row), (BoardWorkbookSchema.ColCategory, Category), (BoardWorkbookSchema.ColName, Name), (BoardWorkbookSchema.ColUrl, Address));

        public static IReadOnlyList<string> EmptyOf(CreditEntry row) =>
            BoardPartRequired.Named(BoardRowRequired.EmptyColumns(row), (BoardWorkbookSchema.ColCategory, What), (BoardWorkbookSchema.ColNameOrHandle, NameOrHandle));

        public static IReadOnlyList<string> EmptyOf(KiCadImportantSignalEntry row) =>
            BoardPartRequired.Named(BoardRowRequired.EmptyColumns(row), (BoardWorkbookSchema.ColKiCadNetName, KiCadNet), (BoardWorkbookSchema.ColDisplayName, SignalName));

        private static List<string> Named(IReadOnlyList<string> columns, params (string Column, string Heading)[] headings) =>
            columns.Select(column => headings.First(heading => heading.Column == column).Heading).ToList();
    }
}
