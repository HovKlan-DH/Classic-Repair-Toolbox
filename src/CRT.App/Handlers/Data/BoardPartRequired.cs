using System;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // A FILE ADDED TO THE BOARD COMES WITH EVERY FIELD IT NEEDS FILLED IN (owner request, 2026-10-30: "make
    // sure that as minimum then 'What' and 'Name or handle' must be set ... In 'Links' then all 3 fields must
    // be filled ... Same for 'Files'"): "Add file" copies the file into the draft at once, and a row with an
    // empty field would leave it cited by a row Submit refuses.
    //
    // What a row of the board's own lists needs is CRT.Data's BoardRowRequired - an error on the cell
    // everywhere, the server's refusal too; the pane marks the empty box "Required" (owner request,
    // 2026-11-01). This class named those fields by the pane's headings too (EmptyOf) until 2026-11-15, when
    // that was found to have no caller but its own tests, and went.
    // ###########################################################################################
    internal static class BoardPartRequired
    {
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
    }
}
