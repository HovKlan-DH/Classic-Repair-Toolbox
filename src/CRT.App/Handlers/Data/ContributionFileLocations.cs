namespace Handlers.DataHandling
{
    // ###########################################################################################
    // The folder a picked file goes in - a folder of the BOARD'S OWN FILE NAMES (CRT 4.0, 2026-10-25).
    //
    // *** IN 4.0 A ROW NAMES A FILE OF ITS BOARD *** ("Scope baseline/U19_1_PAL.png" -
    // BoardFileNames), and the bytes lie in a store laid out by checksum ("Files/ab/<sha256>.png").
    // So the data tree's folders say nothing about where a file of the board goes: it goes where every
    // converted board keeps that kind of file (the trees measured 2026-10-25: schematics at the top,
    // then exactly these four folders, and "Truth tables" on two boards).
    // ###########################################################################################
    internal static class ContributionFileLocations
    {
        // The kinds of file row a board has.
        public enum FileRowKind
        {
            ComponentPicture,
            ScopeBaseline,
            ComponentFile,
            BoardFile
        }

        // ###########################################################################################
        // *** A PICKED FILE GOES IN ITS KIND'S USUAL FOLDER, NEVER AS A BARE NAME (owner report,
        // 2026-09-25: the file sat at the top of the board, beside its schematics) *** - a picture in
        // "Component images", a scope baseline in "Scope baseline", a component's file in "Component
        // local files", a board's in "Board local files".
        // ###########################################################################################
        public static string DefaultFolderFor(FileRowKind kind) => kind switch
        {
            FileRowKind.ComponentPicture => "Component images",
            FileRowKind.ScopeBaseline => DraftFolderLayout.ScopeBaselineFolderName,
            FileRowKind.ComponentFile => "Component local files",
            _ => "Board local files",
        };
    }
}
