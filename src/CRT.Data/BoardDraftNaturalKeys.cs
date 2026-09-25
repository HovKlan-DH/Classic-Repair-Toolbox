namespace Handlers.DataHandling
{
    // ###########################################################################################
    // The one place a BoardData row's identity is decided for drafting and (later) for
    // base-revision diffing at submission time - see NewContributeStrategy.md, "Retiring the
    // UUIDs". A row's UuidV4 is never used here: the whole point of base-revision diffing is that
    // row identity does not have to live in the data, so a draft (and, later, a server-side diff)
    // pairs rows on the same natural keys a human would use to describe "the same row" - the board
    // label, the schematic name, and so on.
    //
    // Every key is built the same way: normalize (trim, treat null as empty), join with a
    // separator that cannot appear in a board label or schematic name in practice, compare
    // case-insensitively at the call site (BoardDraftApplier does the comparison; this class only
    // builds the string). Two rows with the same key are the same row for drafting purposes.
    // ###########################################################################################
    public static class BoardDraftNaturalKeys
    {
        // Not a character any board label, schematic name, region or file name legitimately
        // contains, and distinct from '/' and '\' which board-relative paths already use.
        // public because a key sometimes has to be split back apart for DISPLAY - a Deleted
        // tombstone carries no Entry to read real field values off, so the key's own parts are the
        // only thing left to show (see DraftDriftDetector). Exposed rather than copied, so there is
        // still exactly one definition of what separates a key's parts.
        //
        // It renders as a box in most fonts, so it must never reach the screen as-is.
        public const string Separator = "␟";

        public static string ForSchematic(string schematicName) =>
            Normalize(schematicName);

        // ###########################################################################################
        // A component row is its board label PLUS ITS REGION (changed 2026-09-24).
        //
        // A regionalised component is stored as one row per region, all sharing a label - U1 for
        // PAL and U1 for NTSC - which is an established shape in the data (see
        // ComponentListBuilder.BuildComponentsInScope). Keyed on the label alone, those two rows
        // were ONE row to every comparison: adding the NTSC row next to an existing PAL one was
        // no change at all to the Drafts tab's count, the table editor flagged it as a duplicate
        // instead of colouring it added, and a maintainer's summary never showed it. Reported by
        // the project owner from the table editor. Component images already key on region too.
        //
        // With NO region the key is the bare label, exactly as before, so every row that is not
        // regionalised - nearly all of them - keys identically to how it always did. The
        // consequence worth knowing: changing a row's region is now a delete plus an add, the same
        // way changing its label is.
        // ###########################################################################################
        public static string ForComponent(string boardLabel, string? region = null) =>
            Normalize(region).Length == 0
                ? Normalize(boardLabel)
                : Join(boardLabel, region!);

        // Matches NewContributeStrategy.md's stated key for a component image: label + region +
        // pin + name. All four together, because a single board label can carry more than one
        // oscilloscope capture (different pins, different regions of the same net).
        public static string ForComponentImage(string boardLabel, string region, string pin, string name) =>
            Join(boardLabel, region, pin, name);

        // Matches NewContributeStrategy.md's stated key for a highlight: schematic + label. A
        // board label can appear on more than one schematic page (e.g. a connector shown on both
        // a main board and a daughterboard page), so schematic name must be part of the key too.
        public static string ForComponentHighlight(string schematicName, string boardLabel) =>
            Join(schematicName, boardLabel);

        public static string ForComponentLocalFile(string boardLabel, string name) =>
            Join(boardLabel, name);

        public static string ForComponentLink(string boardLabel, string name) =>
            Join(boardLabel, name);

        public static string ForBoardLocalFile(string category, string name) =>
            Join(category, name);

        public static string ForBoardLink(string category, string name) =>
            Join(category, name);

        public static string ForCredit(string category, string subCategory, string nameOrHandle) =>
            Join(category, subCategory, nameOrHandle);

        public static string ForKiCadImportantSignal(string displayName) =>
            Normalize(displayName);

        // ###########################################################################################
        // The key for one OFFICIAL BoardData row, dispatched on its type - the other side of the
        // pairing from the ForX methods above, which build a key from a drafted row's own fields.
        //
        // This lives here rather than in BoardDraftApplier (where it began life as a private
        // KeyOf<T>) because this class's own header already calls itself "the one place a BoardData
        // row's identity is decided", and having the official-row side of that decision live
        // somewhere else contradicted it. BoardDraftApplier now delegates here, and the drift
        // detector uses the same dispatch - so a new section, or a changed key shape, is a change in
        // exactly one place instead of being silently correct in one caller and wrong in another.
        //
        // Throws for an unknown row type rather than returning a blank key: a blank would silently
        // pair unrelated rows together during a merge, which is far worse than a loud failure.
        // ###########################################################################################
        public static string ForRow(object row) => row switch
        {
            BoardSchematicEntry schematic => ForSchematic(schematic.SchematicName),
            ComponentEntry component => ForComponent(component.BoardLabel, component.Region),
            ComponentImageEntry image => ForComponentImage(image.BoardLabel, image.Region, image.Pin, image.Name),
            ComponentHighlightEntry highlight => ForComponentHighlight(highlight.SchematicName, highlight.BoardLabel),
            ComponentLocalFileEntry componentFile => ForComponentLocalFile(componentFile.BoardLabel, componentFile.Name),
            ComponentLinkEntry componentLink => ForComponentLink(componentLink.BoardLabel, componentLink.Name),
            BoardLocalFileEntry boardFile => ForBoardLocalFile(boardFile.Category, boardFile.Name),
            BoardLinkEntry boardLink => ForBoardLink(boardLink.Category, boardLink.Name),
            CreditEntry credit => ForCredit(credit.Category, credit.SubCategory, credit.NameOrHandle),
            KiCadImportantSignalEntry signal => ForKiCadImportantSignal(signal.DisplayName),

            // ONE calibration per schematic, so the schematic name alone is the key - which is
            // also what KiCadCalibrationDraftWriter builds when it saves one, via ForSchematic.
            // Added 2026-09-22 when calibrations began travelling with submissions and the review
            // summary needed to pair them; before that nothing called ForRow with one and this
            // threw.
            KiCadCalibrationEntry calibration => ForSchematic(calibration.SchematicName),
            _ => throw new System.NotSupportedException($"No natural key rule for row type [{row?.GetType().Name ?? "null"}].")
        };

        private static string Join(params string[] parts)
        {
            for (int i = 0; i < parts.Length; i++)
            {
                parts[i] = Normalize(parts[i]);
            }

            return string.Join(Separator, parts);
        }

        private static string Normalize(string? value) => value?.Trim() ?? string.Empty;
    }
}
