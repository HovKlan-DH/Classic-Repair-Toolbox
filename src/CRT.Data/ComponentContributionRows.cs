namespace Handlers.DataHandling
{
    // ###########################################################################################
    // THE ROW SHAPES THE COMPONENT-CONTRIBUTION WINDOW HANDS OVER WHEN IT SAVES.
    //
    // *** THIS IS ALL THAT SURVIVED ComponentDraftWriter (Phase 6, 2026-09-23). *** That class
    // turned an editing session into BoardDraft row deltas; ComponentBoardWriter now applies the
    // same session straight to the draft's board, so the logic is gone and only the DTOs remain.
    //
    // They are kept - rather than folded into ComponentBoardWriter or replaced with BoardData's own
    // entry types - because they mirror the WINDOW's row models field for field
    // (ContributionComponentRow and friends, in ComponentContribution.axaml.cs), and that is their
    // job: they are the boundary between a UI collection the user edits and a board row. The two
    // differ in one way that matters, which is why BoardData's types cannot be used directly: an
    // attachment row carries FileLocation and File as SEPARATE fields, because the file picker
    // fills them separately, while a board row stores one joined path.
    //
    // The class name is kept as ComponentDraftWriter so the window's own Build*DraftRows methods
    // and their tests need no rename. It is a container for types now, not a writer - if that ever
    // reads as misleading, rename it and the ~30 references together rather than leaving the two
    // halves disagreeing.
    // ###########################################################################################
    public static class ComponentDraftWriter
    {
        // Mirrors ContributionComponentRow's editable fields.
        public sealed class ComponentDraftRow
        {
            public string BoardLabel { get; init; } = string.Empty;
            public string FriendlyName { get; init; } = string.Empty;
            public string TechnicalNameOrValue { get; init; } = string.Empty;
            public string PartNumber { get; init; } = string.Empty;
            public string Category { get; init; } = string.Empty;
            public string Region { get; init; } = string.Empty;
            public string Description { get; init; } = string.Empty;
        }

        // Mirrors ContributionComponentImageRow. FileLocation/File are the data-root-relative
        // folder and file name the attachment's bytes are copied to.
        public sealed class ComponentImageDraftRow
        {
            public string BoardLabel { get; init; } = string.Empty;
            public string Region { get; init; } = string.Empty;
            public string Pin { get; init; } = string.Empty;
            public string Name { get; init; } = string.Empty;
            public string ExpectedOscilloscopeReading { get; init; } = string.Empty;
            public string VoltsDiv { get; init; } = string.Empty;
            public string TimeDiv { get; init; } = string.Empty;
            public string TriggerLevelVolts { get; init; } = string.Empty;
            public string Note { get; init; } = string.Empty;
            public string FileLocation { get; init; } = string.Empty;
            public string File { get; init; } = string.Empty;
        }

        // Mirrors ContributionComponentLocalFileRow / ContributionBoardLocalFileRow - the same
        // shape serves both, distinguished by which list it is passed in as (component-scoped rows
        // carry BoardLabel, board-scoped rows carry Category instead).
        public sealed class LocalFileDraftRow
        {
            public string BoardLabel { get; init; } = string.Empty;
            public string Category { get; init; } = string.Empty;
            public string Name { get; init; } = string.Empty;
            public string FileLocation { get; init; } = string.Empty;
            public string File { get; init; } = string.Empty;
        }

        // Mirrors ContributionComponentLinkRow / ContributionBoardLinkRow.
        public sealed class LinkDraftRow
        {
            public string BoardLabel { get; init; } = string.Empty;
            public string Category { get; init; } = string.Empty;
            public string Name { get; init; } = string.Empty;
            public string Url { get; init; } = string.Empty;
        }
    }
}
