namespace Handlers.DataHandling
{
    // ###########################################################################################
    // system.json - RETIRED (owner decision, 2026-09-25).
    //
    // It was written into every published board's folder (NewContributeStrategy.md Phase 4, task
    // 7) and so downloaded by every user, while nothing in CRT showed anything from it. The
    // project owner: "I do not want this file visible in the source ... it should not be something
    // downloaded by all users, as this file is not relevant for them."
    //
    // Nothing writes or reads it any more. Every fact it carried is in the server's database: the
    // revision and content hash on `systems`, the maintainers in `maintainers`, the origin on
    // `systems.origin`. SystemDescriptor survives as the IN-MEMORY result of a publish - the
    // revision and content hash it records in the database - and is never serialised.
    //
    // Only the name is kept, so the server can remove a copy an earlier build left in a board
    // folder (RetiredSystemDescriptor) and a promotion never carries one to production.
    // ###########################################################################################
    public static class SystemDescriptorStore
    {
        public const string FileName = "system.json";
    }
}
