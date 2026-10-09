using System.Reflection;
using System.Text.Json;
using CRT.Server.Handlers.Submissions;
using Handlers.DataHandling;

namespace CRT.Server.Tests
{
    // ###########################################################################################
    // MIGRATION 0020 - "system" becomes "board" in the schema too (owner decision, 2026-10-09) - and in
    // what the database already STORES: a renamed key inside stored JSON would read as missing.
    //
    // submission_changes.changes_json holds CRT.Data's SubmissionChanges as MySqlSubmissionStore
    // writes it. Its IsNewSystem became IsNewBoard; a row written before the deploy still says
    // IsNewSystem, which the store now skips as unknown - every board created before it would lose
    // its "A new board." in the History view (code review, 2026-10-09). The migration rewrites the
    // key; these hold the migration's rewrite and the store's own writing to the same name.
    // ###########################################################################################
    public sealed class BoardNamesMigrationTests
    {
        private static readonly SubmissionChanges NewBoard = new(true, [], FileChanges.None);

        // The store's own settings, read off the store, so a change there is seen here.
        private static JsonSerializerOptions StoreOptions() =>
            (JsonSerializerOptions)typeof(MySqlSubmissionStore)
                .GetField("JsonOptions", BindingFlags.NonPublic | BindingFlags.Static)!
                .GetValue(null)!;

        private static string Migration() =>
            File.ReadAllText(Path.Combine(AppContext.BaseDirectory, "Migrations", "0020_board_names.sql"));

        [Fact]
        public void The_store_writes_the_new_board_flag_under_the_key_the_migration_rewrites_to()
        {
            string json = JsonSerializer.Serialize(BoardNamesMigrationTests.NewBoard, BoardNamesMigrationTests.StoreOptions());

            Assert.Contains("\"IsNewBoard\":true", json, StringComparison.Ordinal);
            Assert.Contains("REPLACE(changes_json, '\"IsNewSystem\"', '\"IsNewBoard\"')", BoardNamesMigrationTests.Migration(), StringComparison.Ordinal);
        }

        // A row from before the deploy: unread as it stands, read again once the migration's
        // rewrite has been applied to it.
        [Fact]
        public void A_row_stored_before_the_rename_reads_as_a_new_board_once_the_migration_rewrote_it()
        {
            string stored = JsonSerializer.Serialize(BoardNamesMigrationTests.NewBoard, BoardNamesMigrationTests.StoreOptions())
                .Replace("\"IsNewBoard\"", "\"IsNewSystem\"", StringComparison.Ordinal);

            Assert.False(JsonSerializer.Deserialize<SubmissionChanges>(stored, BoardNamesMigrationTests.StoreOptions())!.IsNewBoard);

            string rewritten = stored.Replace("\"IsNewSystem\"", "\"IsNewBoard\"", StringComparison.Ordinal);

            Assert.True(JsonSerializer.Deserialize<SubmissionChanges>(rewritten, BoardNamesMigrationTests.StoreOptions())!.IsNewBoard);
        }

        // ###########################################################################################
        // Nothing else the store keeps as JSON names a "system" - so the one rewrite above is all
        // stored JSON needs. The payload's rows and renames, an amendment's previous rows and files,
        // and the changes, each by its property names as the store writes them.
        // ###########################################################################################
        [Theory]
        [InlineData(typeof(SubmissionChanges))]
        [InlineData(typeof(SectionChanges))]
        [InlineData(typeof(FileChanges))]
        [InlineData(typeof(ChangedRowFact))]
        [InlineData(typeof(RenamedRowFact))]
        [InlineData(typeof(SubmissionRows))]
        [InlineData(typeof(SubmissionRename))]
        [InlineData(typeof(SubmissionFile))]
        public void No_other_type_kept_as_JSON_has_a_member_that_says_system(Type stored)
        {
            Assert.DoesNotContain(
                stored.GetProperties(BindingFlags.Public | BindingFlags.Instance),
                property => property.Name.Contains("System", StringComparison.OrdinalIgnoreCase));
        }
    }
}
