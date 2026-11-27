using System.Text.Json;
using CRT.Server.Configuration;
using Handlers.DataHandling;
using MySqlConnector;

namespace CRT.Server.Handlers.Collections
{
    // ###########################################################################################
    // ICollectionStore on MySQL (migration 0025). An I/O boundary like the other MySql* stores - the rules are
    // CollectionFlows', tested against a fake.
    // ###########################################################################################
    public sealed class MySqlCollectionStore : ICollectionStore
    {
        private const string CollectionColumns =
            "c.id, c.account_id, c.hardware_name, c.board_name, c.region, c.title, c.description, c.forum_links_json, " +
            "c.created_utc, c.updated_utc, c.delete_at_utc, c.reminded_utc, c.deleted_utc, c.deleted_reason, c.unshared_utc, c.components_json";

        private const string ImageColumns =
            "position, sha256, width, height, component_label, pin, signal_name, expected_reading, time_per_division, " +
            "volts_per_division, trigger_level, caption, captured_utc, pinout_sha256";

        private readonly string thisConnectionString;

        public MySqlCollectionStore(ServerOptions options)
        {
            ArgumentNullException.ThrowIfNull(options);
            ArgumentException.ThrowIfNullOrWhiteSpace(options.ConnectionString);

            this.thisConnectionString = options.ConnectionString;
        }

        public Task<CollectionRecord?> FindAsync(string id, CancellationToken cancellationToken = default) =>
            this.FindCoreAsync(id, sha256: null, cancellationToken);

        public Task<CollectionRecord?> FindWithImageAsync(string id, string sha256, CancellationToken cancellationToken = default) =>
            this.FindCoreAsync(id, sha256 ?? throw new ArgumentNullException(nameof(sha256)), cancellationToken);

        // `sha256` null reads every capture; set, only that capture's rows.
        private async Task<CollectionRecord?> FindCoreAsync(string id, string? sha256, CancellationToken cancellationToken)
        {
            await using MySqlConnection connection = await this.OpenAsync(cancellationToken);

            CollectionRecord? collection;

            await using (MySqlCommand select = connection.CreateCommand())
            {
                select.CommandText = $"SELECT {CollectionColumns}, 0 FROM collections c WHERE c.id = @id;";
                select.Parameters.AddWithValue("@id", id);

                await using MySqlDataReader reader = await select.ExecuteReaderAsync(cancellationToken);
                collection = await reader.ReadAsync(cancellationToken) ? MySqlCollectionStore.ReadCollection(reader) : null;
            }

            if (collection is null)
                return null;

            var images = new List<CollectionImageRecord>();

            await using (MySqlCommand select = connection.CreateCommand())
            {
                // A picture is a capture or a capture's pinout drawing (migration 0029).
                select.CommandText =
                    $"SELECT {ImageColumns} FROM collection_images WHERE collection_id = @id" +
                    (sha256 is null ? string.Empty : " AND (sha256 = @sha OR pinout_sha256 = @sha)") + " ORDER BY position;";
                select.Parameters.AddWithValue("@id", id);

                if (sha256 is not null)
                    select.Parameters.AddWithValue("@sha", sha256);

                await using MySqlDataReader reader = await select.ExecuteReaderAsync(cancellationToken);

                while (await reader.ReadAsync(cancellationToken))
                    images.Add(MySqlCollectionStore.ReadImage(reader));
            }

            return collection with { Images = images, ImageCount = images.Count };
        }

        // CollectionColumns, the captures counted, and the comments the page shows counted - none hidden, none of a locked
        // account (AccountLockFlows), none on a capture the collection no longer shows (OnACaptureShown) - with the newest one's
        // id of OTHERS than the owner: what marks "new comments" in CRT, and the owner's own answer on the website is nothing
        // new to them (code review, 2026-11-27).
        private const string CountedColumns =
            $"{CollectionColumns}, (SELECT COUNT(*) FROM collection_images i WHERE i.collection_id = c.id), " +
            "(SELECT COUNT(*) FROM collection_comments m JOIN accounts w ON w.id = m.account_id " +
            $"WHERE m.collection_id = c.id AND m.deleted_utc IS NULL AND m.hidden_utc IS NULL AND w.is_locked = 0 AND {OnACaptureShown}), " +
            "(SELECT COALESCE(MAX(m.id), 0) FROM collection_comments m JOIN accounts w ON w.id = m.account_id " +
            $"WHERE m.collection_id = c.id AND m.deleted_utc IS NULL AND m.hidden_utc IS NULL AND w.is_locked = 0 AND m.account_id <> c.account_id AND {OnACaptureShown})";

        // ###########################################################################################
        // *** A COMMENT COUNTS ONLY ON A CAPTURE ITS COLLECTION SHOWS NOW (code review, 2026-11-27). *** A capture taken out and
        // shared again leaves its comments - nothing deletes them, so the capture shared again brings them back - but no page
        // shows them and no CRT window lists them: counted, a collection read "1 comment" and "new comments" for one nobody
        // could see, hide or delete, and the administrator's newest comments linked to a place no page has.
        // ###########################################################################################
        private const string OnACaptureShown =
            "EXISTS (SELECT 1 FROM collection_images s WHERE s.collection_id = m.collection_id AND s.sha256 = m.image_sha256)";

        public Task<IReadOnlyList<CollectionRecord>> ListLiveForAccountAsync(long accountId, CancellationToken cancellationToken = default) =>
            this.ListCountedAsync(
                $"SELECT {CountedColumns} FROM collections c WHERE c.account_id = @account AND c.deleted_utc IS NULL ORDER BY c.created_utc DESC;",
                cancellationToken,
                ("@account", accountId));

        public Task<IReadOnlyList<CollectionRecord>> ListLiveCommentedOnAsync(long accountId, CancellationToken cancellationToken = default) =>
            this.ListCountedAsync(
                $"SELECT {CountedColumns} FROM collections c " +
                "JOIN (SELECT m.collection_id, MAX(m.id) AS latest FROM collection_comments m WHERE m.account_id = @account AND m.deleted_utc IS NULL " +
                $"AND {OnACaptureShown} GROUP BY m.collection_id) w " +
                "ON w.collection_id = c.id " +
                "WHERE c.account_id <> @account AND c.deleted_utc IS NULL ORDER BY w.latest DESC;",
                cancellationToken,
                ("@account", accountId));

        public Task<IReadOnlyList<CollectionRecord>> ListAllLiveAsync(int count, CancellationToken cancellationToken = default) =>
            this.ListCountedAsync(
                $"SELECT {CountedColumns} FROM collections c WHERE c.deleted_utc IS NULL ORDER BY c.created_utc DESC LIMIT @count;",
                cancellationToken,
                ("@count", count));

        private async Task<IReadOnlyList<CollectionRecord>> ListCountedAsync(string sql, CancellationToken cancellationToken, params (string Name, object Value)[] parameters)
        {
            await using MySqlConnection connection = await this.OpenAsync(cancellationToken);
            await using MySqlCommand select = connection.CreateCommand();

            select.CommandText = sql;

            foreach ((string name, object value) in parameters)
                select.Parameters.AddWithValue(name, value);

            var collections = new List<CollectionRecord>();

            await using MySqlDataReader reader = await select.ExecuteReaderAsync(cancellationToken);

            while (await reader.ReadAsync(cancellationToken))
            {
                collections.Add(MySqlCollectionStore.ReadCollection(reader) with
                {
                    CommentCount = Convert.ToInt32(reader.GetValue(17)),
                    LatestCommentId = Convert.ToInt64(reader.GetValue(18))
                });
            }

            return collections;
        }

        public async Task<int> CountCreatedSinceAsync(long accountId, DateTimeOffset sinceUtc, CancellationToken cancellationToken = default)
        {
            await using MySqlConnection connection = await this.OpenAsync(cancellationToken);
            await using MySqlCommand count = connection.CreateCommand();

            count.CommandText = "SELECT COUNT(*) FROM collections WHERE account_id = @account AND created_utc > @since;";
            count.Parameters.AddWithValue("@account", accountId);
            count.Parameters.AddWithValue("@since", sinceUtc.UtcDateTime);

            return Convert.ToInt32(await count.ExecuteScalarAsync(cancellationToken));
        }

        public async Task<bool> AddAsync(CollectionRecord collection, CancellationToken cancellationToken = default)
        {
            ArgumentNullException.ThrowIfNull(collection);

            await using MySqlConnection connection = await this.OpenAsync(cancellationToken);
            await using MySqlTransaction transaction = await connection.BeginTransactionAsync(cancellationToken);

            await using (MySqlCommand insert = connection.CreateCommand())
            {
                insert.Transaction = transaction;
                insert.CommandText =
                    "INSERT IGNORE INTO collections (id, account_id, client_key, hardware_name, board_name, region, title, description, " +
                    "forum_links_json, components_json, created_utc, updated_utc, delete_at_utc, reminded_utc) VALUES " +
                    "(@id, @account, @clientKey, @hardware, @board, @region, @title, @description, @links, @components, @created, @updated, @deleteAt, @reminded);";
                MySqlCollectionStore.AddContent(insert, collection);
                insert.Parameters.AddWithValue("@account", collection.AccountId);
                insert.Parameters.AddWithValue("@clientKey", (object?)collection.ClientKey ?? DBNull.Value);
                insert.Parameters.AddWithValue("@created", collection.CreatedUtc.UtcDateTime);

                // The id taken already: nothing written, and CollectionFlows draws another.
                if (await insert.ExecuteNonQueryAsync(cancellationToken) == 0)
                    return false;
            }

            await MySqlCollectionStore.InsertImagesAsync(connection, transaction, collection, cancellationToken);
            await transaction.CommitAsync(cancellationToken);

            return true;
        }

        public async Task<CollectionRecord?> FindLiveByClientKeyAsync(long accountId, string clientKey, CancellationToken cancellationToken = default)
        {
            string? id;

            await using (MySqlConnection connection = await this.OpenAsync(cancellationToken))
            await using (MySqlCommand command = connection.CreateCommand())
            {
                command.CommandText =
                    "SELECT id FROM collections WHERE account_id = @account AND client_key = @clientKey AND deleted_utc IS NULL " +
                    "ORDER BY created_utc DESC LIMIT 1;";
                command.Parameters.AddWithValue("@account", accountId);
                command.Parameters.AddWithValue("@clientKey", clientKey);

                id = await command.ExecuteScalarAsync(cancellationToken) as string;
            }

            return id is null ? null : await this.FindAsync(id, cancellationToken);
        }

        public async Task<bool> ReplaceContentAsync(CollectionRecord collection, CancellationToken cancellationToken = default)
        {
            ArgumentNullException.ThrowIfNull(collection);

            await using MySqlConnection connection = await this.OpenAsync(cancellationToken);
            await using MySqlTransaction transaction = await connection.BeginTransactionAsync(cancellationToken);

            await using (MySqlCommand update = connection.CreateCommand())
            {
                update.Transaction = transaction;
                update.CommandText =
                    "UPDATE collections SET hardware_name = @hardware, board_name = @board, region = @region, title = @title, " +
                    "description = @description, forum_links_json = @links, components_json = @components, updated_utc = @updated, " +
                    // A new date takes what the flow decided of its reminder (CollectionFlows.ReminderSettledAt); the same date keeps
                    // what was decided - the sweep's mark too, set meanwhile. Before delete_at_utc's own assignment, which MariaDB
                    // makes in order.
                    "reminded_utc = CASE WHEN @deleteAt <> delete_at_utc THEN @reminded ELSE reminded_utc END, delete_at_utc = @deleteAt, " +
                    // Shared again: "Share" brings a collection whose sharing was stopped back under its link (migration 0028).
                    "unshared_utc = NULL " +
                    "WHERE id = @id AND deleted_utc IS NULL;";
                MySqlCollectionStore.AddContent(update, collection);

                if (await update.ExecuteNonQueryAsync(cancellationToken) == 0)
                    return false;
            }

            await using (MySqlCommand delete = connection.CreateCommand())
            {
                delete.Transaction = transaction;
                delete.CommandText = "DELETE FROM collection_images WHERE collection_id = @id;";
                delete.Parameters.AddWithValue("@id", collection.Id);
                await delete.ExecuteNonQueryAsync(cancellationToken);
            }

            await MySqlCollectionStore.InsertImagesAsync(connection, transaction, collection, cancellationToken);
            await transaction.CommitAsync(cancellationToken);

            return true;
        }

        public async Task<bool> SetUnsharedAsync(string id, DateTimeOffset? unsharedUtc, CancellationToken cancellationToken = default)
        {
            await using MySqlConnection connection = await this.OpenAsync(cancellationToken);
            await using MySqlCommand update = connection.CreateCommand();

            // Stopped again keeps the first stop's time; the row counts as changed by its id alone (ROW_COUNT would read 0 for a
            // value set to what it was, which is no refusal).
            update.CommandText =
                "UPDATE collections SET unshared_utc = CASE WHEN @when IS NULL THEN NULL ELSE COALESCE(unshared_utc, @when) END " +
                "WHERE id = @id AND deleted_utc IS NULL;" +
                "SELECT COUNT(*) FROM collections WHERE id = @id AND deleted_utc IS NULL;";
            update.Parameters.AddWithValue("@id", id);
            update.Parameters.AddWithValue("@when", unsharedUtc is { } when ? when.UtcDateTime : DBNull.Value);

            return Convert.ToInt32(await update.ExecuteScalarAsync(cancellationToken)) > 0;
        }

        public async Task<bool> MarkDeletedAsync(string id, string reason, DateTimeOffset whenUtc, CancellationToken cancellationToken = default)
        {
            await using MySqlConnection connection = await this.OpenAsync(cancellationToken);
            await using MySqlTransaction transaction = await connection.BeginTransactionAsync(cancellationToken);

            await using (MySqlCommand update = connection.CreateCommand())
            {
                update.Transaction = transaction;
                update.CommandText =
                    "UPDATE collections SET deleted_utc = @when, deleted_reason = @reason, title = NULL, description = NULL, " +
                    "forum_links_json = NULL, components_json = NULL WHERE id = @id AND deleted_utc IS NULL;";
                update.Parameters.AddWithValue("@id", id);
                update.Parameters.AddWithValue("@when", whenUtc.UtcDateTime);
                update.Parameters.AddWithValue("@reason", reason);

                if (await update.ExecuteNonQueryAsync(cancellationToken) == 0)
                    return false;
            }

            // Its captures and their comments go with it.
            await using (MySqlCommand delete = connection.CreateCommand())
            {
                delete.Transaction = transaction;
                delete.CommandText = "DELETE FROM collection_images WHERE collection_id = @id; DELETE FROM collection_comments WHERE collection_id = @id;";
                delete.Parameters.AddWithValue("@id", id);
                await delete.ExecuteNonQueryAsync(cancellationToken);
            }

            await transaction.CommitAsync(cancellationToken);

            return true;
        }

        // ---------------------------------------------------------------------------------------
        // Comments (migration 0026).
        // ---------------------------------------------------------------------------------------

        private const string CommentColumns =
            "m.id, m.collection_id, m.image_sha256, m.account_id, a.display_name, a.is_administrator, m.body, m.created_utc, m.edited_utc, m.hidden_utc, a.is_locked";

        public async Task<IReadOnlyList<CollectionCommentRecord>> ListCommentsAsync(string collectionId, CancellationToken cancellationToken = default)
        {
            await using MySqlConnection connection = await this.OpenAsync(cancellationToken);
            await using MySqlCommand select = connection.CreateCommand();

            select.CommandText =
                $"SELECT {CommentColumns} FROM collection_comments m JOIN accounts a ON a.id = m.account_id " +
                "WHERE m.collection_id = @id AND m.deleted_utc IS NULL ORDER BY m.created_utc, m.id;";
            select.Parameters.AddWithValue("@id", collectionId);

            var comments = new List<CollectionCommentRecord>();

            await using MySqlDataReader reader = await select.ExecuteReaderAsync(cancellationToken);

            while (await reader.ReadAsync(cancellationToken))
                comments.Add(MySqlCollectionStore.ReadComment(reader));

            return comments;
        }

        public async Task<CollectionCommentRecord?> FindCommentAsync(long commentId, CancellationToken cancellationToken = default)
        {
            await using MySqlConnection connection = await this.OpenAsync(cancellationToken);
            await using MySqlCommand select = connection.CreateCommand();

            select.CommandText =
                $"SELECT {CommentColumns} FROM collection_comments m JOIN accounts a ON a.id = m.account_id " +
                "WHERE m.id = @id AND m.deleted_utc IS NULL;";
            select.Parameters.AddWithValue("@id", commentId);

            await using MySqlDataReader reader = await select.ExecuteReaderAsync(cancellationToken);

            return await reader.ReadAsync(cancellationToken) ? MySqlCollectionStore.ReadComment(reader) : null;
        }

        public async Task<long> AddCommentAsync(string collectionId, string imageSha256, long accountId, string body, DateTimeOffset whenUtc, CancellationToken cancellationToken = default)
        {
            await using MySqlConnection connection = await this.OpenAsync(cancellationToken);
            await using MySqlCommand insert = connection.CreateCommand();

            insert.CommandText =
                "INSERT INTO collection_comments (collection_id, image_sha256, account_id, body, created_utc) " +
                "VALUES (@collection, @sha, @account, @body, @when); SELECT LAST_INSERT_ID();";
            insert.Parameters.AddWithValue("@collection", collectionId);
            insert.Parameters.AddWithValue("@sha", imageSha256);
            insert.Parameters.AddWithValue("@account", accountId);
            insert.Parameters.AddWithValue("@body", body);
            insert.Parameters.AddWithValue("@when", whenUtc.UtcDateTime);

            return Convert.ToInt64(await insert.ExecuteScalarAsync(cancellationToken));
        }

        public Task<bool> EditCommentAsync(long commentId, string body, DateTimeOffset whenUtc, CancellationToken cancellationToken = default) =>
            this.ChangeCommentAsync(
                "UPDATE collection_comments SET body = @body, edited_utc = @when WHERE id = @id AND deleted_utc IS NULL;",
                commentId,
                whenUtc,
                cancellationToken,
                ("@body", body));

        public Task<bool> DeleteCommentAsync(long commentId, DateTimeOffset whenUtc, CancellationToken cancellationToken = default) =>
            this.ChangeCommentAsync(
                "UPDATE collection_comments SET deleted_utc = @when WHERE id = @id AND deleted_utc IS NULL;",
                commentId,
                whenUtc,
                cancellationToken);

        public Task<bool> SetCommentHiddenAsync(long commentId, bool hidden, DateTimeOffset whenUtc, CancellationToken cancellationToken = default) =>
            this.ChangeCommentAsync(
                hidden
                    ? "UPDATE collection_comments SET hidden_utc = @when WHERE id = @id AND deleted_utc IS NULL;"
                    : "UPDATE collection_comments SET hidden_utc = NULL WHERE id = @id AND deleted_utc IS NULL;",
                commentId,
                whenUtc,
                cancellationToken);

        public async Task<IReadOnlyList<CollectionCommentRecord>> ListNewestCommentsAsync(int count, CancellationToken cancellationToken = default)
        {
            await using MySqlConnection connection = await this.OpenAsync(cancellationToken);
            await using MySqlCommand select = connection.CreateCommand();

            select.CommandText =
                $"SELECT {CommentColumns} FROM collection_comments m JOIN accounts a ON a.id = m.account_id " +
                "JOIN collections c ON c.id = m.collection_id " +
                // A collection whose time is up is gone already, the sweep or not (CollectionFlows.IsGone).
                $"WHERE m.deleted_utc IS NULL AND c.deleted_utc IS NULL AND c.delete_at_utc > UTC_TIMESTAMP(3) AND {OnACaptureShown} " +
                "ORDER BY m.id DESC LIMIT @count;";
            select.Parameters.AddWithValue("@count", count);

            var comments = new List<CollectionCommentRecord>();

            await using MySqlDataReader reader = await select.ExecuteReaderAsync(cancellationToken);

            while (await reader.ReadAsync(cancellationToken))
                comments.Add(MySqlCollectionStore.ReadComment(reader));

            return comments;
        }

        public async Task<int> CountCommentsSinceAsync(long accountId, DateTimeOffset sinceUtc, CancellationToken cancellationToken = default)
        {
            await using MySqlConnection connection = await this.OpenAsync(cancellationToken);
            await using MySqlCommand count = connection.CreateCommand();

            count.CommandText = "SELECT COUNT(*) FROM collection_comments WHERE account_id = @account AND created_utc > @since;";
            count.Parameters.AddWithValue("@account", accountId);
            count.Parameters.AddWithValue("@since", sinceUtc.UtcDateTime);

            return Convert.ToInt32(await count.ExecuteScalarAsync(cancellationToken));
        }

        private async Task<bool> ChangeCommentAsync(string sql, long commentId, DateTimeOffset whenUtc, CancellationToken cancellationToken, params (string Name, object Value)[] more)
        {
            await using MySqlConnection connection = await this.OpenAsync(cancellationToken);
            await using MySqlCommand update = connection.CreateCommand();

            update.CommandText = sql;
            update.Parameters.AddWithValue("@id", commentId);
            update.Parameters.AddWithValue("@when", whenUtc.UtcDateTime);

            foreach ((string name, object value) in more)
                update.Parameters.AddWithValue(name, value);

            return await update.ExecuteNonQueryAsync(cancellationToken) == 1;
        }

        private static CollectionCommentRecord ReadComment(MySqlDataReader reader) =>
            new(
                Convert.ToInt64(reader.GetValue(0)),
                reader.GetString(1),
                reader.GetString(2),
                Convert.ToInt64(reader.GetValue(3)),
                reader.GetString(4),
                reader.GetBoolean(5),
                reader.GetString(6),
                MySqlCollectionStore.Utc(reader.GetDateTime(7)),
                reader.IsDBNull(8) ? null : MySqlCollectionStore.Utc(reader.GetDateTime(8)),
                reader.IsDBNull(9) ? null : MySqlCollectionStore.Utc(reader.GetDateTime(9)),
                reader.GetBoolean(10));

        public async Task<IReadOnlyList<string>> ListExpiredAsync(DateTimeOffset nowUtc, CancellationToken cancellationToken = default)
        {
            await using MySqlConnection connection = await this.OpenAsync(cancellationToken);
            await using MySqlCommand select = connection.CreateCommand();

            select.CommandText = "SELECT id FROM collections WHERE deleted_utc IS NULL AND delete_at_utc <= @now;";
            select.Parameters.AddWithValue("@now", nowUtc.UtcDateTime);

            var ids = new List<string>();

            await using MySqlDataReader reader = await select.ExecuteReaderAsync(cancellationToken);

            while (await reader.ReadAsync(cancellationToken))
                ids.Add(reader.GetString(0));

            return ids;
        }

        public async Task<IReadOnlyList<CollectionRecord>> ListToRemindAsync(DateTimeOffset byUtc, CancellationToken cancellationToken = default)
        {
            await using MySqlConnection connection = await this.OpenAsync(cancellationToken);
            await using MySqlCommand select = connection.CreateCommand();

            select.CommandText =
                $"SELECT {CollectionColumns}, 0 FROM collections c " +
                "WHERE c.deleted_utc IS NULL AND c.reminded_utc IS NULL AND c.delete_at_utc <= @by;";
            select.Parameters.AddWithValue("@by", byUtc.UtcDateTime);

            return await MySqlCollectionStore.ReadCollectionsAsync(select, cancellationToken);
        }

        public async Task MarkRemindedAsync(string id, DateTimeOffset whenUtc, CancellationToken cancellationToken = default)
        {
            await using MySqlConnection connection = await this.OpenAsync(cancellationToken);
            await using MySqlCommand update = connection.CreateCommand();

            update.CommandText = "UPDATE collections SET reminded_utc = @when WHERE id = @id;";
            update.Parameters.AddWithValue("@id", id);
            update.Parameters.AddWithValue("@when", whenUtc.UtcDateTime);
            await update.ExecuteNonQueryAsync(cancellationToken);
        }

        public async Task<IReadOnlyDictionary<string, CollectionUploadRecord>> FindUploadsAsync(
            long accountId,
            IReadOnlyCollection<string> sha256s,
            CancellationToken cancellationToken = default)
        {
            var found = new Dictionary<string, CollectionUploadRecord>(StringComparer.Ordinal);

            if (sha256s.Count == 0)
                return found;

            await using MySqlConnection connection = await this.OpenAsync(cancellationToken);
            await using MySqlCommand select = connection.CreateCommand();

            string[] names = sha256s.Select((_, index) => $"@s{index}").ToArray();

            select.CommandText =
                "SELECT account_id, sha256, byte_count, width, height, uploaded_utc FROM collection_uploads " +
                $"WHERE account_id = @account AND sha256 IN ({string.Join(", ", names)});";
            select.Parameters.AddWithValue("@account", accountId);

            int at = 0;

            foreach (string sha in sha256s)
                select.Parameters.AddWithValue(names[at++], sha);

            await using MySqlDataReader reader = await select.ExecuteReaderAsync(cancellationToken);

            while (await reader.ReadAsync(cancellationToken))
            {
                var upload = new CollectionUploadRecord(
                    Convert.ToInt64(reader.GetValue(0)),
                    reader.GetString(1),
                    reader.GetInt32(2),
                    reader.GetInt32(3),
                    reader.GetInt32(4),
                    MySqlCollectionStore.Utc(reader.GetDateTime(5)));

                found[upload.Sha256] = upload;
            }

            return found;
        }

        public async Task AddUploadAsync(CollectionUploadRecord upload, CancellationToken cancellationToken = default)
        {
            ArgumentNullException.ThrowIfNull(upload);

            await using MySqlConnection connection = await this.OpenAsync(cancellationToken);
            await using MySqlCommand insert = connection.CreateCommand();

            insert.CommandText =
                "INSERT INTO collection_uploads (account_id, sha256, byte_count, width, height, uploaded_utc) " +
                "VALUES (@account, @sha, @bytes, @width, @height, @when) " +
                "ON DUPLICATE KEY UPDATE uploaded_utc = GREATEST(uploaded_utc, VALUES(uploaded_utc));";
            insert.Parameters.AddWithValue("@account", upload.AccountId);
            insert.Parameters.AddWithValue("@sha", upload.Sha256);
            insert.Parameters.AddWithValue("@bytes", upload.ByteCount);
            insert.Parameters.AddWithValue("@width", upload.Width);
            insert.Parameters.AddWithValue("@height", upload.Height);
            insert.Parameters.AddWithValue("@when", upload.UploadedUtc.UtcDateTime);

            await insert.ExecuteNonQueryAsync(cancellationToken);
        }

        public async Task<int> CountUploadsSinceAsync(long accountId, DateTimeOffset sinceUtc, CancellationToken cancellationToken = default)
        {
            await using MySqlConnection connection = await this.OpenAsync(cancellationToken);
            await using MySqlCommand count = connection.CreateCommand();

            count.CommandText = "SELECT COUNT(*) FROM collection_uploads WHERE account_id = @account AND uploaded_utc > @since;";
            count.Parameters.AddWithValue("@account", accountId);
            count.Parameters.AddWithValue("@since", sinceUtc.UtcDateTime);

            return Convert.ToInt32(await count.ExecuteScalarAsync(cancellationToken));
        }

        public async Task<long> StoredBytesAsync(CancellationToken cancellationToken = default)
        {
            await using MySqlConnection connection = await this.OpenAsync(cancellationToken);
            await using MySqlCommand sum = connection.CreateCommand();

            sum.CommandText =
                "SELECT COALESCE(SUM(bytes), 0) FROM (SELECT MAX(byte_count) AS bytes FROM collection_uploads GROUP BY sha256) AS pictures;";

            return Convert.ToInt64(await sum.ExecuteScalarAsync(cancellationToken));
        }

        public async Task<bool> IsUploadedAsync(string sha256, CancellationToken cancellationToken = default)
        {
            await using MySqlConnection connection = await this.OpenAsync(cancellationToken);
            await using MySqlCommand command = connection.CreateCommand();

            command.CommandText = "SELECT EXISTS (SELECT 1 FROM collection_uploads WHERE sha256 = @sha);";
            command.Parameters.AddWithValue("@sha", sha256);

            return Convert.ToInt32(await command.ExecuteScalarAsync(cancellationToken)) != 0;
        }

        public async Task<IReadOnlyList<string>> DeleteUnusedUploadsAsync(DateTimeOffset uploadedBeforeUtc, CancellationToken cancellationToken = default)
        {
            await using MySqlConnection connection = await this.OpenAsync(cancellationToken);
            await using MySqlTransaction transaction = await connection.BeginTransactionAsync(cancellationToken);

            var unused = new List<(long Account, string Sha)>();

            await using (MySqlCommand select = connection.CreateCommand())
            {
                select.Transaction = transaction;
                select.CommandText =
                    "SELECT u.account_id, u.sha256 FROM collection_uploads u WHERE u.uploaded_utc < @before AND NOT EXISTS (" +
                    "SELECT 1 FROM collection_images i JOIN collections c ON c.id = i.collection_id " +
                    // A capture's pinout drawing is shown too (migration 0029).
                    "WHERE c.account_id = u.account_id AND c.deleted_utc IS NULL AND (i.sha256 = u.sha256 OR i.pinout_sha256 = u.sha256)) FOR UPDATE;";
                select.Parameters.AddWithValue("@before", uploadedBeforeUtc.UtcDateTime);

                await using MySqlDataReader reader = await select.ExecuteReaderAsync(cancellationToken);

                while (await reader.ReadAsync(cancellationToken))
                    unused.Add((Convert.ToInt64(reader.GetValue(0)), reader.GetString(1)));
            }

            foreach ((long account, string sha) in unused)
            {
                await using MySqlCommand delete = connection.CreateCommand();

                delete.Transaction = transaction;
                delete.CommandText = "DELETE FROM collection_uploads WHERE account_id = @account AND sha256 = @sha;";
                delete.Parameters.AddWithValue("@account", account);
                delete.Parameters.AddWithValue("@sha", sha);
                await delete.ExecuteNonQueryAsync(cancellationToken);
            }

            var gone = new List<string>();

            foreach (string sha in unused.Select(row => row.Sha).Distinct(StringComparer.Ordinal))
            {
                await using MySqlCommand left = connection.CreateCommand();

                left.Transaction = transaction;
                left.CommandText = "SELECT COUNT(*) FROM collection_uploads WHERE sha256 = @sha;";
                left.Parameters.AddWithValue("@sha", sha);

                if (Convert.ToInt32(await left.ExecuteScalarAsync(cancellationToken)) == 0)
                    gone.Add(sha);
            }

            await transaction.CommitAsync(cancellationToken);

            return gone;
        }

        // ---------------------------------------------------------------------------------------

        private static void AddContent(MySqlCommand command, CollectionRecord collection)
        {
            command.Parameters.AddWithValue("@id", collection.Id);
            command.Parameters.AddWithValue("@hardware", collection.HardwareName);
            command.Parameters.AddWithValue("@board", collection.BoardName);
            command.Parameters.AddWithValue("@region", (object?)collection.Region ?? DBNull.Value);
            command.Parameters.AddWithValue("@title", (object?)collection.Title ?? DBNull.Value);
            command.Parameters.AddWithValue("@description", (object?)collection.Description ?? DBNull.Value);
            command.Parameters.AddWithValue("@links", collection.ForumLinks.Count == 0 ? DBNull.Value : JsonSerializer.Serialize(collection.ForumLinks));
            command.Parameters.AddWithValue("@components", collection.Components is not { Count: > 0 } components ? DBNull.Value : JsonSerializer.Serialize(components));
            command.Parameters.AddWithValue("@updated", collection.UpdatedUtc.UtcDateTime);
            command.Parameters.AddWithValue("@deleteAt", collection.DeleteAtUtc.UtcDateTime);
            command.Parameters.AddWithValue("@reminded", collection.RemindedUtc is { } reminded ? reminded.UtcDateTime : DBNull.Value);
        }

        private static async Task InsertImagesAsync(MySqlConnection connection, MySqlTransaction transaction, CollectionRecord collection, CancellationToken cancellationToken)
        {
            for (int position = 0; position < collection.Images.Count; position++)
            {
                CollectionImageRecord image = collection.Images[position];
                CollectionImageEntry entry = image.Entry;

                await using MySqlCommand insert = connection.CreateCommand();

                insert.Transaction = transaction;
                insert.CommandText =
                    $"INSERT INTO collection_images (collection_id, {ImageColumns}) VALUES " +
                    "(@id, @position, @sha, @width, @height, @label, @pin, @signal, @reading, @tdiv, @vdiv, @trigger, @caption, @captured, @pinout);";
                insert.Parameters.AddWithValue("@id", collection.Id);
                insert.Parameters.AddWithValue("@position", position);
                insert.Parameters.AddWithValue("@sha", entry.Sha256);
                insert.Parameters.AddWithValue("@width", image.Width);
                insert.Parameters.AddWithValue("@height", image.Height);
                insert.Parameters.AddWithValue("@label", entry.ComponentLabel);
                insert.Parameters.AddWithValue("@pin", (object?)entry.Pin ?? DBNull.Value);
                insert.Parameters.AddWithValue("@signal", (object?)entry.SignalName ?? DBNull.Value);
                insert.Parameters.AddWithValue("@reading", (object?)entry.ExpectedReading ?? DBNull.Value);
                insert.Parameters.AddWithValue("@tdiv", (object?)entry.TimePerDivision ?? DBNull.Value);
                insert.Parameters.AddWithValue("@vdiv", (object?)entry.VoltsPerDivision ?? DBNull.Value);
                insert.Parameters.AddWithValue("@trigger", (object?)entry.TriggerLevel ?? DBNull.Value);
                insert.Parameters.AddWithValue("@caption", (object?)entry.Caption ?? DBNull.Value);
                insert.Parameters.AddWithValue("@captured", (object?)entry.CapturedUtc?.UtcDateTime ?? DBNull.Value);
                insert.Parameters.AddWithValue("@pinout", (object?)entry.PinoutSha256 ?? DBNull.Value);

                await insert.ExecuteNonQueryAsync(cancellationToken);
            }
        }

        private static async Task<IReadOnlyList<CollectionRecord>> ReadCollectionsAsync(MySqlCommand select, CancellationToken cancellationToken)
        {
            var collections = new List<CollectionRecord>();

            await using MySqlDataReader reader = await select.ExecuteReaderAsync(cancellationToken);

            while (await reader.ReadAsync(cancellationToken))
                collections.Add(MySqlCollectionStore.ReadCollection(reader));

            return collections;
        }

        // CollectionColumns, then the image count.
        private static CollectionRecord ReadCollection(MySqlDataReader reader)
        {
            string? links = MySqlCollectionStore.Text(reader, 7);

            return new CollectionRecord(
                reader.GetString(0),
                Convert.ToInt64(reader.GetValue(1)),
                reader.GetString(2),
                reader.GetString(3),
                MySqlCollectionStore.Text(reader, 4),
                MySqlCollectionStore.Text(reader, 5),
                MySqlCollectionStore.Text(reader, 6),
                links is null ? [] : JsonSerializer.Deserialize<List<string>>(links) ?? [],
                MySqlCollectionStore.Utc(reader.GetDateTime(8)),
                MySqlCollectionStore.Utc(reader.GetDateTime(9)),
                MySqlCollectionStore.Utc(reader.GetDateTime(10)),
                reader.IsDBNull(11) ? null : MySqlCollectionStore.Utc(reader.GetDateTime(11)),
                reader.IsDBNull(12) ? null : MySqlCollectionStore.Utc(reader.GetDateTime(12)),
                MySqlCollectionStore.Text(reader, 13),
                [],
                Convert.ToInt32(reader.GetValue(16)),
                UnsharedUtc: reader.IsDBNull(14) ? null : MySqlCollectionStore.Utc(reader.GetDateTime(14)),
                Components: MySqlCollectionStore.Text(reader, 15) is { } components
                    ? JsonSerializer.Deserialize<List<CollectionComponent>>(components) ?? []
                    : []);
        }

        // ImageColumns, in order.
        private static CollectionImageRecord ReadImage(MySqlDataReader reader)
        {
            var entry = new CollectionImageEntry(
                reader.GetString(1),
                reader.GetString(4),
                MySqlCollectionStore.Text(reader, 5),
                MySqlCollectionStore.Text(reader, 6),
                MySqlCollectionStore.Text(reader, 7),
                MySqlCollectionStore.Text(reader, 8),
                MySqlCollectionStore.Text(reader, 9),
                MySqlCollectionStore.Text(reader, 10),
                MySqlCollectionStore.Text(reader, 11),
                reader.IsDBNull(12) ? null : MySqlCollectionStore.Utc(reader.GetDateTime(12)),
                MySqlCollectionStore.Text(reader, 13));

            return new CollectionImageRecord(entry, reader.GetInt32(2), reader.GetInt32(3));
        }

        private static string? Text(MySqlDataReader reader, int ordinal) => reader.IsDBNull(ordinal) ? null : reader.GetString(ordinal);

        private static DateTimeOffset Utc(DateTime value) => new(DateTime.SpecifyKind(value, DateTimeKind.Utc));

        private async Task<MySqlConnection> OpenAsync(CancellationToken cancellationToken)
        {
            var connection = new MySqlConnection(this.thisConnectionString);
            await connection.OpenAsync(cancellationToken);

            return connection;
        }
    }
}
