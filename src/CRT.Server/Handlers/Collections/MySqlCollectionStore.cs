using System.Text.Json;
using CRT.Server.Configuration;
using Handlers.DataHandling;
using MySqlConnector;

namespace CRT.Server.Handlers.Collections
{
    // ###########################################################################################
    // ICollectionStore on MySQL (migrations 0025-0031). An I/O boundary like the other MySql* stores - the rules are
    // CollectionFlows', tested against a fake.
    // ###########################################################################################
    public sealed class MySqlCollectionStore : ICollectionStore
    {
        private const string CollectionColumns =
            "c.id, c.owner_key_hash, c.connection, c.hardware_name, c.board_name, c.region, c.title, c.description, c.forum_links_json, " +
            "c.created_utc, c.updated_utc, c.delete_at_utc, c.deleted_utc, c.deleted_reason, c.unshared_utc, c.components_json, " +
            "c.oscilloscopes_json";

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

        public async Task<IReadOnlyList<CollectionRecord>> ListAllLiveAsync(int count, CancellationToken cancellationToken = default)
        {
            await using MySqlConnection connection = await this.OpenAsync(cancellationToken);
            await using MySqlCommand select = connection.CreateCommand();

            select.CommandText =
                $"SELECT {CollectionColumns}, (SELECT COUNT(*) FROM collection_images i WHERE i.collection_id = c.id) " +
                "FROM collections c WHERE c.deleted_utc IS NULL ORDER BY c.created_utc DESC LIMIT @count;";
            select.Parameters.AddWithValue("@count", count);

            var collections = new List<CollectionRecord>();

            await using MySqlDataReader reader = await select.ExecuteReaderAsync(cancellationToken);

            while (await reader.ReadAsync(cancellationToken))
                collections.Add(MySqlCollectionStore.ReadCollection(reader));

            return collections;
        }

        public async Task<int> CountCreatedSinceAsync(string connection, DateTimeOffset sinceUtc, CancellationToken cancellationToken = default)
        {
            await using MySqlConnection database = await this.OpenAsync(cancellationToken);
            await using MySqlCommand count = database.CreateCommand();

            count.CommandText = "SELECT COUNT(*) FROM collections WHERE connection = @connection AND created_utc > @since;";
            count.Parameters.AddWithValue("@connection", connection);
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
                    "INSERT IGNORE INTO collections (id, owner_key_hash, connection, hardware_name, board_name, region, title, description, " +
                    "forum_links_json, components_json, oscilloscopes_json, created_utc, updated_utc, delete_at_utc) VALUES " +
                    "(@id, @owner, @connection, @hardware, @board, @region, @title, @description, @links, @components, @oscilloscopes, @created, " +
                    "@updated, @deleteAt);";
                MySqlCollectionStore.AddContent(insert, collection);
                insert.Parameters.AddWithValue("@owner", (object?)collection.OwnerKeyHash ?? DBNull.Value);
                insert.Parameters.AddWithValue("@connection", (object?)collection.Connection ?? DBNull.Value);
                insert.Parameters.AddWithValue("@created", collection.CreatedUtc.UtcDateTime);

                // The id taken already: nothing written, and CollectionFlows draws another.
                if (await insert.ExecuteNonQueryAsync(cancellationToken) == 0)
                    return false;
            }

            await MySqlCollectionStore.InsertImagesAsync(connection, transaction, collection, cancellationToken);
            await transaction.CommitAsync(cancellationToken);

            return true;
        }

        public async Task<CollectionRecord?> FindLiveByOwnerAsync(string ownerKeyHash, CancellationToken cancellationToken = default)
        {
            string? id;

            await using (MySqlConnection connection = await this.OpenAsync(cancellationToken))
            await using (MySqlCommand command = connection.CreateCommand())
            {
                command.CommandText =
                    "SELECT id FROM collections WHERE owner_key_hash = @owner AND deleted_utc IS NULL ORDER BY created_utc DESC LIMIT 1;";
                command.Parameters.AddWithValue("@owner", ownerKeyHash);

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
                    "description = @description, forum_links_json = @links, components_json = @components, " +
                    "oscilloscopes_json = @oscilloscopes, updated_utc = @updated, delete_at_utc = @deleteAt, " +
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

        public async Task<bool> RemoveImageAsync(string id, string sha256, DateTimeOffset whenUtc, CancellationToken cancellationToken = default)
        {
            await using MySqlConnection connection = await this.OpenAsync(cancellationToken);
            await using MySqlTransaction transaction = await connection.BeginTransactionAsync(cancellationToken);

            // The collection's row first, so a delete or a share running at the same time waits for it - and only while it is
            // not deleted. Its date and its sharing are never named here.
            await using (MySqlCommand update = connection.CreateCommand())
            {
                update.Transaction = transaction;
                update.CommandText =
                    "UPDATE collections SET updated_utc = @updated WHERE id = @id AND deleted_utc IS NULL;" +
                    "SELECT COUNT(*) FROM collections WHERE id = @id AND deleted_utc IS NULL;";
                update.Parameters.AddWithValue("@id", id);
                update.Parameters.AddWithValue("@updated", whenUtc.UtcDateTime);

                if (Convert.ToInt32(await update.ExecuteScalarAsync(cancellationToken)) == 0)
                    return false;
            }

            await using (MySqlCommand delete = connection.CreateCommand())
            {
                delete.Transaction = transaction;
                delete.CommandText = "DELETE FROM collection_images WHERE collection_id = @id AND sha256 = @sha256;";
                delete.Parameters.AddWithValue("@id", id);
                delete.Parameters.AddWithValue("@sha256", sha256);
                await delete.ExecuteNonQueryAsync(cancellationToken);
            }

            await transaction.CommitAsync(cancellationToken);
            return true;
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
                    "forum_links_json = NULL, components_json = NULL, oscilloscopes_json = NULL WHERE id = @id AND deleted_utc IS NULL;";
                update.Parameters.AddWithValue("@id", id);
                update.Parameters.AddWithValue("@when", whenUtc.UtcDateTime);
                update.Parameters.AddWithValue("@reason", reason);

                if (await update.ExecuteNonQueryAsync(cancellationToken) == 0)
                    return false;
            }

            // Its captures go with it - and the comments written before server 17.0.0, which nothing reads any more (migration
            // 0031), so a collection deleted leaves none behind.
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

        public async Task<IReadOnlyDictionary<string, CollectionUploadRecord>> FindUploadsAsync(
            string ownerKeyHash,
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
                "SELECT owner_key_hash, sha256, byte_count, width, height, uploaded_utc, connection FROM collection_uploads " +
                $"WHERE owner_key_hash = @owner AND sha256 IN ({string.Join(", ", names)});";
            select.Parameters.AddWithValue("@owner", ownerKeyHash);

            int at = 0;

            foreach (string sha in sha256s)
                select.Parameters.AddWithValue(names[at++], sha);

            await using MySqlDataReader reader = await select.ExecuteReaderAsync(cancellationToken);

            while (await reader.ReadAsync(cancellationToken))
            {
                var upload = new CollectionUploadRecord(
                    reader.GetString(0),
                    reader.GetString(1),
                    reader.GetInt32(2),
                    reader.GetInt32(3),
                    reader.GetInt32(4),
                    MySqlCollectionStore.Utc(reader.GetDateTime(5)),
                    MySqlCollectionStore.Text(reader, 6));

                found[upload.Sha256] = upload;
            }

            return found;
        }

        public async Task AddUploadAsync(CollectionUploadRecord upload, CancellationToken cancellationToken = default)
        {
            ArgumentNullException.ThrowIfNull(upload);

            await using MySqlConnection connection = await this.OpenAsync(cancellationToken);
            await using MySqlCommand insert = connection.CreateCommand();

            // One row per key and picture (ux_collection_uploads_owner, migration 0031).
            insert.CommandText =
                "INSERT INTO collection_uploads (owner_key_hash, connection, sha256, byte_count, width, height, uploaded_utc) " +
                "VALUES (@owner, @connection, @sha, @bytes, @width, @height, @when) " +
                "ON DUPLICATE KEY UPDATE uploaded_utc = GREATEST(uploaded_utc, VALUES(uploaded_utc));";
            insert.Parameters.AddWithValue("@owner", upload.OwnerKeyHash);
            insert.Parameters.AddWithValue("@connection", (object?)upload.Connection ?? DBNull.Value);
            insert.Parameters.AddWithValue("@sha", upload.Sha256);
            insert.Parameters.AddWithValue("@bytes", upload.ByteCount);
            insert.Parameters.AddWithValue("@width", upload.Width);
            insert.Parameters.AddWithValue("@height", upload.Height);
            insert.Parameters.AddWithValue("@when", upload.UploadedUtc.UtcDateTime);

            await insert.ExecuteNonQueryAsync(cancellationToken);
        }

        public async Task<int> CountUploadsSinceAsync(string connection, DateTimeOffset sinceUtc, CancellationToken cancellationToken = default)
        {
            await using MySqlConnection database = await this.OpenAsync(cancellationToken);
            await using MySqlCommand count = database.CreateCommand();

            count.CommandText = "SELECT COUNT(*) FROM collection_uploads WHERE connection = @connection AND uploaded_utc > @since;";
            count.Parameters.AddWithValue("@connection", connection);
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

            var unused = new List<(long Id, string Sha)>();

            await using (MySqlCommand select = connection.CreateCommand())
            {
                select.Transaction = transaction;

                // An upload is in use while a collection not deleted of ITS OWNER shows it - the same key, or, for an upload from
                // before server 17.0.0 (migration 0031), the same account - a capture, or a capture's pinout drawing (0029).
                select.CommandText =
                    "SELECT u.id, u.sha256 FROM collection_uploads u WHERE u.uploaded_utc < @before AND NOT EXISTS (" +
                    "SELECT 1 FROM collection_images i JOIN collections c ON c.id = i.collection_id " +
                    "WHERE c.deleted_utc IS NULL AND (i.sha256 = u.sha256 OR i.pinout_sha256 = u.sha256) AND (" +
                    "c.owner_key_hash = u.owner_key_hash OR (u.account_id IS NOT NULL AND c.account_id = u.account_id))) FOR UPDATE;";
                select.Parameters.AddWithValue("@before", uploadedBeforeUtc.UtcDateTime);

                await using MySqlDataReader reader = await select.ExecuteReaderAsync(cancellationToken);

                while (await reader.ReadAsync(cancellationToken))
                    unused.Add((Convert.ToInt64(reader.GetValue(0)), reader.GetString(1)));
            }

            foreach ((long id, _) in unused)
            {
                await using MySqlCommand delete = connection.CreateCommand();

                delete.Transaction = transaction;
                delete.CommandText = "DELETE FROM collection_uploads WHERE id = @id;";
                delete.Parameters.AddWithValue("@id", id);
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
            command.Parameters.AddWithValue("@oscilloscopes", collection.Oscilloscopes is not { Count: > 0 } oscilloscopes ? DBNull.Value : JsonSerializer.Serialize(oscilloscopes));
            command.Parameters.AddWithValue("@updated", collection.UpdatedUtc.UtcDateTime);
            command.Parameters.AddWithValue("@deleteAt", collection.DeleteAtUtc.UtcDateTime);
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

        // CollectionColumns, then the image count.
        private static CollectionRecord ReadCollection(MySqlDataReader reader)
        {
            string? links = MySqlCollectionStore.Text(reader, 8);

            return new CollectionRecord(
                reader.GetString(0),
                MySqlCollectionStore.Text(reader, 1),
                reader.GetString(3),
                reader.GetString(4),
                MySqlCollectionStore.Text(reader, 5),
                MySqlCollectionStore.Text(reader, 6),
                MySqlCollectionStore.Text(reader, 7),
                links is null ? [] : JsonSerializer.Deserialize<List<string>>(links) ?? [],
                MySqlCollectionStore.Utc(reader.GetDateTime(9)),
                MySqlCollectionStore.Utc(reader.GetDateTime(10)),
                MySqlCollectionStore.Utc(reader.GetDateTime(11)),
                reader.IsDBNull(12) ? null : MySqlCollectionStore.Utc(reader.GetDateTime(12)),
                MySqlCollectionStore.Text(reader, 13),
                [],
                Convert.ToInt32(reader.GetValue(17)),
                Connection: MySqlCollectionStore.Text(reader, 2),
                UnsharedUtc: reader.IsDBNull(14) ? null : MySqlCollectionStore.Utc(reader.GetDateTime(14)),
                Components: MySqlCollectionStore.Text(reader, 15) is { } components
                    ? JsonSerializer.Deserialize<List<CollectionComponent>>(components) ?? []
                    : [],
                Oscilloscopes: MySqlCollectionStore.Text(reader, 16) is { } oscilloscopes
                    ? JsonSerializer.Deserialize<List<string>>(oscilloscopes) ?? []
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
