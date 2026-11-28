-- ############################################################################################
-- COLLECTIONS SHARED WITH NO ACCOUNT (owner decision, 2026-11-28: "Entirely remove the new "Account" possibility ... I do not
-- want the possibility for everyone to create an account"; asked who then shares a collection: "Anybody, no account";
-- server 17.0.0):
--
--   collections.owner_key_hash         SHA-256 (lower-case hex) of the key CRT made for the collection as it first shared it,
--                                      kept with the collection on that computer alone - only that key changes or deletes the
--                                      page. NULL for one shared with an account before: nobody changes it any more, and it
--                                      goes at its date (or by the administrator's delete on the web).
--   collections.connection             the internet connection it was first shared from - an IPv4 address, an IPv6 /64 - what
--                                      the day's limit counts (CollectionRules.MaximumCollectionsPerDay).
--   collections.account_id             may be NULL from now on; kept for the collections shared before.
--
--   collection_uploads                 an upload belongs to the KEY it was sent with (owner_key_hash), counted by its
--                                      connection; its rows keyed by an id of their own, as account_id may be NULL now.
--
-- The comments (0026) are no longer read or written: the table stays as it is, untouched.
--
-- READ 0001_initial.sql's header before editing this. Once applied, this file is history.
-- ############################################################################################

ALTER TABLE collections
    MODIFY COLUMN account_id BIGINT UNSIGNED NULL,
    ADD COLUMN owner_key_hash CHAR(64) NULL AFTER account_id,
    ADD COLUMN connection VARCHAR(64) NULL AFTER owner_key_hash,
    ADD KEY ix_collections_owner_key (owner_key_hash, deleted_utc),
    ADD KEY ix_collections_connection (connection, created_utc);

ALTER TABLE collection_uploads
    DROP PRIMARY KEY,
    ADD COLUMN id BIGINT UNSIGNED NOT NULL AUTO_INCREMENT PRIMARY KEY FIRST,
    MODIFY COLUMN account_id BIGINT UNSIGNED NULL,
    ADD COLUMN owner_key_hash CHAR(64) NULL AFTER account_id,
    ADD COLUMN connection VARCHAR(64) NULL AFTER owner_key_hash,
    ADD UNIQUE KEY ux_collection_uploads_owner (owner_key_hash, sha256),
    ADD KEY ix_collection_uploads_connection (connection, uploaded_utc);
