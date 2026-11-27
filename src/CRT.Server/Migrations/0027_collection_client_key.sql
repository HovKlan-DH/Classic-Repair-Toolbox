-- ############################################################################################
-- A COLLECTION'S KEY FROM CRT (code review, 2026-11-27 - CollectionFlows.CreateAsync): a FIRST share whose answer was lost -
-- the connection dropped after the server had kept it - made a second page when "Share" was pressed again, and the first
-- one's id never reached CRT, so only the administrator could take it down. CRT now sends the collection's own local id
-- with a first share (the Idempotency-Key header), and a first share sent again with the same key changes the page it made
-- instead of making another.
--
--   client_key  the local id CRT keeps the collection under - 32 lower-case hex digits - or NULL for one shared before.
--               Not unique: a collection deleted keeps its row, and CRT shares it again under the same key - only a
--               collection still shown is looked for by it.
--
-- READ 0001_initial.sql's header before editing this. Once applied, this file is history.
-- ############################################################################################

ALTER TABLE collections
    ADD COLUMN client_key CHAR(32) NULL AFTER account_id,
    ADD KEY ix_collections_client_key (account_id, client_key, deleted_utc);
