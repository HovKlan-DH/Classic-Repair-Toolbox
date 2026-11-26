-- ############################################################################################
-- COMMENTS ON A SHARED COLLECTION'S CAPTURES (step 3, 2026-11-26 - CRT.Data's CollectionComments): written on a collection's
-- page by anybody signed in, and by its owner from CRT.
--
--   collection_id, image_sha256  the capture it is on - by its checksum, which a collection keeps for a capture however
--                                often it is shared again.
--   account_id                   who wrote it - their display name is shown beside it.
--   body                         plain text, as typed (at most 2000 characters).
--   created_utc, edited_utc      written, and last changed by its writer.
--   hidden_utc                   hidden by the collection's owner: gone from the page, still listed for the owner.
--   deleted_utc                  deleted by its writer or an administrator: gone - the row stays a moment for nothing, and
--                                goes with its collection.
--
-- A collection deleted takes its comments with it (CollectionFlows), as its captures.
--
-- READ 0001_initial.sql's header before editing this. Once applied, this file is history.
-- ############################################################################################

CREATE TABLE collection_comments (
    id              BIGINT UNSIGNED  NOT NULL AUTO_INCREMENT,
    collection_id   CHAR(12)         NOT NULL,
    image_sha256    CHAR(64)         NOT NULL,
    account_id      BIGINT UNSIGNED  NOT NULL,
    body            TEXT             NOT NULL,
    created_utc     DATETIME(3)      NOT NULL,
    edited_utc      DATETIME(3)      NULL,
    hidden_utc      DATETIME(3)      NULL,
    deleted_utc     DATETIME(3)      NULL,

    PRIMARY KEY (id),
    KEY ix_collection_comments_collection (collection_id, created_utc),
    KEY ix_collection_comments_account (account_id, created_utc)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;
