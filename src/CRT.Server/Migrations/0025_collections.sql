-- ############################################################################################
-- OSCILLOSCOPE COLLECTIONS (owner requests, 2026-11-25 - CRT.Data's CollectionRules): captures taken in CRT's component info
-- window, shared from CRT by an account as a page anybody with its link can see, for help from a forum.
--
-- collections - one row per collection shared:
--
--   id                    the public id, 12 characters (CollectionRules.NewId) - the page's address.
--   account_id            the account that shared it. NEVER SHOWN on the page (owner decision, 2026-11-25).
--   hardware_name,        the board, as CRT names it; region, the one shown when captured.
--   board_name, region
--   title, description    the owner's words, as typed.
--   forum_links_json      the forum threads it was posted in, a JSON array of web addresses.
--   created_utc, updated_utc
--   delete_at_utc         when the sweep deletes it (CollectionSweeper).
--   reminded_utc          when the owner was mailed that it goes in 30 days - once.
--   deleted_utc           set once deleted - its row stays, its words emptied, so the page can say when it went.
--   deleted_reason        'expired', 'owner' or 'administrator'.
--
-- collection_images - its captures, in the page's order (position from 0). Each names its picture by checksum; the bytes lie
-- under the CollectionsRoot setting, "<ab>/<sha256>.png" (CollectionFiles), once however many collections show them.
--
-- collection_uploads - which account uploaded which picture, and when: a collection may show only pictures its own account
-- uploaded, the day's limit is counted here, and a picture no collection of its account shows any more is deleted by the
-- sweep a day after its upload.
--
-- READ 0001_initial.sql's header before editing this. Once applied, this file is history.
-- ############################################################################################

CREATE TABLE collections (
    id                CHAR(12)         NOT NULL,
    account_id        BIGINT UNSIGNED  NOT NULL,
    hardware_name     VARCHAR(100)     NOT NULL,
    board_name        VARCHAR(100)     NOT NULL,
    region            VARCHAR(100)     NULL,
    title             VARCHAR(200)     NULL,
    description       TEXT             NULL,
    forum_links_json  TEXT             NULL,
    created_utc       DATETIME(3)      NOT NULL,
    updated_utc       DATETIME(3)      NOT NULL,
    delete_at_utc     DATETIME(3)      NOT NULL,
    reminded_utc      DATETIME(3)      NULL,
    deleted_utc       DATETIME(3)      NULL,
    deleted_reason    VARCHAR(16)      NULL,

    PRIMARY KEY (id),
    KEY ix_collections_account (account_id, deleted_utc, created_utc),
    KEY ix_collections_delete_at (deleted_utc, delete_at_utc)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

CREATE TABLE collection_images (
    collection_id                 CHAR(12)      NOT NULL,
    position                      INT           NOT NULL,
    sha256                        CHAR(64)      NOT NULL,
    width                         INT           NOT NULL,
    height                        INT           NOT NULL,
    component_label               VARCHAR(100)  NOT NULL,
    pin                           VARCHAR(100)  NULL,
    signal_name                   VARCHAR(100)  NULL,
    expected_reading              VARCHAR(100)  NULL,
    time_per_division             VARCHAR(40)   NULL,
    volts_per_division            VARCHAR(40)   NULL,
    trigger_level                 VARCHAR(40)   NULL,
    baseline_sha256               CHAR(64)      NULL,
    baseline_file_name            VARCHAR(260)  NULL,
    baseline_time_per_division    VARCHAR(40)   NULL,
    baseline_volts_per_division   VARCHAR(40)   NULL,
    baseline_trigger_level        VARCHAR(40)   NULL,
    caption                       TEXT          NULL,
    looks_wrong                   TINYINT(1)    NOT NULL DEFAULT 0,
    captured_utc                  DATETIME(3)   NULL,

    PRIMARY KEY (collection_id, position),
    KEY ix_collection_images_sha (sha256)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

CREATE TABLE collection_uploads (
    account_id    BIGINT UNSIGNED  NOT NULL,
    sha256        CHAR(64)         NOT NULL,
    byte_count    INT              NOT NULL,
    width         INT              NOT NULL,
    height        INT              NOT NULL,
    uploaded_utc  DATETIME(3)      NOT NULL,

    PRIMARY KEY (account_id, sha256),
    KEY ix_collection_uploads_sha (sha256),
    KEY ix_collection_uploads_when (account_id, uploaded_utc)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;
