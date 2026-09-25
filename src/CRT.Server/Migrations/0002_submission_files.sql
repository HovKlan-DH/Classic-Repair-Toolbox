-- ############################################################################################
-- Phase 4: the files and rows belonging to a submission.
--
-- 0001 created `submissions` as the queue and audit spine. This adds what a submission actually
-- CONTAINS, which 0001 deliberately left out because nothing then needed it.
--
-- READ 0001_initial.sql's header before editing this. The same rules apply: once applied this
-- file is history and must never be changed - the runner records a checksum and refuses to start
-- if it differs. Every later change is a NEW numbered file.
-- ############################################################################################


-- --------------------------------------------------------------------------------------------
-- submission_files - one row per file the submitted system should contain.
--
-- THE MANIFEST IS THE COMPLETE INTENDED STATE, not a list of changes: a file present in the base
-- revision but absent here is a deletion. So this table holds every file of the system as it
-- should read afterwards, not only the ones being uploaded.
--
-- `path` is stored EXACTLY as submitted, case included. The server's filesystem is case-sensitive
-- (Phase 3 put the data tree on Linux), and normalising case here would publish data that works
-- on the contributor's Windows machine and fails everywhere else. The unique index is therefore
-- on a case-sensitive collation - `utf8mb4_bin` - while the rest of the table uses the ordinary
-- one, because two paths differing only in case ARE two different paths to this server.
--
-- `sha256` is the blob's identity and its location in the blob store. It is NOT a foreign key to
-- anything: blobs are content-addressed files on disk, shared across every system that references
-- the same bytes, and giving them a table would mean keeping two records of one truth in step.
--
-- `is_uploaded` records whether the bytes have arrived for THIS submission. A file the server
-- already held is marked uploaded immediately at negotiation time, which is what makes a typo fix
-- cost no upload at all.
-- --------------------------------------------------------------------------------------------
CREATE TABLE submission_files (
    id              BIGINT UNSIGNED  NOT NULL AUTO_INCREMENT,
    submission_id   BIGINT UNSIGNED  NOT NULL,
    path            VARCHAR(500)     NOT NULL COLLATE utf8mb4_bin,
    sha256          CHAR(64)         NOT NULL,
    size_bytes      BIGINT UNSIGNED  NOT NULL,
    is_uploaded     TINYINT(1)       NOT NULL DEFAULT 0,

    PRIMARY KEY (id),
    UNIQUE KEY ux_submission_files_path (submission_id, path),
    KEY ix_submission_files_hash (submission_id, sha256),
    KEY ix_submission_files_pending (submission_id, is_uploaded),

    CONSTRAINT fk_submission_files_submission
        FOREIGN KEY (submission_id) REFERENCES submissions (id)
        ON DELETE CASCADE
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;


-- --------------------------------------------------------------------------------------------
-- submission_payloads - the manifest's ROWS, held as JSON.
--
-- WHY JSON RATHER THAN A TABLE PER BOARDDATA SECTION. Eleven sections, each with its own columns,
-- would mean eleven tables that exist only to be read back out in one piece and handed to
-- CRT.Data - and every schema change to BoardData would need a migration here to match. The rows
-- are never queried field by field on the server: they are stored, diffed as a whole against the
-- base revision, and either published or discarded. A single JSON column matches how the data is
-- actually used, and keeps CRT.Data the one definition of the schema.
--
-- The cost is that the database cannot validate the shape. That is acceptable because
-- SubmissionValidator already has, before this row is ever written, and because a malformed
-- payload here fails at deserialisation with the submission id in hand.
--
-- LONGTEXT rather than JSON: MariaDB's JSON type is an alias for LONGTEXT with a validity check,
-- and the check costs a parse on every write for a guarantee the application has already made.
-- --------------------------------------------------------------------------------------------
CREATE TABLE submission_payloads (
    submission_id   BIGINT UNSIGNED  NOT NULL,
    format_version  INT              NOT NULL,
    rows_json       LONGTEXT         NOT NULL,
    renames_json    LONGTEXT         NULL,
    created_utc     DATETIME(3)      NOT NULL,

    PRIMARY KEY (submission_id),

    CONSTRAINT fk_submission_payloads_submission
        FOREIGN KEY (submission_id) REFERENCES submissions (id)
        ON DELETE CASCADE
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;


-- --------------------------------------------------------------------------------------------
-- submission_findings - what automated validation said about a submission.
--
-- STORED RATHER THAN RECOMPUTED, for two reasons. A rejected submission's findings are what the
-- contributor is shown, possibly days later, and re-running validation to produce them would give
-- a different answer once the base revision has moved on. And a reviewer looking at a queued
-- submission needs to see the warnings that were raised at the time it was accepted.
--
-- Errors and warnings both land here; severity is what distinguishes them. A submission with any
-- error is never queued (see SubmissionValidator.CanBeQueued), so a row with severity 'error'
-- always belongs to a rejected submission.
-- --------------------------------------------------------------------------------------------
CREATE TABLE submission_findings (
    id              BIGINT UNSIGNED  NOT NULL AUTO_INCREMENT,
    submission_id   BIGINT UNSIGNED  NOT NULL,
    severity        VARCHAR(16)      NOT NULL,
    code            VARCHAR(64)      NOT NULL,
    subject         VARCHAR(500)     NULL,
    message         TEXT             NOT NULL,

    PRIMARY KEY (id),
    KEY ix_submission_findings_submission (submission_id, severity),

    CONSTRAINT fk_submission_findings_submission
        FOREIGN KEY (submission_id) REFERENCES submissions (id)
        ON DELETE CASCADE,

    CONSTRAINT ck_submission_findings_severity
        CHECK (severity IN ('warning', 'error'))
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;


-- --------------------------------------------------------------------------------------------
-- submissions gains the columns Phase 4 needs.
--
-- `format_version` so a submission made against an older contract can still be read back, and so
-- a future reader knows which rules produced it.
--
-- `base_revision` already exists in 0001. What is added here is the state a submission passes
-- through while it is being uploaded - 0001's CHECK constraint listed only the states a QUEUED
-- submission can hold, because uploading had not been designed yet.
--
-- MariaDB cannot alter a CHECK constraint in place, so it is dropped and recreated. The new list
-- is a superset: no existing row can violate it.
-- --------------------------------------------------------------------------------------------
ALTER TABLE submissions
    ADD COLUMN format_version INT NOT NULL DEFAULT 1 AFTER state,
    ADD COLUMN expires_utc DATETIME(3) NULL AFTER created_utc;

ALTER TABLE submissions
    DROP CONSTRAINT ck_submissions_state;

-- 'uploading' is the state between creating a submission and finalising it. An abandoned one is
-- garbage-collected on `expires_utc` - blobs from abandoned submissions must be collected or the
-- disk fills quietly, which NewContributeStrategy.md names as a trap.
ALTER TABLE submissions
    ADD CONSTRAINT ck_submissions_state
        CHECK (state IN ('uploading', 'pending', 'changes_requested', 'approved',
                         'rejected', 'withdrawn', 'merged', 'abandoned'));

-- Finding the submissions whose upload window has passed, for the garbage collector.
CREATE INDEX ix_submissions_expiry ON submissions (state, expires_utc);
