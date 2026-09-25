-- ############################################################################################
-- A REVIEWER'S AMENDMENT TO A SUBMISSION (maintainer request, 2026-09-25): the review
-- application's table editor lets a reviewer correct a submission's rows before publishing it.
--
-- submission_payloads and submission_files keep holding the CURRENT content - what review and
-- publish read - so nothing that reads a submission changes. This table keeps what each amendment
-- REPLACED: the first row of a submission holds the contributor's own original rows and files, so
-- the original is never lost and "what did the reviewer change" can always be answered.
--
-- previous_files_json is the file list that was replaced (path, sha256, sizeBytes), because the
-- files follow the rows: a row a reviewer deletes drops the file it cited.
--
-- account_label carries a readable copy of who amended, for the same reason audit.actor_label
-- does: the account may later be renamed or removed.
--
-- READ 0001_initial.sql's header before editing this. Once applied, this file is history.
-- ############################################################################################

CREATE TABLE submission_amendments (
    submission_id         BIGINT UNSIGNED  NOT NULL,
    version               INT              NOT NULL,
    previous_rows_json    LONGTEXT         NOT NULL,
    previous_files_json   LONGTEXT         NOT NULL,
    account_id            BIGINT UNSIGNED  NULL,
    account_label         VARCHAR(255)     NOT NULL,
    amended_utc           DATETIME(3)      NOT NULL,

    PRIMARY KEY (submission_id, version),

    CONSTRAINT fk_submission_amendments_submission
        FOREIGN KEY (submission_id) REFERENCES submissions (id)
        ON DELETE CASCADE,

    CONSTRAINT fk_submission_amendments_account
        FOREIGN KEY (account_id) REFERENCES accounts (id)
        ON DELETE SET NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;
