-- ############################################################################################
-- Security review, 2026-09-25: the system id becomes CASE-SENSITIVE, and each submission records
-- how many bytes it asked to upload.
--
-- READ 0001_initial.sql's header before editing this. Once applied, this file is history.
-- ############################################################################################


-- --------------------------------------------------------------------------------------------
-- system_id compares BYTE FOR BYTE (utf8mb4_bin), in all three tables that hold it.
--
-- *** IT USED TO FOLD CASE AND ACCENTS, AND THAT LET ANYONE HIJACK A BOARD. *** The columns took
-- the database default, utf8mb4_unicode_ci. The `systems` row is created by the FIRST submission
-- for a board, from that submission's own spelling, before any review - so an anonymous
-- submission for "commodore/c64/250425" created the row, and because the key ignored case, every
-- later real submission for "Commodore/C64/250425" attached to THAT row. It then reloaded with
-- the lowercase names and was rejected at finalise with "system identifier does not match",
-- blaming the contributor's client. Every shipped board without a row yet was open to it.
--
-- With a binary key a case-variant is simply a different row, and the rules in
-- SubmissionFileRules and PublishPlan refuse it against the published tree - where a case
-- variant does real harm, because Windows and macOS clients merge the two folders.
--
-- The foreign keys must be dropped first: MariaDB requires a foreign key and the key it
-- references to share a collation, so the three columns change together and the keys return
-- afterwards. No existing row can conflict - a case-insensitive primary key cannot hold two
-- spellings of one id, so every id is already unique under the stricter comparison.
-- --------------------------------------------------------------------------------------------
ALTER TABLE maintainers DROP FOREIGN KEY fk_maintainers_system;
ALTER TABLE submissions DROP FOREIGN KEY fk_submissions_system;

ALTER TABLE systems
    MODIFY system_id VARCHAR(255) NOT NULL COLLATE utf8mb4_bin;

ALTER TABLE submissions
    MODIFY system_id VARCHAR(255) NOT NULL COLLATE utf8mb4_bin;

ALTER TABLE maintainers
    MODIFY system_id VARCHAR(255) NOT NULL COLLATE utf8mb4_bin;

ALTER TABLE submissions
    ADD CONSTRAINT fk_submissions_system
        FOREIGN KEY (system_id) REFERENCES systems (system_id)
        ON DELETE CASCADE;

ALTER TABLE maintainers
    ADD CONSTRAINT fk_maintainers_system
        FOREIGN KEY (system_id) REFERENCES systems (system_id)
        ON DELETE CASCADE;


-- --------------------------------------------------------------------------------------------
-- submissions.bytes_to_upload - how many bytes the server asked this submission to upload.
--
-- What the per-address budget in SubmissionRateLimitPolicy sums: the files the server did NOT
-- already hold, as counted at create. Existing rows default to 0, which is the honest answer for
-- rows that predate the count and costs nothing - they are older than the policy's one-day window
-- by the time this runs anyway.
-- --------------------------------------------------------------------------------------------
ALTER TABLE submissions
    ADD COLUMN bytes_to_upload BIGINT UNSIGNED NOT NULL DEFAULT 0 AFTER summary;
