-- ############################################################################################
-- THE FILES A SUBMISSION LEFT EXACTLY AS THEY WERE PUBLISHED (code review, 2026-10-13).
--
-- A submission carries EVERY file its board cites - the contributor's draft is a whole copy of the
-- board - with the hash each had in the draft. Most of them the contributor never touched: they are
-- byte-identical to what was published at the same path when the submission was created, and the
-- create took them from the server's own disk rather than having them uploaded.
--
-- That fact used to be forgotten. At approve, PublishPlan writes every file whose published hash
-- differs from the submission's - and for a file the contributor left alone, a difference means
-- somebody ELSE changed it since (another submission replacing a shared image, say). The old bytes
-- went back over the new ones, for every board citing the file.
--
-- One row per file the submission carried unchanged, as the SERVER found it at create (never
-- anything the contributor says): its path and the hash both sides had. The approval leaves such a
-- file as it is published now when that copy has moved since (SubmissionLeftAsPublished).
--
--   path    - data-root-relative, as in submission_files (utf8mb4_bin: case matters on the server).
--   sha256  - the hash the submission and the published tree both had at create.
--
-- A TABLE OF ITS OWN rather than a column on submission_files: an amendment rewrites that table
-- whole (MySqlSubmissionStore.AmendAsync), and this fact is about what the CONTRIBUTOR sent. An
-- amended file whose hash differs simply does not match its row here, and is written as changed.
--
-- Goes with its submission (ON DELETE CASCADE), so the data reset's DELETE FROM submissions takes it.
--
-- READ 0001_initial.sql's header before editing this. Once applied, this file is history.
-- ############################################################################################

CREATE TABLE submission_left_as_published (
    submission_id   BIGINT UNSIGNED  NOT NULL,
    path            VARCHAR(500)     NOT NULL COLLATE utf8mb4_bin,
    sha256          CHAR(64)         NOT NULL,

    PRIMARY KEY (submission_id, path),

    CONSTRAINT fk_submission_left_as_published_submission
        FOREIGN KEY (submission_id) REFERENCES submissions (id)
        ON DELETE CASCADE
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;
