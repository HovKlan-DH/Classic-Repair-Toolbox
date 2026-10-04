-- ############################################################################################
-- WHAT A SUBMISSION CHANGED, AS IT WENT INTO BETA (owner request, 2026-10-04: "Is it possible to
-- summarize each submission change in a textual form ... so it is possible to see what has changed
-- over time? Not necessarily the full details, but at least to get an idea, besides the sometimes
-- vague description from the contributor").
--
-- One row per submission published to BETA: the rows, fields and files the publish changed - CRT.Data's
-- SubmissionChanges as JSON, written by ApprovePublishFlow and read back for the Systems screen's
-- History view. It has to be kept HERE, at the publish: no history of the boards is retained, so the
-- board as it was before is gone the moment the publish has written over it.
--
-- A submission published again (pushed back out of BETA, then approved once more) replaces its row
-- (REPLACE INTO): what it changed THAT time is what BETA last took from it.
--
-- A TABLE OF ITS OWN rather than a column on submissions, for 0014's reason: every submissions
-- SELECT is read by column position (MySqlSubmissionStore.ReadSubmission).
--
--   changes_json  - SubmissionChanges, bounded (SubmissionChangeFacts.ListedPerKind per list).
--   recorded_utc  - when the publish recorded it, by the server's clock.
--
-- Submissions published BEFORE this migration have no row, and the History view says nothing about
-- what they changed - there is no board left to compare them with.
--
-- READ 0001_initial.sql's header before editing this. Once applied, this file is history.
-- ############################################################################################

CREATE TABLE submission_changes (
    submission_id   BIGINT UNSIGNED  NOT NULL,
    changes_json    MEDIUMTEXT       NOT NULL,
    recorded_utc    DATETIME(3)      NOT NULL,

    PRIMARY KEY (submission_id),

    CONSTRAINT fk_submission_changes_submission
        FOREIGN KEY (submission_id) REFERENCES submissions (id)
        ON DELETE CASCADE
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;
