-- ############################################################################################
-- A SUBMISSION TAKEN BACK OUT OF BETA, RECORDED (code review, 2026-09-29).
--
-- "Push back to queue" returns every submission merged since the last promotion to `pending`,
-- with the maintainer's reason. The contributor is told "returned" ("Taken back out of BETA -
-- waiting for review again"). Until now that word was INFERRED - "pending, and carrying a decision
-- comment" - which held only while no other path ever left a comment on a pending row: a manual
-- fix, a future note to the contributor, or an amendment would have made CRT say a rollback had
-- happened when none had.
--
-- So the rollback now records it: one row per submission it returned, with the time it did so -
-- the same instant it writes as the submission's decided_utc. A submission reads as returned
-- while it is pending AND this time is not older than its latest decision; any decision after the
-- rollback moves decided_utc on and ends it, with nothing having to remember to delete the row. A
-- second rollback of the same submission moves the time (ON DUPLICATE KEY UPDATE).
--
-- A TABLE OF ITS OWN rather than a column on submissions, for 0014's reason: every submissions
-- SELECT is read by column position (MySqlSubmissionStore.ReadSubmission).
--
--   returned_utc - when the rollback returned it, by the server's clock.
--
-- Submissions returned BEFORE this migration have no row and read as plain "pending" from now on;
-- they are few (the rollback shipped 2026-09-27), and their maintainer's comment is still shown.
--
-- READ 0001_initial.sql's header before editing this. Once applied, this file is history.
-- ############################################################################################

CREATE TABLE submission_beta_returns (
    submission_id   BIGINT UNSIGNED  NOT NULL,
    returned_utc    DATETIME(3)      NOT NULL,

    PRIMARY KEY (submission_id),

    CONSTRAINT fk_submission_beta_returns_submission
        FOREIGN KEY (submission_id) REFERENCES submissions (id)
        ON DELETE CASCADE
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;
