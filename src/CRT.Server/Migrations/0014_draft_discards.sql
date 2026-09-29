-- ############################################################################################
-- THE CONTRIBUTOR DISCARDED THEIR OWN DRAFT (owner request, 2026-09-28): "if the contributor has
-- discarded his own data, I think this needs to be validated and by sure it should be clearly
-- visible on the system and for the maintainer(s)".
--
-- One row per submission whose contributor discarded their draft of that board in CRT after
-- sending it - CRT reports it once per live submission (CRT.Data's DraftDiscardContract), proving
-- ownership with the submission's capability token. The review queue, "Beta > Prod" and the
-- Systems screen show it beside the submission, so a maintainer does not publish work its own
-- author has thrown away without asking them first.
--
-- A TABLE OF ITS OWN rather than a column on submissions: every submissions SELECT is read by
-- column position (MySqlSubmissionStore.ReadSubmission), and one more column there would have to
-- be added at the end of six queries in step. The row is also the whole record: a second notice
-- for the same submission keeps the FIRST time (INSERT IGNORE).
--
--   discarded_utc - when the notice arrived, by the server's clock.
--
-- READ 0001_initial.sql's header before editing this. Once applied, this file is history.
-- ############################################################################################

CREATE TABLE submission_draft_discards (
    submission_id   BIGINT UNSIGNED  NOT NULL,
    discarded_utc   DATETIME(3)      NOT NULL,

    PRIMARY KEY (submission_id),

    CONSTRAINT fk_submission_draft_discards_submission
        FOREIGN KEY (submission_id) REFERENCES submissions (id)
        ON DELETE CASCADE
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;
