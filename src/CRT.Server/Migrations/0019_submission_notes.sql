-- ############################################################################################
-- A NEW SYSTEM'S NOTES, AS ITS CONTRIBUTOR WROTE THEM (owner request, 2026-10-05: "that note needs
-- to be sent also to the server, as this notes needs to go into the main Excel in the 'Hardware
-- and Board' sheet and in the column 'Hardware notes in "Overview" tab'").
--
-- One row per submission that carries notes - what the contributor typed under "Notes (optional)"
-- in CRT's "Create system" (CRT.Data's SubmissionManifest.HardwareNotes). Written when the
-- submission is created; read by the Systems screen's placement (SystemListingFlow), whose Notes box
-- starts with the newest notes of the system's live submissions. The maintainer's SAVED placement
-- is what goes into the main Excel data file, so these are a suggestion, never written there as
-- they are.
--
-- A TABLE OF ITS OWN rather than a column on submissions, for 0014's reason: every submissions
-- SELECT is read by column position (MySqlSubmissionStore.ReadSubmission).
--
--   hardware_notes - at most MasterListing.MaximumNotesLength (2,000) characters, which
--                    SubmissionValidator holds a submission to before this row is written.
--
-- Goes with its submission (ON DELETE CASCADE), so the data reset's DELETE FROM submissions takes it.
--
-- READ 0001_initial.sql's header before editing this. Once applied, this file is history.
-- ############################################################################################

CREATE TABLE submission_notes (
    submission_id   BIGINT UNSIGNED  NOT NULL,
    hardware_notes  TEXT             NOT NULL,

    PRIMARY KEY (submission_id),

    CONSTRAINT fk_submission_notes_submission
        FOREIGN KEY (submission_id) REFERENCES submissions (id)
        ON DELETE CASCADE
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;
