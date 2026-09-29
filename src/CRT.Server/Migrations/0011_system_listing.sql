-- ############################################################################################
-- WHERE A NEW SYSTEM GOES IN THE DROP-DOWN LISTS (owner request, 2026-09-27): "When a system is
-- added to BETA, and it is a NEW system, can you then make sure it gets added also to the main
-- Excel data file in the Data root? The maintainer should order the new system, so it becomes
-- visible in the right location for the drop-down lists. This must be done before it can be pushed
-- to BETA."
--
-- A maintainer places the system in the Systems screen; this is what they chose. The publish to
-- BETA writes it as one row of the newest main Excel data file (MasterListing), and a promotion to
-- production inserts the same row next to the same neighbours there.
--
--   listing_hardware_name / listing_board_name - the names in CRT's two drop-downs.
--   listing_notes                              - the "Hardware notes in Overview tab" text.
--   listing_after                              - the board workbook ("Excel data file") of the row it
--                                                goes after; NULL puts it first.
--   listing_set_by / listing_set_utc           - who placed it, and when.
--
-- All NULL for every system nobody has placed: an existing board is already listed, and a new
-- one waits here until it is placed. listing_hardware_name being NULL is "not placed".
--
-- READ 0001_initial.sql's header before editing this. Once applied, this file is history.
-- ############################################################################################
ALTER TABLE systems
    ADD COLUMN listing_hardware_name VARCHAR(100) NULL,
    ADD COLUMN listing_board_name VARCHAR(100) NULL,
    ADD COLUMN listing_notes TEXT NULL,
    ADD COLUMN listing_after VARCHAR(512) NULL,
    ADD COLUMN listing_set_by BIGINT NULL,
    ADD COLUMN listing_set_utc DATETIME(3) NULL;
