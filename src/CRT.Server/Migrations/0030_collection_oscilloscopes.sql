-- ############################################################################################
-- THE OSCILLOSCOPES A COLLECTION'S CAPTURES WERE TAKEN WITH (owner request, 2026-11-28: "the oscilloscope brand and model
-- should be shown there" - the page's top panel; server 16.1.0):
--
--   collections.oscilloscopes_json     each oscilloscope once, its vendor and model as CRT's "Oscilloscope" tab named it
--                                      ("Rigol DS1054Z"), JSON. NULL for none known - a CRT before 16.1.0, or captures taken
--                                      before CRT kept it.
--
-- READ 0001_initial.sql's header before editing this. Once applied, this file is history.
-- ############################################################################################

ALTER TABLE collections
    ADD COLUMN oscilloscopes_json TEXT NULL AFTER components_json;
