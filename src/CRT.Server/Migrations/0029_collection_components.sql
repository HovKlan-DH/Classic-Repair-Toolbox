-- ############################################################################################
-- A CAPTURE SEEN LARGE TELLS OF ITS COMPONENT, AND NO BASELINE (owner requests, 2026-11-27 - server 16.0.0):
--
--   collections.components_json        what CRT's component info window said of each component captured, as it was captured
--                                      (CollectionComponent, JSON) - shown beside a capture seen large. NULL for none.
--   collection_images.pinout_sha256    CRT's drawing of the component's pinout with this capture's pin marked - a picture
--                                      the account uploaded, as the captures are. NULL for none.
--   collection_images.baseline_*       dropped - CRT's own baseline beside a capture is no more ("Do not include the baseline
--                                      image, as that is not relevant" - removed entirely, owner decision); nothing reads or
--                                      writes them since server 16.0.0.
--
-- READ 0001_initial.sql's header before editing this. Once applied, this file is history.
-- ############################################################################################

ALTER TABLE collections
    ADD COLUMN components_json MEDIUMTEXT NULL AFTER forum_links_json;

ALTER TABLE collection_images
    ADD COLUMN pinout_sha256 CHAR(64) NULL AFTER sha256,
    DROP COLUMN baseline_sha256,
    DROP COLUMN baseline_file_name,
    DROP COLUMN baseline_time_per_division,
    DROP COLUMN baseline_volts_per_division,
    DROP COLUMN baseline_trigger_level;
