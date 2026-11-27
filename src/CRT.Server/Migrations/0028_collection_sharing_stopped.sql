-- ############################################################################################
-- SHARING STOPPED, AND THE "LOOKS WRONG" MARK GONE (owner requests, 2026-11-27 - CollectionFlows.SetSharedAsync):
--
--   collections.unshared_utc         when the owner pressed "Stop sharing" in CRT, or NULL while it is shared. The page is then
--                                    offline - its link says it is not shared at the moment, its captures and comments are
--                                    kept - and "Share" brings it back under the SAME link (owner decision: "Offline, same
--                                    link back"). Its deletion date keeps counting.
--   collection_images.looks_wrong    dropped - the owner's "Looks wrong" mark on a capture is no more ("No need for the 'Looks
--                                    wrong' checkbox - remove it"); nothing reads or writes it since server 15.1.0.
--
-- READ 0001_initial.sql's header before editing this. Once applied, this file is history.
-- ############################################################################################

ALTER TABLE collections
    ADD COLUMN unshared_utc DATETIME(3) NULL AFTER deleted_reason;

ALTER TABLE collection_images
    DROP COLUMN looks_wrong;
