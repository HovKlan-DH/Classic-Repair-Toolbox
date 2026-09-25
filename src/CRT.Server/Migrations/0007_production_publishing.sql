-- ############################################################################################
-- Publishing to PRODUCTION (maintainer request, 2026-09-25): BETA first, then Production once a
-- reviewer has checked it there.
--
-- current_revision and content_hash already say what is in BETA - every publish writes them.
-- These three say what was last copied to PRODUCTION, so "BETA is ahead of production" is a
-- comparison of two columns (ProductionPromotionRules.IsAwaitingProduction) rather than a walk of
-- two trees.
--
-- All three nullable: a system never promoted has none, which is exactly "waiting for
-- production" once it has a BETA content hash. A shipped board nobody has published through the
-- pipeline has neither, and is waiting for nothing.
--
-- production_published_utc is also what tells a contributor their merged work is live for
-- everyone: a merged submission decided at or before it went out with that promotion
-- (ProductionPromotionRules.ContributorFacingState).
--
-- READ 0001_initial.sql's header before editing this. Once applied, this file is history.
-- ############################################################################################
ALTER TABLE systems
    ADD COLUMN production_revision VARCHAR(64) NULL,
    ADD COLUMN production_content_hash VARCHAR(64) NULL,
    ADD COLUMN production_published_utc DATETIME(3) NULL;
