-- ############################################################################################
-- "SYSTEM" BECOMES "BOARD" - IN THE SCHEMA TOO (owner decision, 2026-10-09: rename everything,
-- "code too", from the words people read down to the database).
--
-- Every row of `systems` was always one BOARD of a piece of hardware ("Commodore/C64/250407"),
-- and CRT's own drop-down lists have always said Hardware and Board. So the table becomes
-- `boards`, every `system_id` becomes `board_id`, crt_board_views.systemId becomes boardId, and the
-- keys, indexes and the check constraint that carried the old word are renamed with them.
--
-- NOTHING CHANGES BUT NAMES: no type, collation, nullability, order or ON DELETE rule moves.
-- Every column is restated exactly as 0001/0005/0008/0012/0013 left it, because CHANGE COLUMN
-- takes the whole definition. Column POSITIONS are unchanged too, which matters: every submissions
-- SELECT is read by position (MySqlSubmissionStore.ReadSubmission).
--
-- THE FOREIGN KEYS GO FIRST AND COME BACK LAST, as in 0005: the four tables that point at
-- `systems (system_id)` are released, the parent and the children renamed, and the keys re-added
-- under their new names, pointing at `boards (board_id)`.
--
-- STORED WORDS FOLLOW: the audit trail's action words ('system.placed', 'system.deleted',
-- 'system.list_ordered') and the finding codes kept with refused submissions
-- ('system.id-invalid', 'identity.system_id_mismatch', ...) are what the code now writes as
-- 'board.*', so a board's history keeps showing its old placements and deletions. And the one
-- key renamed inside stored JSON: submission_changes.changes_json holds CRT.Data's
-- SubmissionChanges as the store writes it (PascalCase), whose IsNewSystem is now IsNewBoard -
-- left alone, every board created before this would lose its "A new board." in its History.
--
-- The grant is on crt_review.* (INSTALLING.md), so renaming a table keeps the service's access.
--
-- READ 0001_initial.sql's header before editing this. Once applied, this file is history.
-- ############################################################################################


-- --------------------------------------------------------------------------------------------
-- 1. Release the four foreign keys that point at systems (system_id).
-- --------------------------------------------------------------------------------------------
ALTER TABLE maintainers DROP FOREIGN KEY fk_maintainers_system;
ALTER TABLE submissions DROP FOREIGN KEY fk_submissions_system;
ALTER TABLE production_approvals DROP FOREIGN KEY fk_production_approvals_system;
ALTER TABLE maintainer_invitations DROP FOREIGN KEY fk_maintainer_invitations_system;


-- --------------------------------------------------------------------------------------------
-- 2. The table itself.
-- --------------------------------------------------------------------------------------------
RENAME TABLE systems TO boards;

ALTER TABLE boards
    CHANGE COLUMN system_id board_id VARCHAR(255) NOT NULL COLLATE utf8mb4_bin,
    DROP INDEX ix_systems_manufacturer,
    ADD INDEX ix_boards_manufacturer (manufacturer, hardware, board),
    DROP CONSTRAINT ck_systems_origin,
    ADD CONSTRAINT ck_boards_origin CHECK (origin IN ('shipped', 'contributed'));


-- --------------------------------------------------------------------------------------------
-- 3. The tables that name a board by its id.
-- --------------------------------------------------------------------------------------------
ALTER TABLE maintainers
    CHANGE COLUMN system_id board_id VARCHAR(255) NOT NULL COLLATE utf8mb4_bin;

ALTER TABLE submissions
    CHANGE COLUMN system_id board_id VARCHAR(255) NOT NULL COLLATE utf8mb4_bin,
    DROP INDEX ix_submissions_system,
    ADD INDEX ix_submissions_board (board_id, state);

ALTER TABLE production_approvals
    CHANGE COLUMN system_id board_id VARCHAR(255) NOT NULL COLLATE utf8mb4_bin;

ALTER TABLE maintainer_invitations
    CHANGE COLUMN system_id board_id VARCHAR(255) NOT NULL COLLATE utf8mb4_bin,
    DROP INDEX ix_maintainer_invitations_system,
    ADD INDEX ix_maintainer_invitations_board (board_id);

ALTER TABLE crt_board_views
    CHANGE COLUMN systemId boardId VARCHAR(200) NOT NULL COLLATE utf8mb4_bin,
    DROP INDEX ix_crt_board_views_system,
    ADD INDEX ix_crt_board_views_board (boardId, viewedUtc);


-- --------------------------------------------------------------------------------------------
-- 4. The foreign keys again, under their new names - each ON DELETE CASCADE, as before.
-- --------------------------------------------------------------------------------------------
ALTER TABLE maintainers
    ADD CONSTRAINT fk_maintainers_board
        FOREIGN KEY (board_id) REFERENCES boards (board_id)
        ON DELETE CASCADE;

ALTER TABLE submissions
    ADD CONSTRAINT fk_submissions_board
        FOREIGN KEY (board_id) REFERENCES boards (board_id)
        ON DELETE CASCADE;

ALTER TABLE production_approvals
    ADD CONSTRAINT fk_production_approvals_board
        FOREIGN KEY (board_id) REFERENCES boards (board_id)
        ON DELETE CASCADE;

ALTER TABLE maintainer_invitations
    ADD CONSTRAINT fk_maintainer_invitations_board
        FOREIGN KEY (board_id) REFERENCES boards (board_id)
        ON DELETE CASCADE;


-- --------------------------------------------------------------------------------------------
-- 5. The words already stored.
-- --------------------------------------------------------------------------------------------
UPDATE audit
    SET action = REPLACE(action, 'system.', 'board.')
    WHERE action IN ('system.placed', 'system.deleted', 'system.list_ordered');

UPDATE submission_findings
    SET code = REPLACE(code, 'system', 'board')
    WHERE code IN ('system.id-invalid', 'system.id-mismatch', 'system.case-collision',
                   'identity.system_id_malformed', 'identity.system_id_mismatch',
                   'system.closed', 'system_edit.file_unknown', 'promote.outside_system');

UPDATE submission_changes
    SET changes_json = REPLACE(changes_json, '"IsNewSystem"', '"IsNewBoard"')
    WHERE changes_json LIKE '%"IsNewSystem"%';
