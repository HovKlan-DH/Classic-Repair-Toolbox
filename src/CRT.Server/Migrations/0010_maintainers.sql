-- ############################################################################################
-- The role is called MAINTAINER again (owner decision, 2026-09-25): "Reviewer" becomes
-- "maintainer" everywhere, and the review application becomes "CRT Maintainer".
--
-- 0006 renamed the pool table from `maintainers` to `reviewers`; this names it back. Its index and
-- foreign-key constraints kept their `maintainers` prefix through that rename, so after this one
-- the table and its constraints agree again.
--
-- The approval tables STORE the role as text, and 0008 pinned the allowed values with CHECK
-- constraints: each is dropped, its rows rewritten, and it is added back with the new word.
-- MySqlSubmissionStore's RoleText writes exactly these words. The audit trail stores the grant and
-- revoke ACTIONS as text too, and is rewritten the same way, so one event keeps one name.
-- MaintainerRoleMigrationTests holds all of it to the code.
--
-- *** EVERY STATEMENT BEFORE THE LAST CAN RUN TWICE, AND THE RENAME IS THE LAST. *** MariaDB commits
-- each DDL statement on its own, so a failure part-way leaves what already ran done and this
-- migration unrecorded - and the next start runs it again from the top. Each constraint is dropped
-- IF EXISTS, so a second run finds either the old one or the new one and removes it; the UPDATEs
-- change nothing the second time; and the one statement that cannot be repeated, the RENAME, comes
-- last, so it only runs once everything before it has succeeded.
--
-- READ 0001_initial.sql's header before editing this. Once applied, this file is history.
-- ############################################################################################


ALTER TABLE submission_approvals
    DROP CONSTRAINT IF EXISTS ck_submission_approvals_role;

UPDATE submission_approvals
    SET role = 'maintainer'
    WHERE role = 'reviewer';

ALTER TABLE submission_approvals
    ADD CONSTRAINT ck_submission_approvals_role CHECK (role IN ('maintainer', 'administrator'));


ALTER TABLE production_approvals
    DROP CONSTRAINT IF EXISTS ck_production_approvals_role;

UPDATE production_approvals
    SET role = 'maintainer'
    WHERE role = 'reviewer';

ALTER TABLE production_approvals
    ADD CONSTRAINT ck_production_approvals_role CHECK (role IN ('maintainer', 'administrator'));


UPDATE audit
    SET action = 'maintainer.granted'
    WHERE action = 'reviewer.granted';

UPDATE audit
    SET action = 'maintainer.revoked'
    WHERE action = 'reviewer.revoked';


RENAME TABLE reviewers TO maintainers;
