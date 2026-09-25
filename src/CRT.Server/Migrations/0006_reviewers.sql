-- ############################################################################################
-- Phase 6 roles (maintainer decision, 2026-09-25): TWO roles, Administrator and Reviewer.
--
-- The strategy document planned four. The maintainer collapsed them: a REVIEWER is somebody the
-- administrator assigns to one or more systems, and that person reviews AND publishes changes
-- for exactly those systems. There is no recommend-only role, and being a reviewer is not a flag
-- on the account - it is being in a system's pool. The administrator is in every pool by
-- definition (accounts.is_administrator, unchanged).
--
-- READ 0001_initial.sql's header before editing this. Once applied, this file is history.
-- ############################################################################################


-- --------------------------------------------------------------------------------------------
-- The pool table is named for what it now holds. Everything the maintainers table said about
-- itself in 0001 still applies - a POOL, not an owner; administrators not listed; granted_by
-- for the audit question - only the word changed, and it changed everywhere a person reads it.
-- The foreign-key constraint names keep their old prefix; MariaDB carries them across a rename
-- and nothing reads them by name.
-- --------------------------------------------------------------------------------------------
RENAME TABLE maintainers TO reviewers;


-- --------------------------------------------------------------------------------------------
-- accounts.is_reviewer goes. It meant "may see the queue but never publish", a role that no
-- longer exists; under the new model an account reviews what its pool rows say and nothing
-- else. Dropping the column rather than leaving it dead means no code path can ever read it
-- again and quietly grant the old, global-queue behaviour.
-- --------------------------------------------------------------------------------------------
ALTER TABLE accounts
    DROP COLUMN is_reviewer;


-- --------------------------------------------------------------------------------------------
-- submissions.touches_shared_files - set at creation, from the manifest against the published
-- tree (SubmissionSharedFiles in CRT.Data).
--
-- Shared files belong to no system, so a submission that adds or changes one is the
-- ADMINISTRATOR's to review, whichever board it names. Stored on the row so the queue can be
-- filtered and a request refused without loading the payload every time.
-- --------------------------------------------------------------------------------------------
ALTER TABLE submissions
    ADD COLUMN touches_shared_files TINYINT(1) NOT NULL DEFAULT 0;
