-- ############################################################################################
-- Reconciling the schema with the code that was written against it (2026-09-21).
--
-- THIS MIGRATION EXISTS BECAUSE THREE DISAGREEMENTS WERE FOUND BY READING, NOT BY TESTING. The
-- whole submission pipeline is covered by tests, and every one of them uses an in-memory fake
-- store - so no test has ever executed MySqlSubmissionStore against MariaDB, and none of these
-- would have been caught before the first real submission failed on the live box.
--
-- The three:
--
--   1. submissions.state's CHECK constraint did not allow the states the code writes. 0001 listed
--      the REVIEW lifecycle ('pending', 'changes_requested', 'approved', 'rejected', 'withdrawn',
--      'merged'), which is Phase 5's vocabulary. Phase 4 then added the TRANSPORT states the code
--      actually writes first - 'uploading' when a submission is created, and 'abandoned' when its
--      upload window expires - and neither was in the list. Every create would have failed.
--
--   2. submissions.system_id is NOT NULL with a foreign key to systems(system_id), but nothing
--      ever inserted into `systems`. The table is empty, so every submission would have failed on
--      the foreign key. See the new stored behaviour in MySqlSubmissionStore: a submission now
--      ensures its system row exists first.
--
--   3. Nothing recorded a system's CURRENT REVISION as published, which task 7's system.json
--      needs and which `systems.current_revision` was already there to hold.
--
-- READ 0001_initial.sql's header before editing this. Once applied, this file is history.
-- ############################################################################################


-- --------------------------------------------------------------------------------------------
-- submissions.state - the full lifecycle, transport and review together.
--
-- A submission's life is: uploading -> pending -> (a reviewer decides) -> approved/rejected/
-- changes_requested -> merged, with 'withdrawn' available to the contributor and 'abandoned' for
-- an upload window that expired without being finalised.
--
-- 'uploading' and 'abandoned' are the two this adds. They are not review states and a reviewer
-- never sees them - which is exactly why they were missed when 0001 was written from the review
-- model, and why the constraint has to describe the whole life rather than one phase of it.
--
-- MariaDB requires the old constraint to be dropped before an identically-named one is added.
-- --------------------------------------------------------------------------------------------
ALTER TABLE submissions
    DROP CONSTRAINT ck_submissions_state;

ALTER TABLE submissions
    ADD CONSTRAINT ck_submissions_state
    CHECK (state IN (
        -- Transport (Phase 4): the contribution is still arriving, or never finished.
        'uploading',
        'abandoned',

        -- Review (Phase 5): it arrived intact and is somebody's to decide.
        'pending',
        'changes_requested',
        'approved',
        'rejected',
        'withdrawn',
        'merged'
    ));


-- --------------------------------------------------------------------------------------------
-- systems.origin - say what the two values are, in the schema rather than only in code.
--
-- 'shipped' for a system that came with CRT, 'contributed' for one that arrived through this
-- pipeline and was vetted. The column already existed and already carried that intent in its
-- comment; this makes the database refuse a third value rather than trusting every writer.
--
-- SET ONCE, WHEN THE SYSTEM ROW IS CREATED, and never recomputed - a system that arrived through
-- the pipeline stays 'contributed' however many times it is later revised, including by the
-- maintainer. It records where the system CAME FROM, not who touched it last.
-- --------------------------------------------------------------------------------------------
ALTER TABLE systems
    ADD CONSTRAINT ck_systems_origin
    CHECK (origin IN ('shipped', 'contributed'));


-- --------------------------------------------------------------------------------------------
-- systems.current_revision already exists and stays nullable, deliberately.
--
-- NULL means "this system has never been published" - which is the state of every system created
-- by a first-time contribution, from the moment its first submission arrives until a maintainer
-- publishes it. That is a real and common state, not missing data, so it must not be defaulted to
-- an empty string: "" and NULL would then both occur and mean the same thing, which is how a
-- column stops being answerable.
--
-- Nothing to do here. Recorded so the next person does not "fix" the nullability.
-- --------------------------------------------------------------------------------------------


-- --------------------------------------------------------------------------------------------
-- systems.content_hash - what the published tree currently hashes to (task 7).
--
-- Mirrors system.json's own ContentHash, so the database can answer "is this client holding the
-- current revision" without reading a file out of the published tree - which the service cannot
-- do anyway for Production, since it has no read-write access there by design.
--
-- Nullable for the same reason current_revision is: an unpublished system has no published
-- content to hash.
-- --------------------------------------------------------------------------------------------
ALTER TABLE systems
    ADD COLUMN content_hash VARCHAR(64) NULL AFTER current_revision;
