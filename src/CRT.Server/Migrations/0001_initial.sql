-- ############################################################################################
-- CRT review database - initial schema (Phase 3, step 4).
--
-- READ THIS BEFORE EDITING. Once this file has been applied to any database it is HISTORY and
-- must never be changed: the runner records a checksum and refuses to start if it differs, for
-- the reason MigrationPlan's header gives. Every later change is a NEW numbered file.
--
-- Conventions, applied uniformly so nobody has to remember exceptions:
--
--   - utf8mb4 everywhere. MySQL's old three-byte "utf8" silently mangles anything outside the
--     Basic Multilingual Plane, and this database stores contributor display names and free-text
--     review comments from people all over the world.
--   - InnoDB everywhere, for real foreign keys and transactions.
--   - Times are DATETIME(3) holding UTC. Not TIMESTAMP: that carries implicit timezone conversion
--     tied to the server's setting, so moving the box or changing its zone silently rewrites the
--     meaning of every stored row. The service converts for display; the database stores UTC.
--   - Every table that records a human action names the ACCOUNT, never a display name, because
--     display names change and an audit trail that says "Dennis approved it" stops being true.
--
-- SEVEN tables, where the strategy document's task list named five. `sessions` and
-- `account_tokens` are added here because steps 5 and 6 need them, and a table that belongs in
-- the initial schema should not arrive as a second migration.
-- ############################################################################################


-- --------------------------------------------------------------------------------------------
-- accounts - one row per person.
--
-- EMAIL IS THE IDENTITY and is unique. There is no separate username: a second identifier would
-- be a second thing to recover, and this audience does not want a public handle.
--
-- email_normalised is what the UNIQUE index is on, NOT email. Addresses are case-insensitive in
-- practice for the local part on every mail system CRT's users actually have, so storing
-- "Dennis@example.com" and "dennis@example.com" as two accounts would let one person register
-- twice and would make "an address gets ONE verification mail" untrue. The original spelling is
-- kept in `email` so mail is addressed the way the person wrote it.
--
-- password_hash holds a full Argon2id encoded string ("$argon2id$v=19$m=65536,t=3,p=2$..."),
-- which carries its own parameters. That is what lets the cost be raised later without
-- invalidating existing passwords - verification uses the parameters in the stored hash, and the
-- login path rehashes when they are below the current configuration. 255 chars is comfortably
-- above the ~100 an Argon2id encoding needs.
--
-- is_verified starts 0. An unverified account may log in but may not submit - the distinction
-- matters because blocking login outright makes "resend verification" impossible to reach.
--
-- is_administrator is a column rather than a row in `maintainers`, because the administrator is
-- a maintainer of EVERY system by definition (Phase 6's roles table). Expressing that as rows
-- would mean inserting one per system forever and would make "is this person an administrator"
-- a question about the absence of gaps.
-- --------------------------------------------------------------------------------------------
CREATE TABLE accounts (
    id                  BIGINT UNSIGNED  NOT NULL AUTO_INCREMENT,
    email               VARCHAR(320)     NOT NULL,
    email_normalised    VARCHAR(320)     NOT NULL,
    password_hash       VARCHAR(255)     NOT NULL,
    display_name        VARCHAR(100)     NOT NULL,
    is_verified         TINYINT(1)       NOT NULL DEFAULT 0,
    is_administrator    TINYINT(1)       NOT NULL DEFAULT 0,
    is_reviewer         TINYINT(1)       NOT NULL DEFAULT 0,
    is_locked           TINYINT(1)       NOT NULL DEFAULT 0,
    created_utc         DATETIME(3)      NOT NULL,
    last_login_utc      DATETIME(3)      NULL,

    PRIMARY KEY (id),
    UNIQUE KEY ux_accounts_email_normalised (email_normalised)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;


-- --------------------------------------------------------------------------------------------
-- account_tokens - single-use tokens for email verification and password reset.
--
-- SEPARATE FROM sessions on purpose. These are one-shot, short-lived and granted BEFORE the
-- person has proven anything; a session is granted after. Mixing them in one table means one
-- bug in a WHERE clause turns a password-reset token into a login.
--
-- ONLY THE HASH IS STORED, never the token itself. A database read - a backup left somewhere, a
-- SQL injection in some future endpoint - must not yield working password-reset links. The
-- service hashes the presented token and looks that up. SHA-256 is right here where it would be
-- wrong for passwords: these are 256-bit random values, so there is nothing to brute-force.
--
-- consumed_utc rather than deleting the row: a used token must stay visible long enough for
-- "this link has already been used" to be answerable, and for the audit trail to show that a
-- reset actually completed.
-- --------------------------------------------------------------------------------------------
CREATE TABLE account_tokens (
    id              BIGINT UNSIGNED  NOT NULL AUTO_INCREMENT,
    account_id      BIGINT UNSIGNED  NOT NULL,
    purpose         VARCHAR(32)      NOT NULL,
    token_hash      CHAR(64)         NOT NULL,
    created_utc     DATETIME(3)      NOT NULL,
    expires_utc     DATETIME(3)      NOT NULL,
    consumed_utc    DATETIME(3)      NULL,

    PRIMARY KEY (id),
    UNIQUE KEY ux_account_tokens_hash (token_hash),
    KEY ix_account_tokens_account (account_id, purpose),
    KEY ix_account_tokens_expiry (expires_utc),

    CONSTRAINT fk_account_tokens_account
        FOREIGN KEY (account_id) REFERENCES accounts (id)
        ON DELETE CASCADE
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;


-- --------------------------------------------------------------------------------------------
-- sessions - one row per issued refresh token.
--
-- TOKENS ARE OPAQUE AND DATABASE-BACKED, NOT STATELESS JWTs. This is forced by Phase 6's
-- definition of done: "removing a maintainer takes effect immediately, proven by a test using an
-- already-issued token". No stateless token can satisfy that - its whole point is being valid
-- without a lookup. The cost is one indexed single-row SELECT per authenticated request, which is
-- nothing at this audience size.
--
-- Same hash-only rule as account_tokens, for the same reason.
--
-- ROTATION WITH REUSE DETECTION is what makes a 30-day refresh lifetime acceptable. Each refresh
-- issues a new row and marks the old one rotated, pointing at its successor. If a token that has
-- ALREADY been rotated is presented again, that means two parties hold it - the legitimate client
-- and a thief - and the correct response is to revoke the whole chain rather than to guess which
-- is which. replaced_by_id is what makes that chain walkable.
--
-- user_agent and ip are recorded so a person can recognise their own sessions in a future "signed
-- in devices" view, and so a stolen-credential incident can be reconstructed. Both are truncated
-- rather than trusted: they come from the request.
-- --------------------------------------------------------------------------------------------
CREATE TABLE sessions (
    id                  BIGINT UNSIGNED  NOT NULL AUTO_INCREMENT,
    account_id          BIGINT UNSIGNED  NOT NULL,
    refresh_token_hash  CHAR(64)         NOT NULL,
    created_utc         DATETIME(3)      NOT NULL,
    expires_utc         DATETIME(3)      NOT NULL,
    last_used_utc       DATETIME(3)      NULL,
    revoked_utc         DATETIME(3)      NULL,
    revoked_reason      VARCHAR(64)      NULL,
    replaced_by_id      BIGINT UNSIGNED  NULL,
    user_agent          VARCHAR(255)     NULL,
    created_ip          VARCHAR(45)      NULL,

    PRIMARY KEY (id),
    UNIQUE KEY ux_sessions_refresh_hash (refresh_token_hash),
    KEY ix_sessions_account (account_id, revoked_utc),
    KEY ix_sessions_expiry (expires_utc),

    CONSTRAINT fk_sessions_account
        FOREIGN KEY (account_id) REFERENCES accounts (id)
        ON DELETE CASCADE,

    -- SET NULL, not CASCADE: deleting a successor row must not delete the history that points at
    -- it. The chain losing a link is acceptable; losing the whole chain is not.
    CONSTRAINT fk_sessions_replaced_by
        FOREIGN KEY (replaced_by_id) REFERENCES sessions (id)
        ON DELETE SET NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;


-- --------------------------------------------------------------------------------------------
-- systems - one row per hardware system that can receive contributions.
--
-- system_id is the STRING key the data tree already uses (manufacturer/hardware/board), not a
-- surrogate number, because it is what the desktop app, the file tree and every submission
-- already name. The three parts are stored separately as well so the review app can list by
-- manufacturer without parsing.
--
-- current_revision is the revision now published. It is the base a contributor's work is diffed
-- against, which is why `submissions` records the base revision it started from: a submission
-- built against an older revision needs re-basing rather than blind merging.
--
-- origin records where the system came from (shipped with CRT, or contributed and vetted). It
-- exists so a future "new systems awaiting review" list does not need a separate table.
-- --------------------------------------------------------------------------------------------
CREATE TABLE systems (
    system_id           VARCHAR(255)     NOT NULL,
    manufacturer        VARCHAR(100)     NOT NULL,
    hardware            VARCHAR(100)     NOT NULL,
    board               VARCHAR(100)     NOT NULL,
    current_revision    VARCHAR(64)      NULL,
    origin              VARCHAR(32)      NOT NULL,
    is_accepting        TINYINT(1)       NOT NULL DEFAULT 1,
    created_utc         DATETIME(3)      NOT NULL,

    PRIMARY KEY (system_id),
    KEY ix_systems_manufacturer (manufacturer, hardware, board)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;


-- --------------------------------------------------------------------------------------------
-- maintainers - the POOL of people who may approve work on a system.
--
-- A POOL, NOT AN OWNER. Phase 6 is explicit: any maintainer of a system may approve work on it,
-- there is no core maintainer who outranks the others, and there is no ownership to transfer.
-- That is what keeps a system moving when one person goes quiet. The composite primary key says
-- exactly this - a set of (system, account) pairs with no rank column to argue about.
--
-- Administrators are NOT listed here; see accounts.is_administrator.
--
-- granted_by names the account that granted it, so the audit question "who made this person a
-- maintainer" is answerable from this row alone. It is nullable only for the first administrator,
-- who is granted by nobody.
-- --------------------------------------------------------------------------------------------
CREATE TABLE maintainers (
    system_id       VARCHAR(255)     NOT NULL,
    account_id      BIGINT UNSIGNED  NOT NULL,
    granted_by      BIGINT UNSIGNED  NULL,
    granted_utc     DATETIME(3)      NOT NULL,

    PRIMARY KEY (system_id, account_id),
    KEY ix_maintainers_account (account_id),

    CONSTRAINT fk_maintainers_system
        FOREIGN KEY (system_id) REFERENCES systems (system_id)
        ON DELETE CASCADE,

    CONSTRAINT fk_maintainers_account
        FOREIGN KEY (account_id) REFERENCES accounts (id)
        ON DELETE CASCADE,

    -- The grantor's account being deleted must not delete the grant, and must not be allowed to
    -- silently rewrite who granted it either - so RESTRICT would block the delete. SET NULL keeps
    -- the grant and loses only the attribution, which the audit table still records permanently.
    CONSTRAINT fk_maintainers_granted_by
        FOREIGN KEY (granted_by) REFERENCES accounts (id)
        ON DELETE SET NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;


-- --------------------------------------------------------------------------------------------
-- submissions - one row per contribution offered for review.
--
-- base_revision is the systems.current_revision the contributor started from. Stored rather than
-- looked up, because by review time the system may have moved on, and "was this built against
-- what is published now?" is the question that decides whether it can be merged directly.
--
-- state is a string, not an enum or a lookup table: the set is small, stable and meaningful in a
-- log line. A CHECK constraint keeps it honest.
--
-- The blob itself is NOT here. Submission payloads are files on disk (Phase 4); this table is the
-- queue and the audit spine. Putting multi-megabyte board data in the database would make every
-- listing query drag it along.
-- --------------------------------------------------------------------------------------------
CREATE TABLE submissions (
    id                  BIGINT UNSIGNED  NOT NULL AUTO_INCREMENT,
    system_id           VARCHAR(255)     NOT NULL,
    account_id          BIGINT UNSIGNED  NULL,
    base_revision       VARCHAR(64)      NULL,
    state               VARCHAR(32)      NOT NULL,
    payload_path        VARCHAR(500)     NULL,
    summary             VARCHAR(500)     NULL,
    created_utc         DATETIME(3)      NOT NULL,
    decided_utc         DATETIME(3)      NULL,
    decided_by          BIGINT UNSIGNED  NULL,
    decision_comment    TEXT             NULL,

    PRIMARY KEY (id),
    KEY ix_submissions_state (state, created_utc),
    KEY ix_submissions_system (system_id, state),
    KEY ix_submissions_account (account_id, created_utc),

    CONSTRAINT fk_submissions_system
        FOREIGN KEY (system_id) REFERENCES systems (system_id)
        ON DELETE CASCADE,

    -- SET NULL, not CASCADE: deleting an account must not erase the record that a submission was
    -- made and decided. The audit trail outlives the account.
    CONSTRAINT fk_submissions_account
        FOREIGN KEY (account_id) REFERENCES accounts (id)
        ON DELETE SET NULL,

    CONSTRAINT fk_submissions_decided_by
        FOREIGN KEY (decided_by) REFERENCES accounts (id)
        ON DELETE SET NULL,

    CONSTRAINT ck_submissions_state
        CHECK (state IN ('pending', 'changes_requested', 'approved', 'rejected', 'withdrawn', 'merged'))
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;


-- --------------------------------------------------------------------------------------------
-- audit - append-only record of who did what.
--
-- NOTHING EVER UPDATES OR DELETES A ROW HERE. There is no code path that should, and if one
-- appears it is a defect. Phase 6's definition of done requires every decision to be attributable,
-- and a mutable audit trail attributes nothing.
--
-- actor_account_id is nullable and actor_label carries a copy of the identity as text, because
-- two kinds of actor have no account row: the system itself (automated rejections, token
-- cleanups) and a deleted account. The label is what a human reads months later, when the account
-- it referred to may be gone.
--
-- detail is JSON-shaped free text rather than columns. What is worth recording differs per action
-- and would otherwise mean a wide table of mostly-null columns, or a new migration per action.
-- --------------------------------------------------------------------------------------------
CREATE TABLE audit (
    id                  BIGINT UNSIGNED  NOT NULL AUTO_INCREMENT,
    actor_account_id    BIGINT UNSIGNED  NULL,
    actor_label         VARCHAR(255)     NOT NULL,
    action              VARCHAR(64)      NOT NULL,
    subject             VARCHAR(255)     NULL,
    detail              TEXT             NULL,
    at_utc              DATETIME(3)      NOT NULL,

    PRIMARY KEY (id),
    KEY ix_audit_at (at_utc),
    KEY ix_audit_actor (actor_account_id, at_utc),
    KEY ix_audit_subject (subject, at_utc),

    CONSTRAINT fk_audit_actor
        FOREIGN KEY (actor_account_id) REFERENCES accounts (id)
        ON DELETE SET NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;
