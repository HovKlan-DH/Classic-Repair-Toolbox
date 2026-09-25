-- ############################################################################################
-- TWO APPROVALS FOR A SHARED-FILE CHANGE (maintainer decision, 2026-09-25): "in case of changes to
-- any shared file, then both the reviewer and the admin should approve before publishing to BETA
-- or production. If there is no shared files changed, then normal reviewer is sufficient."
--
-- Who must approve is ApprovalRules (CRT.Data). These two tables only remember who HAS, so the
-- first of two approvals survives until the second arrives. One row per role: the second reviewer
-- of a board does not count as the administrator.
--
-- account_label carries a readable copy of who approved, for the same reason audit.actor_label
-- does: the account may later be renamed or removed, and "approved by somebody" helps nobody.
--
-- READ 0001_initial.sql's header before editing this. Once applied, this file is history.
-- ############################################################################################


-- --------------------------------------------------------------------------------------------
-- Approvals of a SUBMISSION (the BETA publish). The first of two moves the submission to
-- 'approved' - a state the schema has allowed since 0001 and nothing wrote until now - and it
-- stays in the review queue until the second one publishes it.
-- --------------------------------------------------------------------------------------------
CREATE TABLE submission_approvals (
    submission_id   BIGINT UNSIGNED  NOT NULL,
    role            VARCHAR(16)      NOT NULL,
    account_id      BIGINT UNSIGNED  NULL,
    account_label   VARCHAR(255)     NOT NULL,
    approved_utc    DATETIME(3)      NOT NULL,

    PRIMARY KEY (submission_id, role),

    CONSTRAINT ck_submission_approvals_role CHECK (role IN ('reviewer', 'administrator')),

    CONSTRAINT fk_submission_approvals_submission
        FOREIGN KEY (submission_id) REFERENCES submissions (id)
        ON DELETE CASCADE,

    CONSTRAINT fk_submission_approvals_account
        FOREIGN KEY (account_id) REFERENCES accounts (id)
        ON DELETE SET NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;


-- --------------------------------------------------------------------------------------------
-- Approvals of publishing a SYSTEM to production, keyed by the BETA content hash they were given
-- against: a publish landing in BETA afterwards changes the hash, and an approval of the earlier
-- state must not carry over to a state nobody looked at.
--
-- system_id is utf8mb4_bin like every other copy of it (migration 0005): a foreign key and the key
-- it references must share a collation.
-- --------------------------------------------------------------------------------------------
CREATE TABLE production_approvals (
    system_id           VARCHAR(255)     NOT NULL COLLATE utf8mb4_bin,
    beta_content_hash   VARCHAR(64)      NOT NULL,
    role                VARCHAR(16)      NOT NULL,
    account_id          BIGINT UNSIGNED  NULL,
    account_label       VARCHAR(255)     NOT NULL,
    approved_utc        DATETIME(3)      NOT NULL,

    PRIMARY KEY (system_id, beta_content_hash, role),

    CONSTRAINT ck_production_approvals_role CHECK (role IN ('reviewer', 'administrator')),

    CONSTRAINT fk_production_approvals_system
        FOREIGN KEY (system_id) REFERENCES systems (system_id)
        ON DELETE CASCADE,

    CONSTRAINT fk_production_approvals_account
        FOREIGN KEY (account_id) REFERENCES accounts (id)
        ON DELETE SET NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;
