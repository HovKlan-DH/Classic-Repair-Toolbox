-- ############################################################################################
-- INVITING A NEW MAINTAINER BY EMAIL (owner request, 2026-09-27): "I should be able to either
-- select an existing maintainer or invite a new maintainer via email."
--
-- The administrator names an address that has no account yet. The server mails it a one-time code
-- and keeps only the code's SHA-256 here (token_hash), exactly as account_tokens does. Accepting
-- the code in CRT Maintainer creates the account - verified, since the code proves the mailbox -
-- and puts it in the pool of every system that address has an open invitation to.
--
-- A TABLE OF ITS OWN, not a pool row for a half-made account: nothing that reads `maintainers`
-- (the queue, who is mailed, who may approve) can see an invited person before they have accepted.
--
--   system_id           - the system invited to; utf8mb4_bin like every other system_id (0005).
--   email / _normalised - as typed, and as matched (the same normalising accounts use).
--   invited_by          - the administrator; SET NULL if that account is ever deleted, the audit
--                         table keeps the name.
--   expires_utc         - an invitation is good for a limited time (MaintainerInvitationFlows).
--   accepted_utc / _account_id, withdrawn_utc - how it ended; all three NULL while it is open.
--
-- READ 0001_initial.sql's header before editing this. Once applied, this file is history.
-- ############################################################################################
CREATE TABLE maintainer_invitations (
    id                   BIGINT UNSIGNED  NOT NULL AUTO_INCREMENT,
    system_id            VARCHAR(255)     NOT NULL COLLATE utf8mb4_bin,
    email                VARCHAR(320)     NOT NULL,
    email_normalised     VARCHAR(320)     NOT NULL,
    token_hash           CHAR(64)         NOT NULL,
    invited_by           BIGINT UNSIGNED  NULL,
    created_utc          DATETIME(3)      NOT NULL,
    expires_utc          DATETIME(3)      NOT NULL,
    accepted_utc         DATETIME(3)      NULL,
    accepted_account_id  BIGINT UNSIGNED  NULL,
    withdrawn_utc        DATETIME(3)      NULL,

    PRIMARY KEY (id),
    UNIQUE KEY ux_maintainer_invitations_hash (token_hash),
    KEY ix_maintainer_invitations_email (email_normalised),
    KEY ix_maintainer_invitations_system (system_id),

    CONSTRAINT fk_maintainer_invitations_system
        FOREIGN KEY (system_id) REFERENCES systems (system_id)
        ON DELETE CASCADE,

    CONSTRAINT fk_maintainer_invitations_invited_by
        FOREIGN KEY (invited_by) REFERENCES accounts (id)
        ON DELETE SET NULL,

    CONSTRAINT fk_maintainer_invitations_accepted_account
        FOREIGN KEY (accepted_account_id) REFERENCES accounts (id)
        ON DELETE SET NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;
