-- ############################################################################################
-- THE RATE LIMITER'S OWN TABLE, SO THE AUDIT TRAIL IS APPEND-ONLY AGAIN (code review, 2026-10-13).
--
-- 0001_initial.sql promises the audit table is never updated. Two things broke that: failed
-- sign-ins were counted from `audit` rows and a successful sign-in UPDATEd them to
-- 'login.failed.cleared' to take them out of the count, and every mail-sending request wrote a
-- 'mail.requested' budget row there. A table counted by a limiter is working state, not history.
--
-- One row per counted event:
--
--   kind         - 'login.failed' (AuthRateLimitPolicy's sign-in buckets) or 'mail.requested'
--                  (its mail budget per address).
--   account_key  - the normalised address signed in as; NULL for a mail request (the address a
--                  mail would go to is the victim's in the abuse case and is never recorded).
--   address      - the client's IP address, NULL when it was not known.
--   at_utc       - when.
--
-- A successful sign-in DELETES its account's 'login.failed' rows: they are counters, and the audit
-- trail keeps its own append-only 'login.failed' row for each failure. Rows older than a day are
-- deleted by the hourly sweep (MySqlAccountStore.DeleteRateLimitEventsBeforeAsync) - every window
-- the limiter counts is far shorter.
--
-- The rows already in `audit` are left exactly as they are: history.
--
-- READ 0001_initial.sql's header before editing this. Once applied, this file is history.
-- ############################################################################################

CREATE TABLE rate_limit_events (
    id           BIGINT UNSIGNED  NOT NULL AUTO_INCREMENT,
    kind         VARCHAR(32)      NOT NULL,
    account_key  VARCHAR(255)     NULL,
    address      VARCHAR(64)      NULL,
    at_utc       DATETIME(3)      NOT NULL,

    PRIMARY KEY (id),
    KEY ix_rate_limit_events_account (kind, account_key, at_utc),
    KEY ix_rate_limit_events_address (kind, address, at_utc),
    KEY ix_rate_limit_events_at (at_utc)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;
