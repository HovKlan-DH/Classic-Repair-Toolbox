-- ############################################################################################
-- OSCILLOSCOPE MODELS SENT IN FROM CRT (CRT 4.0 step 5, 2026-10-29 - the plan: "Submit this model
-- sends it to the administrator's queue on the Maintainer tab, which publishes it into the main file
-- for everybody. This is a new kind of submission on the server").
--
-- A TABLE OF ITS OWN, not a row of `submissions`: a board submission is a board's rows and files,
-- uploaded in parts, reviewed by that board's maintainers and published in two stages. A model is
-- one entry of the main file, sent in one request, decided by the administrator alone and written
-- into the BETA and the stable main files at once. Nothing of the board pipeline applies to it.
--
-- One row per model sent in:
--
--   vendor, series_or_model - its name, as sent (the main file's "brand" and "seriesOrModel").
--   values_json             - every other value, keyed by the main file's names for them
--                             (OscilloscopeModelValues), as JSON.
--   contact_email           - what the contributor typed; NULL when sent signed in (the account's
--                             address is used then, as for a board submission).
--   contact_name            - the name the contributor gave, if any.
--   account_id              - the account it was sent with, if signed in.
--   message                 - what the contributor wrote to the administrator.
--   full_test_suite_passed  - CRT's word that "Run full test suite" went through on the model sent.
--   client_version          - the CRT that sent it (User-Agent), for the administrator to see.
--   sender_address          - the client's internet connection, for the daily limit: its IPv4 address,
--                             or the /64 network of its IPv6 address; NULL when unknown.
--   state                   - 'pending', 'in_beta' (a publish wrote it into BETA's main file and the
--                             stable source could not take it - Publish again finishes it), 'published',
--                             'rejected', or 'replaced' (a newer one of the same name from the same
--                             contributor took its place while it waited, or another model of its name
--                             was written into BETA after it).
--   decided_utc, decided_by_account_id, decision_note - the administrator's decision, and the
--                             reason told the contributor of a rejection.
--
-- READ 0001_initial.sql's header before editing this. Once applied, this file is history.
-- ############################################################################################

CREATE TABLE oscilloscope_submissions (
    id                      BIGINT UNSIGNED  NOT NULL AUTO_INCREMENT,
    created_utc             DATETIME(3)      NOT NULL,
    vendor                  VARCHAR(100)     NOT NULL,
    series_or_model         VARCHAR(100)     NOT NULL,
    values_json             MEDIUMTEXT       NOT NULL,
    contact_email           VARCHAR(320)     NULL,
    contact_name            VARCHAR(200)     NULL,
    account_id              BIGINT UNSIGNED  NULL,
    message                 TEXT             NULL,
    full_test_suite_passed  TINYINT(1)       NOT NULL DEFAULT 0,
    client_version          VARCHAR(64)      NULL,
    sender_address          VARCHAR(64)      NULL,
    state                   VARCHAR(16)      NOT NULL DEFAULT 'pending',
    decided_utc             DATETIME(3)      NULL,
    decided_by_account_id   BIGINT UNSIGNED  NULL,
    decision_note           TEXT             NULL,

    PRIMARY KEY (id),
    KEY ix_oscilloscope_submissions_state (state, created_utc),
    KEY ix_oscilloscope_submissions_sender (sender_address, created_utc)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;
