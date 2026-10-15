-- ############################################################################################
-- MAIL THE MAIL SYSTEM REFUSED, KEPT UNTIL IT GOES (code review, 2026-10-13).
--
-- Every mail used to be fire-and-forget: SmtpEmailSender logged a refusal and the message was gone.
-- With postfix down for an hour, every decision mail, every "waiting for your review" and every
-- approval-needed mail of that hour reached nobody, and nothing said so.
--
-- One row per mail that could not be handed to the local mail system. MailOutboxSender tries it
-- again every few minutes, waiting longer after each failure, and gives it up after
-- MailOutboxFlow.GiveUpAfter. Account > "Server version" shows how many rows there are.
--
-- *** A MAIL CARRYING A WORKING CREDENTIAL IS NEVER STORED WITH ITS WORDS. *** Only token HASHES
-- are stored (CLAUDE.md: sessions, verification and reset links, submission tokens) - a database
-- read must not yield a working link or code. A verification, password reset, address-change or
-- invitation mail that fails is recorded WITHOUT its bodies (body_text NULL): never retried, it is
-- counted so the administrator sees the mail system failed, and the person asks for a new one
-- (or the administrator invites again).
--
--   to_address        - where it goes.
--   subject           - its subject.
--   body_text         - the plain-text body; NULL for a credential mail (not retried).
--   body_html         - the HTML body (MailBody), NULL without one.
--   reply_to          - the Reply-To address, NULL without one.
--   created_utc       - when it first failed.
--   attempts          - how many times it has been tried (1 after the first failure).
--   next_attempt_utc  - not tried again before this.
--   last_error        - the last refusal, shortened, for the log reader - never shown to anybody.
--
-- READ 0001_initial.sql's header before editing this. Once applied, this file is history.
-- ############################################################################################

CREATE TABLE mail_outbox (
    id                BIGINT UNSIGNED  NOT NULL AUTO_INCREMENT,
    to_address        VARCHAR(320)     NOT NULL,
    subject           VARCHAR(500)     NOT NULL,
    body_text         MEDIUMTEXT       NULL,
    body_html         MEDIUMTEXT       NULL,
    reply_to          VARCHAR(320)     NULL,
    created_utc       DATETIME(3)      NOT NULL,
    attempts          INT UNSIGNED     NOT NULL DEFAULT 1,
    next_attempt_utc  DATETIME(3)      NOT NULL,
    last_error        VARCHAR(500)     NULL,

    PRIMARY KEY (id),
    KEY ix_mail_outbox_due (next_attempt_utc),
    KEY ix_mail_outbox_created (created_utc)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;
