-- ############################################################################################
-- A MAINTAINER CHANGING THEIR OWN EMAIL ADDRESS (owner request, 2026-10-03).
--
-- The address is the account's identity (0001_initial.sql), so a new one is proved before it is
-- used: the maintainer asks for the change in the Maintainer tab's "Your account" window, the
-- server mails a one-time code to the NEW address, and the address changes only when that code
-- is pasted back (AccountSelfServiceFlows).
--
-- The code is one more one-shot token of account_tokens - purpose 'email_change' - with the same
-- hash-only storage, expiry and consumption as verification and password reset. What it needs
-- that they do not is the address it is FOR, kept until the code is used:
--
--   pending_email            - the new address as the maintainer typed it, which mail is sent to.
--   pending_email_normalised - the same, normalised (AccountRules.NormaliseEmail), which is what
--                              accounts.email_normalised is unique on.
--
-- NULL on every other kind of token. Columns rather than a table of their own: account_tokens is
-- read by NAMED columns (MySqlAccountStore.FindTokenByHashAsync), not by position, so adding two
-- moves nothing that is already read.
--
-- READ 0001_initial.sql's header before editing this. Once applied, this file is history.
-- ############################################################################################

ALTER TABLE account_tokens
    ADD COLUMN pending_email            VARCHAR(320)  NULL AFTER purpose,
    ADD COLUMN pending_email_normalised VARCHAR(320)  NULL AFTER pending_email;
