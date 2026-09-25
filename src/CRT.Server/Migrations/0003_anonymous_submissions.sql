-- ############################################################################################
-- Contributing requires no account (maintainer's decision, 2026-09-21).
--
-- 0001 and 0002 were written assuming a submission always belonged to a verified account. It does
-- not: anyone may contribute, giving only an email address so they can be told whether their work
-- was accepted. An account appears only when somebody is made a MAINTAINER of a system, which the
-- maintainer grants rather than the contributor signing up for.
--
-- See NewContributeStrategy.md, "CONTRIBUTING NEEDS NO ACCOUNT", for why - in short, the threat
-- model never claimed accounts defended against bad submissions, and an account that exists only
-- because somebody once fixed a typo is a credential to protect in exchange for nothing.
--
-- READ 0001_initial.sql's header before editing this. Once applied, this file is history.
-- ############################################################################################


-- --------------------------------------------------------------------------------------------
-- The contributor's contact address.
--
-- REQUIRED IN PRACTICE, NULLABLE IN THE SCHEMA. A submission made by a signed-in maintainer has
-- an account_id and needs no separate address - the account already carries one - so the column
-- cannot be NOT NULL. The application requires one or the other; the database cannot express
-- "exactly one of these two" without a CHECK that would also have to know which states are
-- legitimate, and a constraint that encodes application rules is one that fights the next change.
--
-- THIS IS CONTACT INFORMATION, NOT A CREDENTIAL. Nothing signs in with it, there is no password
-- beside it, and it is never used to identify a returning contributor - two submissions from the
-- same address are two unrelated submissions. It exists so a reviewer can say "accepted" or
-- "here is what needs changing".
--
-- Stored as given rather than normalised, because it is only ever read by a human and mailed to.
-- The normalisation in AccountRules exists to make a unique index behave; there is no unique index
-- here and deliberately so - requiring uniqueness would turn this into an identity.
-- --------------------------------------------------------------------------------------------
ALTER TABLE submissions
    ADD COLUMN contact_email VARCHAR(320) NULL AFTER account_id;

-- Finding a contributor's submissions to tell them the outcome. Not unique - see above.
CREATE INDEX ix_submissions_contact ON submissions (contact_email, created_utc);


-- --------------------------------------------------------------------------------------------
-- Where an anonymous submission came from, for rate limiting.
--
-- WITH NO ACCOUNT THERE IS NOTHING ELSE TO LIMIT AGAINST. Per-account limiting was going to carry
-- this load; without accounts it falls to the address, which is weaker - one address can be a
-- household, an office or a whole country behind CGNAT, and an attacker can rotate addresses.
--
-- That is a real and acknowledged cost of allowing anonymous submission (see the strategy
-- document). It is why the blob size cap and the 24-hour abandoned-upload sweep matter more here
-- than they would have done: they bound what one submission can cost regardless of who sent it.
-- --------------------------------------------------------------------------------------------
ALTER TABLE submissions
    ADD COLUMN created_ip VARCHAR(45) NULL AFTER contact_email;

CREATE INDEX ix_submissions_ip ON submissions (created_ip, created_utc);


-- --------------------------------------------------------------------------------------------
-- The capability token that proves ownership of a submission in flight.
--
-- WITH NO ACCOUNT, A SUBMISSION ID CANNOT BE WHAT AUTHORISES ANYTHING. It is a small consecutive
-- integer, so anyone could upload into - or finalise - a stranger's submission by guessing one.
-- Creating a submission therefore returns a random 256-bit token that every later call for it
-- must present.
--
-- ONLY THE HASH IS STORED, the same rule as every other token in this service (see SecureToken):
-- a database read must not hand somebody the ability to act on submissions that are in flight.
--
-- NOT NULL with no default: every submission has one, and a row without one would be a submission
-- nobody could ever finish. Existing rows cannot exist yet - 0002 shipped in the same sitting and
-- no submission has been made - so a plain NOT NULL is safe here. If that ever stops being true,
-- this needs a default and a backfill instead.
-- --------------------------------------------------------------------------------------------
ALTER TABLE submissions
    ADD COLUMN upload_token_hash CHAR(64) NOT NULL AFTER created_ip;

CREATE INDEX ix_submissions_upload_token ON submissions (upload_token_hash);
