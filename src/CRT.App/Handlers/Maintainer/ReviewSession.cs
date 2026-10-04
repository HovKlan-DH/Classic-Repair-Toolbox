using System;

namespace Handlers.MaintainerHandling
{
    // ###########################################################################################
    // The maintainer's signed-in session (NewContributeStrategy.md Phase 5, task 2).
    //
    // *** THE TOKEN NAMING IS A TRAP, AND THIS TYPE EXISTS PARTLY TO DEFUSE IT. *** The login
    // endpoint answers with a field called `refreshToken`, which reads as "the token you exchange
    // for a real one". It is not: AccountFlows.AuthenticateAsync takes THAT value and looks it up
    // as the session token, so it is what goes straight into `Authorization: Bearer`. A client
    // author who trusts the name will go looking for an access-token exchange that does not
    // exist, get a 401 from every call, and blame the server. So the value is named for what it
    // DOES here.
    //
    // *** IT IS A CREDENTIAL, AND IT IS NOW PERSISTED - UNDER CONDITIONS. *** It authorises
    // reading every queued contribution, and - for an administrator - publishing to every user of
    // a system. This header used to say it was never written to disk, and to forbid adding
    // "remember me" without first deciding where that file lives and who can read it. Those
    // questions have since been answered (owner request, 2026-09-22), so the rule has
    // changed rather than been broken:
    //
    //   - WHERE: one file beside CRT's own AppData settings. See ReviewSessionStore.
    //   - WHO CAN READ IT: on Windows, only the logged-in user, because the token is DPAPI-
    //     encrypted to that account - a copy taken to another machine or another user is inert.
    //     On every other platform there is no equivalent, so NOTHING IS STORED AT ALL rather than
    //     a plaintext token pretending to be protected. See ReviewSessionProtection.
    //
    // The reason the old rule gave - "the session is short-lived by design" - is what actually
    // changed. The server now slides a session's expiry forward as it is used, so the choice is no
    // longer between a short session and a long one, but between asking a maintainer for a complex
    // password on a schedule and storing one encrypted value they can revoke by signing out.
    //
    // *** WHAT MUST STILL NEVER HAPPEN: CRT must not call /api/accounts/refresh. *** That
    // endpoint ROTATES the token and poisons the old value immediately, so a crash between the
    // server rotating and CRT saving the replacement revokes EVERY session for the account
    // and writes an audit entry reading as an attack. Extension is done server-side without
    // rotation precisely so a file-backed token is never in that position.
    //
    // Pure and immutable, so the tab holds one of these and nothing else has to know how a
    // session is shaped.
    // ###########################################################################################
    public sealed record ReviewSession(
        string BearerToken,
        DateTimeOffset ExpiresUtc,
        long AccountId,
        string Email,
        string DisplayName)
    {
        // ###########################################################################################
        // Whether this session is still usable at the given moment.
        //
        // TAKES THE TIME RATHER THAN READING THE CLOCK, so expiry is testable without waiting and
        // without a fake clock abstraction - the same reason every flow in CRT.Server takes a
        // DateTimeOffset.
        //
        // The margin matters: a token that expires while a request is in flight fails with a 401
        // the user cannot act on, and they see it as the app logging them out at random. Treating
        // an almost-expired session as already expired turns that into a clean, explainable
        // sign-in.
        // ###########################################################################################
        public bool IsUsableAt(DateTimeOffset now)
        {
            if (string.IsNullOrWhiteSpace(this.BearerToken))
                return false;

            // *** NEITHER SIDE MAY BE SHIFTED BY THE MARGIN, BECAUSE BOTH ENDS OVERFLOW. ***
            // Two bugs were found here in a row, and both threw ArgumentOutOfRangeException:
            //
            //   - `ExpiresUtc - ExpiryMargin > now` throws when ExpiresUtc is MinValue, which is
            //     exactly what ReviewApiParser records for an expiry it could not read. The one
            //     case this method exists to handle safely would have crashed the app.
            //   - `ExpiresUtc > now + ExpiryMargin` moves the overflow to the other end and
            //     throws when `now` is MaxValue.
            //
            // A method that decides whether to keep using a credential must be TOTAL - it is on
            // the path of every request and is called from UI code where an exception is a crash,
            // not a caught error. So the margin is applied to the DIFFERENCE, which is a TimeSpan
            // and cannot overflow a DateTime in either direction.
            return this.ExpiresUtc - now > ReviewSession.ExpiryMargin;
        }

        // One minute: long enough to cover a slow request over a domestic connection, short
        // enough not to throw away a usable session.
        public static readonly TimeSpan ExpiryMargin = TimeSpan.FromMinutes(1);

        // ###########################################################################################
        // The same session - the same token and expiry - carrying the name and address the server
        // now holds (2026-10-03): after a change in the "Your account" window, and when the
        // remembered sign-in is read again at launch. They are what the Feedback tab and the
        // Submit dialog use while signed in, so they must not stay as they were at sign-in.
        //
        // An answer about a DIFFERENT account changes nothing - it cannot be this session's. An
        // account id this session could not read (0, ReviewApiParser.ParseLogin's fallback) takes
        // the answer's, since the answer came back for this very token.
        // ###########################################################################################
        public ReviewSession WithAccount(Handlers.DataHandling.AccountAnswer account)
        {
            ArgumentNullException.ThrowIfNull(account);

            if (this.AccountId != 0 && account.Id != this.AccountId)
                return this;

            return this with { AccountId = account.Id, Email = account.Email, DisplayName = account.DisplayName };
        }
    }
}
