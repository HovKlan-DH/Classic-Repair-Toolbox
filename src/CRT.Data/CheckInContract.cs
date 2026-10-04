using System.Collections.Generic;
using System.Net.Http;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // *** THE LAUNCH CHECK-IN, AS IT TRAVELS (owner request, 2026-10-03: retire the app-checkin PHP
    // page "so this should now be fully handled by the backend server"). *** One small form CRT
    // posts once per launch: the operating system, its version and the CPU, plus a "control" field
    // that must say "CRT". The CRT VERSION is not in the form - it is the request's User-Agent
    // ("CRT 2026.10.0"), which is how the PHP page always read it. CRT builds the form with
    // BuildForm, the server (CRT.Server's CheckInFormReader) reads it with these names.
    //
    // *** THE FORM IS THE OLD PHP PAGE'S, UNCHANGED, ON PURPOSE. *** Every CRT already installed posts
    // exactly this to https://classic-repair-toolbox.dk/app-checkin/, and Apache forwards that old
    // address to the server's route - so older CRTs keep counting the day the PHP is removed, which
    // only holds while these names, the "CRT" control value and the User-Agent stay as they are.
    // CRT ignores the answer; ThanksAnswer is the PHP page's, kept so the two read alike in a log.
    // ###########################################################################################
    public static class CheckInContract
    {
        // The route, under "/api" - the server maps "/api/" + this, CRT posts to CrtServerBaseUrl + "/" + this.
        public const string PathUnderApi = "usage/check-in";

        public const string ControlField = "control";
        public const string OsHighlevelField = "osHighlevel";
        public const string OsVersionField = "osVersion";
        public const string CpuField = "cpu";

        // What the control field must say - the PHP page stored nothing for anything else.
        public const string ControlValue = "CRT";

        // What the server answers, as the whole body, to a check-in it accepted.
        public const string ThanksAnswer = "Thanks for checking :-)";

        // The form CRT sends - URL-encoded, as it always was.
        public static FormUrlEncodedContent BuildForm(string osHighlevel, string osVersion, string cpu) =>
            new(new List<KeyValuePair<string, string>>
            {
                new(CheckInContract.ControlField, CheckInContract.ControlValue),
                new(CheckInContract.OsHighlevelField, osHighlevel ?? string.Empty),
                new(CheckInContract.OsVersionField, osVersion ?? string.Empty),
                new(CheckInContract.CpuField, cpu ?? string.Empty)
            });
    }
}
