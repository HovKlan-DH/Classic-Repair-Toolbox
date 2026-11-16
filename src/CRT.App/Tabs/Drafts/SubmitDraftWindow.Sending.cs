using Handlers.DataHandling;
using Handlers.MaintainerHandling;
using Handlers.Online;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace CRT
{
    // ###########################################################################################
    // SENDING THE ONE BOARD (SubmitAsync, in the dialog's main part) and what came of it. ONE BOARD PER SUBMIT (owner
    // request, 2026-11-15: "do not allow the user to send multiple submissions in one submit, as this will complicate the
    // "Main contribution" textfield, so only do allow to send ONE contribution/board per submit"): the other drafts were
    // offered to go with it from CRT 4.0 step 2's slice 5 until then, each a submission of its own under the same message.
    // ###########################################################################################
    public partial class SubmitDraftWindow
    {
        // Stands in for sending the board (SubmitAsync, which goes to the server) - a test sends nothing over the network
        // (test rule 6). null is the shipped path.
        internal Func<SubmitDraftTarget, CancellationToken, Task<SubmissionResult>>? SendOverrideForTests { get; set; }

        // ###########################################################################################
        // What came of sending the board: the server's answer (taken or turned down) - or, when it never got there, the
        // sentence ShowFailure says and what the server found.
        // ###########################################################################################
        private sealed record BoardOutcome(
            SubmissionResult? Result,
            string FailureMessage,
            IReadOnlyList<ValidationFinding> Findings);

        // ###########################################################################################
        // Sends the board and says what came of it. Only the contributor's cancel leaves here (code review, 2026-10-13):
        // every other way it can end is an outcome.
        // ###########################################################################################
        private async Task<BoardOutcome> SendOneAsync(SubmitDraftTarget target, IProgress<SubmissionProgress> progress, CancellationToken token)
        {
            try
            {
                SubmissionResult result = this.SendOverrideForTests is { } send
                    ? await send(target, token)
                    : await Task.Run(() => this.SubmitAsync(target, progress, token), token);

                return new BoardOutcome(result, string.Empty, result.Findings ?? []);
            }
            catch (OperationCanceledException) when (token.IsCancellationRequested)
            {
                throw;
            }
            catch (SubmissionRejectedException ex)
            {
                return new BoardOutcome(null, ex.Message, ex.Findings);
            }
            catch (Exception ex) when (ex is System.Net.Http.HttpRequestException or IOException or TaskCanceledException)
            {
                // A request that timed out is a TaskCanceledException too - not the contributor's cancel, which the first catch
                // took - so it reads as "could not be sent".
                return new BoardOutcome(null, SubmitDraftWindow.NoConnectionMessage, []);
            }
            catch (Exception ex)
            {
                // A file the hashing cannot read (UnauthorizedAccessException is not an IOException) was the known way here;
                // anything else is reported the same way and logged in full.
                Logger.Critical($"Submitting the draft of [{target.DisplayName}] failed unexpectedly - [{ex}]", alreadyShown: true);
                return new BoardOutcome(null, SubmitDraftWindow.UnexpectedMessage, []);
            }
        }
    }

    // ###########################################################################################
    // The board to send: what the dialog needs of it - the board as it should read, who it is, where its draft's files lie,
    // its KiCad calibrations, and what the draft held as it was read (kept on the receipt).
    // ###########################################################################################
    internal sealed record SubmitDraftTarget(
        string DisplayName,
        BoardData Board,
        SubmissionIdentity Identity,
        string BoardFolder,
        IReadOnlyList<KiCadCalibrationEntry> Calibrations,
        string Fingerprint);
}
