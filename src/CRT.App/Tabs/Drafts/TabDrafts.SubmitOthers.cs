using Handlers.DataHandling;
using Handlers.Online;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace CRT
{
    // ###########################################################################################
    // THE OTHER DRAFTS SENT WITH ONE (CRT 4.0 step 2, slice 5 - the plan: "Submit offers to send all
    // the drafts at once, with one message"). The Submit dialog lists every other draft "My drafts"
    // holds (OtherDraftOffers, from the rows as last read), and once Submit is pressed it asks this
    // tab to make each ticked one ready (PrepareOtherDraftAsync) - the same steps a Submit on that
    // board takes before its dialog opens, with one difference: a draft with errors is not sent and
    // says so, since no table can open on another board's draft from inside the dialog.
    // ###########################################################################################
    public partial class TabDrafts
    {
        // Every draft but the one Submit was pressed on, as its row shows it - with why it cannot go.
        internal IReadOnlyList<OtherDraftOffer> OtherDraftOffers(HardwareBoardEntry entry) =>
            this.Drafts
                .Where(row => !string.Equals(row.ExcelDataFile, entry.ExcelDataFile, StringComparison.OrdinalIgnoreCase))
                .Select(row => new OtherDraftOffer(row.Entry, row.DisplayName, row.RowSummary, row.NotOfferedWithOthers))
                .ToList();

        // ###########################################################################################
        // Another draft made ready to send, in the order its own Submit goes: what is held in memory
        // saved into its file, nothing sent that holds nothing or was just sent as it is, the board
        // read as a submit reads it, its errors counted - and the credit typed in the dialog put into
        // it for its own work (CreditSuggestions), as this board's ticked credits are. Or why not.
        // ###########################################################################################
        internal async Task<OtherDraftReady> PrepareOtherDraftAsync(HardwareBoardEntry entry, string creditName, string creditContact)
        {
            DraftStatus? status = DraftStatusReader.Resolve(DataManager.DataRoot, DraftManager.DraftsRoot, entry.ExcelDataFile, entry.IsPublished);

            if (status is null)
            {
                return OtherDraftReady.NotSent(SubmitDraftsWording.DraftGone);
            }

            // No table or pane is open on another board, so only the draft held in memory can hold
            // a change its file lacks - and a submit sends the file.
            DraftSession? session = DraftManager.Sessions.Find(DraftManager.DraftsRoot, entry.ExcelDataFile);

            if (session is { HasUnsavedChanges: true } && await session.SaveAsync() != DraftWorkbookEditOutcome.Saved)
            {
                return OtherDraftReady.NotSent(SubmitDraftsWording.NotSaved);
            }

            if (DraftStatusReader.CountChangesCached(status) == 0 && !DraftStatusReader.CountOtherChangesCached(status, DraftManager.DraftsRoot).Any)
            {
                return OtherDraftReady.NotSent(SubmitDraftsWording.NothingToSend);
            }

            string fingerprint = TabDrafts.FingerprintOf(entry, status);

            SubmissionReceipt? lastSubmission = SubmissionReceiptPresenter.LatestForBoard(
                this.ReceiptsOverrideForTests ?? SubmissionReceiptStore.All,
                BoardDescriptorRules.BoardIdFromExcelDataFile(entry.ExcelDataFile),
                status.CreatedUtc);

            if (SubmissionReceiptPresenter.IsAlreadySent(lastSubmission, fingerprint))
            {
                return OtherDraftReady.NotSent(SubmitDraftsWording.AlreadySent);
            }

            if (await TabDrafts.BoardToSendAsync(status) is not { } board)
            {
                return OtherDraftReady.NotSent(SubmitDraftsWording.NotRead);
            }

            int errors = await Task.Run(() => TabDrafts.CountErrors(entry, board));

            if (errors > 0)
            {
                return OtherDraftReady.NotSent(SubmitDraftsWording.ErrorsFirst(errors));
            }

            List<CreditEntry> credits = CreditSuggestions.ToAdd(
                board,
                CreditSuggestions.For(await this.PublishedBoardForAsync(entry, status), board),
                creditName,
                creditContact);

            if (credits.Count > 0)
            {
                if (await this.AddCreditsToDraftAsync(entry, status, credits) is not { } added)
                {
                    return OtherDraftReady.NotSent(SubmitDraftsWording.CreditNotAdded);
                }

                (board, fingerprint) = added;
            }

            return OtherDraftReady.Ready(new SubmitDraftTarget(
                entry.ToString(),
                board,
                TabDrafts.IdentityFor(entry, status),
                DraftFolderLayout.GetBoardFolder(DraftManager.DraftsRoot, entry.ExcelDataFile),
                TabDrafts.CalibrationsFor(entry, board),
                fingerprint,
                Credited: credits.Count > 0));
        }
    }
}
