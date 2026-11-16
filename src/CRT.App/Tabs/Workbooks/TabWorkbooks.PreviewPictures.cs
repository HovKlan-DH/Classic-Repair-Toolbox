using Handlers.DataHandling;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace CRT
{
    // ###########################################################################################
    // THE BOARD PANE'S PICTURES: read through SchematicPreviewPictures, held by the pane while it draws them, and a read
    // that came to nothing not asked again by the pass its own end started (code review, 2026-11-16 - moved out of
    // TabWorkbooks.BoardPreviews.cs, which had grown past the ~1,500 lines a part may have). The pane itself is
    // TabWorkbooks.BoardPreviews.cs; see TabWorkbooks.axaml.cs for the file map.
    // ###########################################################################################
    public partial class TabWorkbooks
    {
        // ###########################################################################################
        // THE PANE'S PICTURES ARE SchematicPreviewPictures' (2026-11-16, performance walkthrough): each decoded at most
        // 1600 wide, ONCE while its file stays as it is, held across rebuilds, tab switches and boards, and never
        // disposed by hand - so neither a rebuild nor the editor a pill opens (ShowDialog does not block the
        // dispatcher, so RefreshWorklogBar can re-enter under it) can draw a disposed bitmap, the render-thread crash
        // this tab once had. Each comes with the picture's own size, which every worklog area is mapped by.
        //
        // While the tab is on screen a picture not read yet is read on the pool and the pane drawn again once it is in
        // (the pane says it is reading them meanwhile); a tab not on screen - a test's - waits for it, as nothing else
        // would draw the pane again. Main reaches this tab only while it is shown (RefreshWorkbooksIfShown).
        // ###########################################################################################
        private readonly HashSet<string> thisPreviewReadsAwaited = new(StringComparer.OrdinalIgnoreCase);

        // Whether the last pass left a picture to be read - the pane says so rather than "could not be found".
        private bool thisPreviewsBeingRead;

        private const string ReadingPicturesMessage = "Reading the schematic images...";

        // ###########################################################################################
        // The picture at `fullPath` as the pane shows it, or null: not readable (one bad path drops that schematic
        // rather than the whole pane), or - on screen - still being read, the pane drawn again once it is.
        // ###########################################################################################
        private SchematicPreviewPicture? GetOrDecodeSchematicPicture(string fullPath)
        {
            // A read that came to nothing for the reads now ending is not asked again by the pass they start (below).
            if (this.thisPreviewReadsFailed.Contains(fullPath))
                return null;

            string? stamp = SchematicPreviewPictures.StampOf(fullPath);

            if (stamp is not null && this.thisPanePictures.TryGetValue(fullPath, out (string Stamp, SchematicPreviewPicture Picture) held) &&
                string.Equals(held.Stamp, stamp, StringComparison.Ordinal))
            {
                this.thisPanePicturesThisPass[fullPath] = held;
                return held.Picture;
            }

            Task<SchematicPreviewPicture?> read = SchematicPreviewPictures.Read(fullPath);

            if (read.IsCompleted)
                return this.HoldForPane(fullPath, stamp, read.Result, this.thisPanePicturesThisPass);

            if (this.VisualRoot is null)
                return this.HoldForPane(fullPath, stamp, read.GetAwaiter().GetResult(), this.thisPanePicturesThisPass);

            this.thisPreviewsBeingRead = true;

            if (this.thisPreviewReadsAwaited.Add(fullPath))
            {
                read.ContinueWith(
                    _ => Avalonia.Threading.Dispatcher.UIThread.Post(() =>
                    {
                        this.thisPreviewReadsAwaited.Remove(fullPath);

                        // Held for the pass this starts, whatever SchematicPreviewPictures has let go of by then - or, come to
                        // nothing, not asked again by it.
                        if (read.IsCompletedSuccessfully && read.Result is not null)
                            this.HoldForPane(fullPath, stamp, read.Result, this.thisPanePictures);
                        else
                            this.thisPreviewReadsFailed.Add(fullPath);

                        // Only once every picture asked for is in: one redraw, not one per picture.
                        if (this.thisPreviewReadsAwaited.Count == 0 && this.VisualRoot is not null)
                        {
                            this.thisPassFromReadsEnding = true;

                            try
                            {
                                this.StartFreshBoardPass();
                            }
                            finally
                            {
                                this.thisPassFromReadsEnding = false;
                            }
                        }
                    }),
                    TaskScheduler.Default);
            }

            return null;
        }

        // ###########################################################################################
        // *** THE PANE HOLDS THE PICTURES IT SHOWS (code review, 2026-11-16). *** SchematicPreviewPictures keeps at most Kept;
        // a pane of more pictures than that read each one over another (the cache letting go of the oldest), and every batch
        // of reads ending began another pass that read them again - for ever, the pane saying "Reading..." while the CPU
        // decoded. So a picture read for the pane is held by it, with the stamp it was read at, for as long as the pane draws
        // it: what one pass drew is what the next finds (thisPanePictures, replaced at the end of each pass by what it drew).
        // A picture that could not be read is not held: the next pass asks again.
        // ###########################################################################################
        private Dictionary<string, (string Stamp, SchematicPreviewPicture Picture)> thisPanePictures = new(StringComparer.OrdinalIgnoreCase);

        private Dictionary<string, (string Stamp, SchematicPreviewPicture Picture)> thisPanePicturesThisPass = new(StringComparer.OrdinalIgnoreCase);

        private SchematicPreviewPicture? HoldForPane(
            string fullPath,
            string? stamp,
            SchematicPreviewPicture? picture,
            Dictionary<string, (string Stamp, SchematicPreviewPicture Picture)> held)
        {
            if (picture is not null && stamp is not null)
                held[fullPath] = (stamp, picture);

            return picture;
        }

        // ###########################################################################################
        // *** A PICTURE THAT CAME TO NOTHING IS NOT ASKED AGAIN BY THE PASS ITS READ STARTED (code review, 2026-11-16). *** Its
        // read is not kept (SchematicPreviewPictures), so that pass asked again, whose read ended, which started another pass -
        // for ever: the pane read "Reading the schematic images..." and the CPU decoded a damaged picture without end, a
        // warning logged at every round. The pictures that came to nothing are left out until a pass something else asks for
        // (a refresh, a board, the tab shown again) - which asks for each once more.
        // ###########################################################################################
        private readonly HashSet<string> thisPreviewReadsFailed = new(StringComparer.OrdinalIgnoreCase);

        // True while the pass the reads ending start runs.
        private bool thisPassFromReadsEnding;

        // At the end of a pass: the pane holds what it drew, and lets go of the rest.
        private void KeepPanePicturesOfThisPass()
        {
            this.thisPanePictures = this.thisPanePicturesThisPass;
            this.thisPanePicturesThisPass = new Dictionary<string, (string Stamp, SchematicPreviewPicture Picture)>(StringComparer.OrdinalIgnoreCase);
        }
    }
}
