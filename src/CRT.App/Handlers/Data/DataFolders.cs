using CRT;
using System;
using System.Collections.Generic;
using System.IO;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // CRT 4 KEEPS USING THE "Data" FOLDER (owner request, 2026-11-22: "once the users "Data" folder
    // has been migrated to the new version 4.0.0 version, that we then move all data back from the
    // "Data-v4" to the "Data" folder, so we keep using that folder").
    //
    // In CRT's own app data folder, at every launch with no "--data-root=" (Arrange):
    //
    //   - "Data" holding CRT 3's Excel tree (IsCrt3Data) is SET ASIDE as "Data CRT 3" - a rename, so
    //     nothing is copied and nothing is lost. CRT 3 is not installed any more: the update replaced it.
    //   - "Data v4", which the 4.0 pre-releases used, then takes the name "Data"; with neither there,
    //     DataManager seeds a new "Data" from the install package, as for any new data folder.
    //   - A folder that could not be renamed - a file in it held open by another program - is tried
    //     again at the next launch, and this one runs as the pre-releases did: "Data v4", with "Data"
    //     read as CRT 3's folder. So does a CRT 3 "Data" while "Data CRT 3" is there already (CRT 3
    //     installed again in between).
    //
    // CRT 3's folder is READ while what is in it is taken over - its drafts, "_UserContribution"
    // boards and own oscilloscope models (LocalDataMigration, OwnOscilloscopeFile) and the captures
    // 3.0.0 saved into it (OscilloscopeCaptureRelocation) - and DELETED once nothing in it is needed
    // any more (ReasonsToKeep; until then it stays, and the log says why). Only "Data CRT 3" is ever
    // deleted, and only by a launch on the default data and drafts folders (DataManager): a "Data" that
    // could not be renamed may be CRT 4's next time, and another drafts folder says nothing of the
    // drafts CRT 3 left.
    //
    // A delete renames the folder to "Data CRT 3 (deleting)" first and deletes it off the UI thread:
    // one stopped half-way is finished at the next launch (Layout.LeftToDelete), and never read as
    // CRT 3's folder in the meantime.
    // ###########################################################################################
    internal static class DataFolders
    {
        internal const string DeletingSuffix = " (deleting)";

        // ###########################################################################################
        // How a launch found the folders: where CRT's data is this run; CRT 3's folder to take over
        // from ("" for none); whether that folder may be deleted once taken over (it is the set-aside
        // "Data CRT 3"); whether this very launch set it aside; and a folder an earlier delete left
        // half-way, to finish ("" for none).
        // ###########################################################################################
        internal sealed record Layout(string DataRoot, string Crt3Root, bool Crt3MayGo, bool Crt3SetAsideNow, string LeftToDelete);

        // ###########################################################################################
        // Puts the folders in `appFolder` in order - see the header. `move` is Directory.Move but in a
        // test, which hands one that refuses: a file held open stops a rename on Windows alone.
        // ###########################################################################################
        internal static Layout Arrange(string appFolder, Action<string, string>? move = null)
        {
            move ??= Directory.Move;

            string data = Path.Combine(appFolder, AppConfig.DataFolderName);
            string preRelease = Path.Combine(appFolder, AppConfig.PreReleaseDataFolderName);
            string aside = Path.Combine(appFolder, AppConfig.Crt3DataFolderName);
            string leftToDelete = Directory.Exists(aside + DataFolders.DeletingSuffix) ? aside + DataFolders.DeletingSuffix : string.Empty;

            bool setAsideNow = DataFolders.IsCrt3Data(data) &&
                               !Directory.Exists(aside) &&
                               DataFolders.TryMove(move, data, aside, "CRT 3's data folder set aside while what is in it is taken over");

            if (!Directory.Exists(data) &&
                Directory.Exists(preRelease) &&
                !DataFolders.TryMove(move, preRelease, data, "The CRT 4 pre-releases' data folder is CRT's data folder again"))
            {
                return new Layout(preRelease, Directory.Exists(aside) ? aside : string.Empty, Directory.Exists(aside), setAsideNow, leftToDelete);
            }

            if (DataFolders.IsCrt3Data(data))
            {
                Logger.Info($"CRT 3's data folder [{data}] could not be set aside this time - CRT's data is in [{preRelease}] for this run");
                return new Layout(preRelease, data, false, false, leftToDelete);
            }

            return new Layout(data, Directory.Exists(aside) ? aside : string.Empty, Directory.Exists(aside), setAsideNow, leftToDelete);
        }

        // ###########################################################################################
        // CRT 3's folder a "--data-root=" run reads what CRT 3 left from - and never moves or deletes:
        // "Data CRT 3" when there, else "Data" while it holds CRT 3's tree, else "".
        // ###########################################################################################
        internal static string Crt3RootToRead(string appFolder)
        {
            string aside = Path.Combine(appFolder, AppConfig.Crt3DataFolderName);
            string data = Path.Combine(appFolder, AppConfig.DataFolderName);

            return Directory.Exists(aside) ? aside : DataFolders.IsCrt3Data(data) ? data : string.Empty;
        }

        // ###########################################################################################
        // Whether `folder` holds CRT 3's Excel tree: a main workbook ("Classic-Repair-Toolbox....xlsx")
        // at its top and no CRT 4 main file ("Classic-Repair-Toolbox.v4.0.0.json"). A CRT 4 folder CRT 3
        // was run on again holds both, and stays CRT 4's. A folder that is not there, or cannot be
        // listed, is not CRT 3's - nothing is moved on a guess.
        // ###########################################################################################
        internal static bool IsCrt3Data(string folder)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(folder) || !Directory.Exists(folder))
                    return false;

                bool workbook = false;

                foreach (string path in Directory.EnumerateFiles(folder, AppConfig.AppFolderName + "*"))
                {
                    string name = Path.GetFileName(path);

                    if (name.StartsWith(AppConfig.MainExcelFileNamePrefix, StringComparison.OrdinalIgnoreCase) &&
                        name.EndsWith(".json", StringComparison.OrdinalIgnoreCase))
                    {
                        return false;
                    }

                    workbook |= name.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase);
                }

                return workbook;
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                Logger.Warning($"The data folder [{folder}] could not be looked at - [{ex.Message}]");
                return false;
            }
        }

        // ###########################################################################################
        // The capture folder to use once CRT 3's "Data" (`data`) is set aside as `aside` - null when it
        // stays as it is. A folder the user chose INSIDE CRT 3's "Data" went with it, and the captures
        // in it are there: without this the setting would name a folder inside CRT 4's new "Data".
        // (ReasonsToKeep then keeps "Data CRT 3" for good - the user's captures are in it.)
        // ###########################################################################################
        internal static string? CaptureFolderAfterSettingAside(string captureFolder, string data, string aside)
        {
            if (string.IsNullOrWhiteSpace(captureFolder) || !OscilloscopeCaptureRelocation.IsSameOrInside(captureFolder, data))
                return null;

            string relative = Path.GetRelativePath(Path.GetFullPath(data), Path.GetFullPath(captureFolder));

            return relative == "." ? aside : Path.Combine(aside, relative);
        }

        // ###########################################################################################
        // Why CRT 3's folder `crt3Root` must stay - one sentence each, empty when it may go: the
        // oscilloscope models a user added in CRT 3 not carried over yet (their file, `ownModelsPath`,
        // is made once they are); the capture folder inside it; anything of CRT 3's drafts and boards
        // still to take over (LocalDataMigration.StillToTakeOver); and a picture at its top a capture
        // could not be moved out from, or one that may be the user's own (IsImageInDoubt) - CRT 3's
        // data had no picture there. Captures are moved out before this is asked (DataManager).
        // ###########################################################################################
        internal static IReadOnlyList<string> ReasonsToKeep(
            string crt3Root,
            string draftsRoot,
            string ownModelsPath,
            IReadOnlyCollection<string> publishedBoardIds,
            string captureFolder)
        {
            var reasons = new List<string>();

            if (string.IsNullOrWhiteSpace(ownModelsPath) || !File.Exists(ownModelsPath))
                reasons.Add("the oscilloscope models you added in CRT 3 are not carried over yet");

            if (!string.IsNullOrWhiteSpace(captureFolder) && OscilloscopeCaptureRelocation.IsSameOrInside(captureFolder, crt3Root))
                reasons.Add($"oscilloscope captures are saved in [{captureFolder}], inside it");

            reasons.AddRange(LocalDataMigration.StillToTakeOver(draftsRoot, crt3Root, publishedBoardIds));

            try
            {
                foreach (string path in Directory.EnumerateFiles(crt3Root))
                {
                    string name = Path.GetFileName(path);

                    if (OscilloscopeCaptureRelocation.LooksLikeCapture(name) || OscilloscopeCaptureRelocation.IsImageInDoubt(name))
                        reasons.Add($"the picture [{name}] in it may be one of your own");
                }
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                reasons.Add($"it could not be looked through ({ex.Message})");
            }

            return reasons;
        }

        // ###########################################################################################
        // Renames CRT 3's folder to "... (deleting)" - the path to delete, or null (logged) when it
        // could not be renamed: then it stays as it is, and the next launch tries again.
        // ###########################################################################################
        internal static string? SetAsideForDeleting(string crt3Root, Action<string, string>? move = null)
        {
            string deleting = crt3Root + DataFolders.DeletingSuffix;

            return DataFolders.TryMove(move ?? Directory.Move, crt3Root, deleting, "CRT 3's data folder taken over - it is deleted")
                ? deleting
                : null;
        }

        // Deletes `folder` and all in it. Never throws: what could not be deleted is finished at the next launch.
        internal static void Delete(string folder)
        {
            try
            {
                Directory.Delete(folder, recursive: true);
                Logger.Info($"Deleted [{folder}]");
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                Logger.Warning($"[{folder}] could not be deleted all the way - [{ex.Message}] - the rest goes at the next launch");
            }
        }

        private static bool TryMove(Action<string, string> move, string from, string to, string what)
        {
            try
            {
                move(from, to);
                Logger.Info($"{what}: [{from}] -> [{to}]");
                return true;
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                Logger.Warning($"{what} could not be done - [{from}] -> [{to}] [{ex.Message}] - tried again at the next launch");
                return false;
            }
        }
    }
}
