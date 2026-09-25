using OfficeOpenXml;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // Centralises the EPPlus licence call so it exists once rather than once per read/write
    // method. Every EPPlus entry point in CRT.Data (BoardDataReader.LoadAsync,
    // BoardDataReader.TryCollectReferencedLocalFiles, BoardDataWriter's save path) calls this
    // instead of ExcelPackage.License.SetNonCommercialPersonal(...) directly - Phase 1's original
    // move left the calls duplicated in place (a straight file move, not a rewrite), and this is
    // the follow-up that removes the duplication now that CRT.Data exists to hold it.
    //
    // Calling SetNonCommercialPersonal repeatedly is already how the app behaves today (it is
    // called once per LoadAsync/save invocation, with no guard, across several files including
    // DataManager.cs in CRT.App), so centralising it changes nothing about WHEN or HOW OFTEN it
    // runs - only that the licence holder name is written once.
    //
    // NOTE for the project owner, carried over from the strategy document: the Community licence
    // this uses is chosen for a desktop application used by an individual. Whether it also covers
    // running EPPlus inside a server process (from Phase 3 onward) is a licence question, not a
    // technical one, and must be confirmed before CRT.Server ships anything that reads or writes
    // Excel.
    // ###########################################################################################
    internal static class EpplusLicense
    {
        public static void Ensure() => ExcelPackage.License.SetNonCommercialPersonal("Dennis Helligsø");
    }
}
