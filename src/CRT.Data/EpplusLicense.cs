using System.IO;
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

        // ###########################################################################################
        // *** EVERY ExcelPackage IS MADE HERE, WITH EPPLUS'S GC.Collect SWITCHED OFF (2026-10-03). ***
        //
        // EPPlus calls GC.Collect() when a package is DISPOSED, unless its
        // Settings.DoGarbageCollectOnDispose is false - and it is true by default. That is a full,
        // blocking collection of the whole process for every workbook read or written, and it costs
        // in proportion to everything else alive in the process. In CRT that is the board images,
        // on the UI thread, several times per draft on every Drafts tab refresh. In the CRT.App test
        // process, whose heap reaches about 500 MB as the run goes on, it was ~0.4 s per workbook
        // and some 85 % of a seven-minute run (measured with dotnet-counters and a GC trace: one
        // induced gen2 collection a second, every one from ExcelPackage.Dispose). Nothing here needs
        // it: the package's memory is reclaimed like any other object's.
        //
        // EpplusPackagesTests fails on any "new ExcelPackage(" in the source outside this file, so
        // a new reader or writer cannot quietly bring the collections back.
        // ###########################################################################################
        public static ExcelPackage NewPackage()
        {
            EpplusLicense.Ensure();
            return EpplusLicense.WithoutCollectOnDispose(new ExcelPackage());
        }

        public static ExcelPackage OpenPackage(Stream stream)
        {
            EpplusLicense.Ensure();
            return EpplusLicense.WithoutCollectOnDispose(new ExcelPackage(stream));
        }

        public static ExcelPackage OpenPackage(FileInfo file)
        {
            EpplusLicense.Ensure();
            return EpplusLicense.WithoutCollectOnDispose(new ExcelPackage(file));
        }

        private static ExcelPackage WithoutCollectOnDispose(ExcelPackage package)
        {
            package.Settings.DoGarbageCollectOnDispose = false;
            return package;
        }
    }
}
