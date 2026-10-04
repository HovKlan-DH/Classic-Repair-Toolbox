using System;
using System.IO;
using System.Linq;
using Handlers.DataHandling;
using OfficeOpenXml;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// *** NO WORKBOOK FORCES A GARBAGE COLLECTION (2026-10-03). *** EPPlus calls GC.Collect() when a
// package is disposed unless Settings.DoGarbageCollectOnDispose is false - true by default. In
// CRT that was a full, blocking collection on the UI thread for every board read or written; in
// the CRT.App test process, ~0.4 s per workbook once the heap had grown, and most of a seven-minute
// run (found with dotnet-counters and a GC trace). EpplusLicense makes every package with it off;
// these hold that, and that no production code makes a package any other way.
// ###########################################################################################
public sealed class EpplusPackagesTests
{
    [Fact]
    public void Every_package_EpplusLicense_makes_leaves_garbage_collection_alone()
    {
        using ExcelPackage created = EpplusLicense.NewPackage();
        Assert.False(created.Settings.DoGarbageCollectOnDispose);

        created.Workbook.Worksheets.Add("Sheet");
        using var stream = new MemoryStream(created.GetAsByteArray());

        using ExcelPackage opened = EpplusLicense.OpenPackage(stream);
        Assert.False(opened.Settings.DoGarbageCollectOnDispose);
    }

    // A reader or writer that made its own package would bring the collections back without
    // anything failing - only slower.
    [Fact]
    public void No_production_code_makes_an_ExcelPackage_except_through_EpplusLicense()
    {
        string src = EpplusPackagesTests.RepositoryPath("src");

        string[] offenders = Directory.EnumerateFiles(src, "*.cs", SearchOption.AllDirectories)
            .Where(path => !path.Contains($"{Path.DirectorySeparatorChar}bin{Path.DirectorySeparatorChar}", StringComparison.Ordinal))
            .Where(path => !path.Contains($"{Path.DirectorySeparatorChar}obj{Path.DirectorySeparatorChar}", StringComparison.Ordinal))
            .Where(path => !string.Equals(Path.GetFileName(path), "EpplusLicense.cs", StringComparison.Ordinal))
            .Where(path => File.ReadAllText(path).Contains("new ExcelPackage(", StringComparison.Ordinal))
            .Select(path => Path.GetRelativePath(src, path))
            .ToArray();

        Assert.Empty(offenders);
    }

    // Walks up from the test binary to the repository root (the folder holding the solution).
    private static string RepositoryPath(string relative)
    {
        string? folder = AppContext.BaseDirectory;

        while (folder is not null && !File.Exists(Path.Combine(folder, "Classic-Repair-Toolbox.slnx")))
            folder = Path.GetDirectoryName(folder);

        Assert.NotNull(folder);

        return Path.Combine(folder!, relative);
    }
}
