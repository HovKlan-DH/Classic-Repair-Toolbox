using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Xml.Linq;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// *** A PINNED PACKAGE IS REMOVED ONCE IT IS NO LONGER NEEDED (owner request, 2026-10-14). ***
//
// A PackageReference marked `PinnedOver="<package>"` exists only to raise a package another one
// brings in: CRT.Data pins System.Security.Cryptography.Pkcs and .Xml at 10.0.13 because EPPlus
// 8.7.1 asks for 10.0.10, which known vulnerabilities cover. Left behind after EPPlus catches up,
// such a line does nothing but hold the version where it was. This reads every marked line in the
// src project files and fails, naming the line to remove, once the package it names requires that
// version or later itself, no longer brings the pinned package in, or is no longer referenced.
//
// What a package requires is read from the restore's own record (obj/project.assets.json) - the
// minimums NuGet resolved against, so the test needs no network and no package cache.
// ###########################################################################################
public sealed class PinnedPackagesTests
{
    // One marked PackageReference, as its project file has it.
    internal sealed record Pin(string Project, string Package, string Version, string Over);

    [Fact]
    public void Every_pinned_package_still_raises_what_the_package_it_names_asks_for()
    {
        var problems = new List<string>();

        foreach (Pin pin in PinnedPackagesTests.Pins())
        {
            string assets = Path.Combine(
                PinnedPackagesTests.RepositoryPath(Path.Combine("src", pin.Project)), "obj", "project.assets.json");

            Assert.True(File.Exists(assets), $"{assets} is missing - restore the solution first.");

            foreach ((string? overVersion, string? overRequires) in PinnedPackagesTests.WhatItRequires(assets, pin))
            {
                string? problem = PinnedPackagesTests.Verdict(pin, overVersion, overRequires);

                if (problem is not null)
                    problems.Add(problem);
            }
        }

        Assert.True(problems.Count == 0, string.Join(Environment.NewLine, problems));
    }

    // The pins this was written for are found - so the test above is not passing on nothing.
    [Fact]
    public void The_marked_lines_are_found_in_the_project_files()
    {
        List<Pin> pins = PinnedPackagesTests.Pins();

        Assert.All(pins, pin => Assert.False(string.IsNullOrWhiteSpace(pin.Version), $"{pin.Package} has no Version"));
        Assert.Contains(pins, pin => pin is { Project: "CRT.Data", Package: "System.Security.Cryptography.Pkcs", Over: "EPPlus" });
    }

    // ###########################################################################################
    // The rule itself, on its own: a pin stays only while the package it names asks for LESS.
    // ###########################################################################################
    [Theory]
    [InlineData("10.0.10", null)]
    [InlineData("10.0.12", null)]
    [InlineData("[10.0.10, )", null)]
    [InlineData("10.0.13-rc.1", null)]
    [InlineData("10.0.13", "itself requires")]
    [InlineData("10.0.15", "itself requires")]
    [InlineData("[10.0.14, 11.0.0)", "itself requires")]
    [InlineData("11.0.0", "itself requires")]
    public void A_pin_stays_only_while_the_package_it_names_asks_for_less(string requires, string? problem)
    {
        var pin = new Pin("CRT.Data", "System.Security.Cryptography.Pkcs", "10.0.13", "EPPlus");

        string? verdict = PinnedPackagesTests.Verdict(pin, "8.7.1", requires);

        if (problem is null)
        {
            Assert.Null(verdict);
        }
        else
        {
            Assert.NotNull(verdict);
            Assert.Contains(problem, verdict, StringComparison.Ordinal);
            Assert.Contains("remove", verdict, StringComparison.Ordinal);
        }
    }

    [Fact]
    public void A_pin_whose_package_no_longer_brings_it_in_or_is_gone_is_to_be_removed()
    {
        var pin = new Pin("CRT.Data", "System.Security.Cryptography.Pkcs", "10.0.13", "EPPlus");

        Assert.Contains("no longer brings in", PinnedPackagesTests.Verdict(pin, "9.0.0", overRequires: null), StringComparison.Ordinal);
        Assert.Contains("no longer referenced", PinnedPackagesTests.Verdict(pin, overVersion: null, overRequires: null), StringComparison.Ordinal);
    }

    // ###########################################################################################
    // What is wrong with a pin - null while it is still needed. `overVersion` is the version of the
    // package it names as restored (null: not in the graph at all); `overRequires` that package's
    // own requirement of the pinned one, as NuGet writes it ("10.0.10", "[10.0.10, 11.0.0)"), or
    // null when it no longer depends on it.
    // ###########################################################################################
    internal static string? Verdict(Pin pin, string? overVersion, string? overRequires)
    {
        string line = $"{pin.Project}.csproj: {pin.Package} {pin.Version}";

        if (overVersion is null)
            return $"{line} is pinned over {pin.Over}, which is no longer referenced - remove the pin.";

        if (overRequires is null)
            return $"{line}: {pin.Over} {overVersion} no longer brings in {pin.Package} - remove the pin.";

        string minimum = PinnedPackagesTests.MinimumOf(overRequires);

        return PinnedPackagesTests.Compare(minimum, pin.Version) >= 0
            ? $"{line}: {pin.Over} {overVersion} itself requires {minimum} or later, so the pin is no longer needed - remove it."
            : null;
    }

    // "[10.0.10, 11.0.0)" or "10.0.10" -> "10.0.10".
    private static string MinimumOf(string range) =>
        range.Trim().TrimStart('[', '(').Split(',')[0].Trim().TrimEnd(']', ')');

    // Numeric parts compared with missing ones as 0; at equal numbers a pre-release is the lower.
    private static int Compare(string left, string right)
    {
        static (int[] Numbers, bool PreRelease) Parse(string version)
        {
            string core = version.Split('+')[0];
            bool preRelease = core.Contains('-', StringComparison.Ordinal);

            return (core.Split('-')[0].Split('.').Select(int.Parse).ToArray(), preRelease);
        }

        (int[] a, bool aPre) = Parse(left);
        (int[] b, bool bPre) = Parse(right);

        for (int index = 0; index < Math.Max(a.Length, b.Length); index++)
        {
            int difference = (index < a.Length ? a[index] : 0).CompareTo(index < b.Length ? b[index] : 0);

            if (difference != 0)
                return difference;
        }

        return aPre == bPre ? 0 : aPre ? -1 : 1;
    }

    // Every PackageReference with a PinnedOver attribute, in every project file under src.
    private static List<Pin> Pins()
    {
        string src = PinnedPackagesTests.RepositoryPath("src");

        return Directory.EnumerateFiles(src, "*.csproj", SearchOption.AllDirectories)
            .SelectMany(path => XDocument.Load(path).Descendants()
                .Where(element => element.Name.LocalName == "PackageReference" && element.Attribute("PinnedOver") is not null)
                .Select(element => new Pin(
                    Path.GetFileNameWithoutExtension(path),
                    (string?)element.Attribute("Include") ?? string.Empty,
                    (string?)element.Attribute("Version") ?? string.Empty,
                    (string)element.Attribute("PinnedOver")!)))
            .ToList();
    }

    // For each framework the project restored: the version of the package the pin names, and that
    // package's own requirement of the pinned one.
    private static IEnumerable<(string? OverVersion, string? OverRequires)> WhatItRequires(string assetsPath, Pin pin)
    {
        using JsonDocument assets = JsonDocument.Parse(File.ReadAllText(assetsPath));

        var answers = new List<(string?, string?)>();

        // "net10.0", not "net10.0/linux-x64" - a runtime's graph has the same dependencies.
        foreach (JsonProperty target in assets.RootElement.GetProperty("targets").EnumerateObject()
                     .Where(target => !target.Name.Contains('/', StringComparison.Ordinal)))
        {
            JsonProperty? over = target.Value.EnumerateObject()
                .Where(library => library.Name.StartsWith(pin.Over + "/", StringComparison.OrdinalIgnoreCase))
                .Select(library => (JsonProperty?)library)
                .FirstOrDefault();

            if (over is not JsonProperty found)
            {
                answers.Add((null, null));
                continue;
            }

            string? requires = found.Value.TryGetProperty("dependencies", out JsonElement dependencies)
                ? dependencies.EnumerateObject()
                    .Where(dependency => string.Equals(dependency.Name, pin.Package, StringComparison.OrdinalIgnoreCase))
                    .Select(dependency => dependency.Value.GetString())
                    .FirstOrDefault()
                : null;

            answers.Add((found.Name[(pin.Over.Length + 1)..], requires));
        }

        return answers;
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
