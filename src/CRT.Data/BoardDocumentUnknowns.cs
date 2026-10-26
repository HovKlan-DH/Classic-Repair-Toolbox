using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Text.Json.Serialization;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // WHAT A BOARD FILE HOLDS THAT THIS CRT DOES NOT KNOW, CARRIED TO THE FILE WRITTEN OVER IT
    // (code review, 2026-10-25).
    //
    // A board file is rebuilt from BoardData at every write (BoardDocumentMapping.ToDocument), and
    // BoardData's rows have no place for a field this CRT does not know. Only the file's TOP-LEVEL
    // fields were kept, so a field a later CRT adds to a component row was dropped from every row by
    // the first edit an installed 4.0.0 made - and an installed CRT stays installed for good.
    //
    // So every writer that replaces a board file hands this the file as it was and the file it is
    // about to write, and each new row takes the unknown fields of ITS row in the old file:
    //
    //   1. the row with exactly the same values - the row nobody touched (and of two identical rows,
    //      each keeps its own);
    //   2. else the row with the same identity - BoardDraftNaturalKeys, the pairing rule the table
    //      and the differ use: a component whose description was edited is still that component.
    //
    // A row whose identity changed, and a new row, take nothing. A field the new file already has
    // is never overwritten.
    // ###########################################################################################
    public static class BoardDocumentUnknowns
    {
        public static void CarryOver(BoardDocument? previous, BoardDocument next)
        {
            ArgumentNullException.ThrowIfNull(next);

            if (previous is null)
                return;

            next.Unknown ??= previous.Unknown;

            Carry(previous.Schematics, next.Schematics, row => BoardDraftNaturalKeys.ForSchematic(row.Name));
            Carry(previous.Components, next.Components, row => BoardDraftNaturalKeys.ForComponent(row.BoardLabel, row.Region));
            Carry(previous.ComponentImages, next.ComponentImages, row => BoardDraftNaturalKeys.ForComponentImage(row.BoardLabel, row.Region, row.Pin, row.Name));
            Carry(previous.ScopeBaselines, next.ScopeBaselines, row => BoardDraftNaturalKeys.ForComponentImage(row.BoardLabel, row.Region, row.Pin, row.Name));
            Carry(previous.ComponentFiles, next.ComponentFiles, row => BoardDraftNaturalKeys.ForComponentLocalFile(row.BoardLabel, row.Name));
            Carry(previous.ComponentLinks, next.ComponentLinks, row => BoardDraftNaturalKeys.ForComponentLink(row.BoardLabel, row.Name));
            Carry(previous.ComponentPins, next.ComponentPins, row => BoardDraftNaturalKeys.ForComponentPin(row.BoardLabel, row.Region, row.Pin));
            Carry(previous.ComponentReplacements, next.ComponentReplacements, row => BoardDraftNaturalKeys.ForComponentReplacement(row.BoardLabel, row.Region, row.PartNumber));
            Carry(previous.BoardFiles, next.BoardFiles, row => BoardDraftNaturalKeys.ForBoardLocalFile(row.Category, row.Name));
            Carry(previous.BoardLinks, next.BoardLinks, row => BoardDraftNaturalKeys.ForBoardLink(row.Category, row.Name));
            Carry(previous.Credits, next.Credits, row => BoardDraftNaturalKeys.ForCredit(row.Category, row.SubCategory, row.NameOrHandle));
            Carry(previous.ImportantSignals, next.ImportantSignals, row => BoardDraftNaturalKeys.ForKiCadImportantSignal(row.DisplayName, row.KiCadNetName));
            Carry(previous.Highlights, next.Highlights, row => BoardDraftNaturalKeys.ForComponentHighlight(row.Schematic, row.BoardLabel));
            Carry(previous.KiCadCalibrations, next.KiCadCalibrations, row => BoardDraftNaturalKeys.ForSchematic(row.Schematic));
            Carry(previous.SchematicRenames, next.SchematicRenames, naturalKey: null);
            Carry(previous.Files, next.Files, row => row.Name.Trim().ToUpperInvariant());
        }

        private static void Carry<T>(List<T> previous, List<T> next, Func<T, string>? naturalKey)
            where T : class, IKeepsUnknownFields
        {
            List<T> holding = previous.Where(row => row.Unknown is { Count: > 0 }).ToList();

            if (holding.Count == 0)
                return;

            var byValues = holding
                .GroupBy(BoardDocumentUnknowns.ValuesOf, StringComparer.Ordinal)
                .ToDictionary(group => group.Key, group => new Queue<T>(group), StringComparer.Ordinal);

            var taken = new HashSet<T>(ReferenceEqualityComparer.Instance);
            var waiting = new List<T>();

            foreach (T row in next.Where(row => row.Unknown is not { Count: > 0 }))
            {
                if (byValues.TryGetValue(BoardDocumentUnknowns.ValuesOf(row), out Queue<T>? same) && same.TryDequeue(out T? match))
                {
                    row.Unknown = match.Unknown;
                    taken.Add(match);
                }
                else
                {
                    waiting.Add(row);
                }
            }

            if (naturalKey is null || waiting.Count == 0)
                return;

            var byIdentity = holding
                .Where(row => !taken.Contains(row))
                .GroupBy(naturalKey, StringComparer.OrdinalIgnoreCase)
                .ToDictionary(group => group.Key, group => new Queue<T>(group), StringComparer.OrdinalIgnoreCase);

            foreach (T row in waiting)
            {
                if (byIdentity.TryGetValue(naturalKey(row), out Queue<T>? same) && same.TryDequeue(out T? match))
                    row.Unknown = match.Unknown;
            }
        }

        // ---- a row's known values, as one text -------------------------------------------------

        private static readonly ConcurrentDictionary<Type, PropertyInfo[]> KnownProperties = new();

        private static string ValuesOf<T>(T row) where T : class
        {
            PropertyInfo[] properties = BoardDocumentUnknowns.KnownProperties.GetOrAdd(
                typeof(T),
                type => type.GetProperties(BindingFlags.Public | BindingFlags.Instance)
                    .Where(property => property.GetCustomAttribute<JsonPropertyNameAttribute>() is not null)
                    .OrderBy(property => property.Name, StringComparer.Ordinal)
                    .ToArray());

            return string.Join('\u001F', properties.Select(property => property.GetValue(row) switch
            {
                null => "\u0000",
                IFormattable formattable => formattable.ToString(null, CultureInfo.InvariantCulture),
                object value => value.ToString(),
            }));
        }
    }
}
