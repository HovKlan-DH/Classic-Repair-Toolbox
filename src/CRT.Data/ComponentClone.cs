using System;
using System.Collections.Generic;
using System.Linq;

namespace Handlers.DataHandling
{
    // What "Clone component from another board" says of the component picked (ComponentClone.FactsOf):
    // its row and what it holds, read in `Region` - and whether the other regions' come along, or are left
    // out because this component has no place for them.
    public sealed record ComponentCloneFacts(
        ComponentEntry Row,
        string Region,
        int Pins,
        int Replacements,
        int Pictures,
        int Baselines,
        int Files,
        int Links,
        bool EveryRegionComes,
        bool OtherRegionsLeftOut);

    // One file the rows taken cite: its name there, its name here, and whether its bytes have to be
    // copied in - not when this draft lists the very same file under the very same name.
    public sealed record ComponentCloneFile(string NameThere, string NameHere, bool CopyIn);

    // ###########################################################################################
    // CLONE A COMPONENT FROM ANOTHER BOARD (owner request, 2026-11-04: "I would like to get a selector, so I
    // can select a hardware, a board and then a component ... clone all its data to this component. It will
    // not change the component name, but all other data should be identical to the one it copies from").
    //
    // "Copy from another board" until then, which offered only the same part number and let parts be ticked.
    // ComponentCopy stays as what "Also on other boards" puts on other boards.
    //
    //   - EVERYTHING BUT THE LABEL COMES AS ONE STEP: the fields, category and package included, the pins,
    //     replacement parts, pictures, oscilloscope baselines, files and links. What the component held of
    //     those goes; a file of the other board comes with its bytes (the host copies them in, FilesCited
    //     and NamesHere saying which).
    //   - WHERE IT IS MARKED STAYS. A rectangle is a place on THIS board's pictures - another board's would
    //     mark nothing, or the wrong spot - so the highlights are left alone.
    //   - REGIONS: AS IT STANDS THERE, AS FAR AS THIS COMPONENT HAS ROOM (ReadingOf):
    //       * in every region here, and there too: its rows as they are, each saying what it says - while the
    //         two boards come in the same regions; one in regions this board has not (PAL and NTSC onto EU and
    //         US) comes as one region there shows, for every region here;
    //       * in every region here, in one region there: that region's, for every region - this component
    //         must not vanish from the other region for being cloned;
    //       * in one region here: what that region shows there, in this one - the component staying in it
    //         whatever its row says there, and the rest of what is for every region staying for every region;
    //       * on a board with no regions: what the region shown shows there, nothing saying a region -
    //         a row that said one would switch this board's PAL/NTSC buttons on.
    //   - A ROW CITING A FILE THAT DID NOT COME IS LEFT OUT (Apply's `fileNames`), and the pane says so.
    //
    // PURE: boards in, a board out. The same board when nothing would change, so no step is made.
    // ###########################################################################################
    public static class ComponentClone
    {
        // Which rows of the other board are taken, and in which region each is filed here. OnlyRegion: the
        // one region this component is in, "" on a board with no regions - the rows of Components say WHERE it
        // is, so it gets ONE, the one that region shows. One for every region there must not make it one for
        // every region here (code review, 2026-11-04: a chip with no region there took U8 of NTSC into PAL too),
        // nor a row for every region and one of the region two rows the same (code review, 2026-11-04), while
        // its pins and links for every region stay so.
        private sealed record Reading(string? Region, Func<string?, string> Filed, string? OnlyRegion = null)
        {
            public bool Takes(string? rowRegion) => this.Region is null || ComponentPaneEdits.InRegion(rowRegion, this.Region);
        }

        // Every component of a board once, as its rows spell it, in the order they stand.
        public static List<string> Labels(BoardData board)
        {
            ArgumentNullException.ThrowIfNull(board);

            return board.Components
                .Select(row => row.BoardLabel?.Trim() ?? string.Empty)
                .Where(label => label.Length > 0)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        // ###########################################################################################
        // How the component `sourceLabel` of `source` would be read for the component `label` of `target`,
        // `shown` being the region CRT shows. Null when either has no such component.
        // ###########################################################################################
        private static Reading? ReadingOf(BoardData target, string? label, BoardData source, string? sourceLabel, string? shown)
        {
            List<ComponentEntry> here = target.Components.Where(row => ComponentPaneEdits.IsOfComponent(row.BoardLabel, label)).ToList();
            List<ComponentEntry> there = source.Components.Where(row => ComponentPaneEdits.IsOfComponent(row.BoardLabel, sourceLabel)).ToList();

            if (here.Count == 0 || there.Count == 0)
            {
                return null;
            }

            // The region wanted if the component is in it there, else the region it is in.
            string RegionThere(string wanted) =>
                there.Exists(row => ComponentPaneEdits.InRegion(row.Region, wanted)) ? wanted : there[0].Region?.Trim() ?? string.Empty;

            if (!BoardRegions.Has(target))
            {
                return new Reading(RegionThere(string.IsNullOrWhiteSpace(shown) ? BoardRegions.Of(source).FirstOrDefault() ?? BoardRegions.Pal : shown.Trim()), _ => string.Empty, string.Empty);
            }

            string oneHere = here.Count == 1 ? here[0].Region?.Trim() ?? string.Empty : string.Empty;

            if (oneHere.Length > 0)
            {
                return new Reading(RegionThere(oneHere), region => string.IsNullOrWhiteSpace(region) ? string.Empty : oneHere, oneHere);
            }

            string oneThere = there.Count == 1 ? there[0].Region?.Trim() ?? string.Empty : string.Empty;

            if (oneThere.Length > 0)
            {
                return new Reading(oneThere, _ => string.Empty);
            }

            // In every region on both: its rows as they are - only while the two boards come in the same regions, or nothing
            // of it says a region there (code review, 2026-11-09: a PAL/NTSC board's chip cloned onto a board in EU and US came
            // filed as PAL and NTSC, which that board shows in no region). Else what one region shows there - the one shown
            // here when that board has it, else its first - for every region here.
            if (ComponentClone.SameRegions(target, source) || !ComponentClone.RegionsThere(source, sourceLabel).Any())
            {
                return new Reading(null, region => region?.Trim() ?? string.Empty);
            }

            IReadOnlyList<string> regionsThere = BoardRegions.Of(source);
            string read = regionsThere.FirstOrDefault(region => string.Equals(region, shown?.Trim(), StringComparison.OrdinalIgnoreCase))
                          ?? regionsThere.FirstOrDefault()
                          ?? ComponentClone.RegionsThere(source, sourceLabel).First();

            return new Reading(read, _ => string.Empty);
        }

        // Whether two boards come in the same regions, whatever their order and capitals.
        private static bool SameRegions(BoardData target, BoardData source) =>
            new HashSet<string>(BoardRegions.Of(target), StringComparer.OrdinalIgnoreCase).SetEquals(BoardRegions.Of(source));

        // ###########################################################################################
        // What the panel says of the component picked, before anything changes - null without one. What
        // it holds is counted in one region (the one taken, else the one shown): both regions added up would
        // read as twice the pins.
        // ###########################################################################################
        public static ComponentCloneFacts? FactsOf(BoardData target, string? label, BoardData source, string? sourceLabel, string? shown)
        {
            ArgumentNullException.ThrowIfNull(target);
            ArgumentNullException.ThrowIfNull(source);

            if (ComponentClone.ReadingOf(target, label, source, sourceLabel, shown) is not Reading reading)
            {
                return null;
            }

            string region = reading.Region ?? shown?.Trim() ?? string.Empty;

            bool InView(string? rowLabel, string? rowRegion) => ComponentPaneEdits.InScope(rowLabel, rowRegion, sourceLabel, region);

            List<ComponentImageEntry> images = source.ComponentImages.Where(row => InView(row.BoardLabel, row.Region)).ToList();

            bool saysARegion = ComponentClone.RegionsThere(source, sourceLabel).Any();

            return new ComponentCloneFacts(
                ComponentPaneEdits.RowShown(source, sourceLabel, region)!,
                region,
                source.ComponentPins.Count(row => InView(row.BoardLabel, row.Region)),
                source.ComponentReplacements.Count(row => InView(row.BoardLabel, row.Region)),
                images.Count(row => !ComponentImageKinds.IsScopeBaseline(row)),
                images.Count(ComponentImageKinds.IsScopeBaseline),
                source.ComponentLocalFiles.Count(row => InView(row.BoardLabel, row.Region)),
                source.ComponentLinks.Count(row => InView(row.BoardLabel, row.Region)),
                reading.Region is null && saysARegion,
                reading.Region is not null && ComponentClone.RegionsThere(source, sourceLabel).Any(other => !string.Equals(other, reading.Region, StringComparison.OrdinalIgnoreCase)));
        }

        // Every region a row of the component says there.
        private static IEnumerable<string> RegionsThere(BoardData source, string? sourceLabel)
        {
            bool Of(string? rowLabel) => ComponentPaneEdits.IsOfComponent(rowLabel, sourceLabel);

            return source.Components.Where(row => Of(row.BoardLabel)).Select(row => row.Region)
                .Concat(source.ComponentImages.Where(row => Of(row.BoardLabel)).Select(row => row.Region))
                .Concat(source.ComponentPins.Where(row => Of(row.BoardLabel)).Select(row => row.Region))
                .Concat(source.ComponentReplacements.Where(row => Of(row.BoardLabel)).Select(row => row.Region))
                .Concat(source.ComponentLocalFiles.Where(row => Of(row.BoardLabel)).Select(row => row.Region))
                .Concat(source.ComponentLinks.Where(row => Of(row.BoardLabel)).Select(row => row.Region))
                .Where(region => !string.IsNullOrWhiteSpace(region))
                .Select(region => region.Trim())
                .Distinct(StringComparer.OrdinalIgnoreCase);
        }

        // ###########################################################################################
        // The files the rows taken cite, each once, as BoardFileNames spells it - what has to come along.
        // ###########################################################################################
        public static List<string> FilesCited(BoardData target, string? label, BoardData source, string? sourceLabel, string? shown)
        {
            ArgumentNullException.ThrowIfNull(target);
            ArgumentNullException.ThrowIfNull(source);

            if (ComponentClone.ReadingOf(target, label, source, sourceLabel, shown) is not Reading reading)
            {
                return [];
            }

            bool Taken(string? rowLabel, string? rowRegion) => ComponentPaneEdits.IsOfComponent(rowLabel, sourceLabel) && reading.Takes(rowRegion);

            return source.ComponentImages.Where(row => Taken(row.BoardLabel, row.Region)).Select(row => row.File)
                .Concat(source.ComponentLocalFiles.Where(row => Taken(row.BoardLabel, row.Region)).Select(row => row.File))
                .Select(BoardFileNames.Normalise)
                .Where(name => name.Length > 0)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        // ###########################################################################################
        // Where each of `cited` goes in this draft. The very same file is already here when this draft lists
        // its bytes (NameOfSameFileHere); anything else comes under a name no file of the board has or an undo
        // can bring back (`taken`), so it replaces nothing another row of this board cites.
        // ###########################################################################################
        public static List<ComponentCloneFile> NamesHere(
            IEnumerable<string> cited,
            IReadOnlyDictionary<string, BoardDocument.StoredFileRow>? filesThere,
            IReadOnlyDictionary<string, BoardDocument.StoredFileRow>? filesHere,
            IEnumerable<string> taken)
        {
            ArgumentNullException.ThrowIfNull(cited);

            var used = new List<string>(taken ?? []);
            var files = new List<ComponentCloneFile>();

            foreach (string name in cited)
            {
                if (ComponentClone.NameOfSameFileHere(name, filesThere, filesHere) is { } here)
                {
                    files.Add(new ComponentCloneFile(name, here, CopyIn: false));
                    continue;
                }

                string free = BoardFileNames.FreeName(name, used);

                used.Add(free);
                files.Add(new ComponentCloneFile(name, free, CopyIn: true));
            }

            return files;
        }

        // ###########################################################################################
        // The name this draft lists the same bytes under: its own name first, else another in the same
        // folder. A file that needed a free name the first time is that same file the second time (code
        // review, 2026-11-04: cloning again copied it under a third name, one more step and "Cloned" for no
        // change). Only the same folder, which says what kind of file it is: a picture is no datasheet because
        // the bytes happen to be the same. Null when it has to be copied in.
        // ###########################################################################################
        private static string? NameOfSameFileHere(
            string name,
            IReadOnlyDictionary<string, BoardDocument.StoredFileRow>? filesThere,
            IReadOnlyDictionary<string, BoardDocument.StoredFileRow>? filesHere)
        {
            if (filesHere is null ||
                filesThere is null ||
                !filesThere.TryGetValue(name, out BoardDocument.StoredFileRow? there) ||
                string.IsNullOrWhiteSpace(there.Sha256))
            {
                return null;
            }

            bool SameBytes(BoardDocument.StoredFileRow file) => string.Equals(file.Sha256, there.Sha256, StringComparison.OrdinalIgnoreCase);

            if (filesHere.TryGetValue(name, out BoardDocument.StoredFileRow? own) && SameBytes(own))
            {
                return name;
            }

            string folder = ComponentClone.FolderOf(name);

            return filesHere
                .Where(file => SameBytes(file.Value) && string.Equals(ComponentClone.FolderOf(file.Key), folder, StringComparison.OrdinalIgnoreCase))
                .Select(file => file.Key)
                .OrderBy(key => key, StringComparer.OrdinalIgnoreCase)
                .FirstOrDefault();
        }

        private static string FolderOf(string name)
        {
            int slash = name.LastIndexOf('/');

            return slash < 0 ? string.Empty : name[..slash];
        }

        // ###########################################################################################
        // `target` with the component `label` holding what `sourceLabel` holds on `source`: every row of it,
        // in every section, replaced - the label and the highlights staying. `fileNames` is each file's name
        // there to its name here; a row citing one it does not list is left out (the file did not come).
        // The same board when that changes nothing.
        // ###########################################################################################
        public static BoardData Apply(
            BoardData target,
            string? label,
            BoardData source,
            string? sourceLabel,
            string? shown,
            IReadOnlyDictionary<string, string> fileNames)
        {
            ArgumentNullException.ThrowIfNull(target);
            ArgumentNullException.ThrowIfNull(source);
            ArgumentNullException.ThrowIfNull(fileNames);

            if (ComponentClone.ReadingOf(target, label, source, sourceLabel, shown) is not Reading reading)
            {
                return target;
            }

            string own = target.Components.First(row => ComponentPaneEdits.IsOfComponent(row.BoardLabel, label)).BoardLabel.Trim();

            bool Taken(string? rowLabel, string? rowRegion) => ComponentPaneEdits.IsOfComponent(rowLabel, sourceLabel) && reading.Takes(rowRegion);
            bool Here(string? rowLabel) => ComponentPaneEdits.IsOfComponent(rowLabel, label);

            // A row with no file keeps none - a note on its own is the row's. One citing a file that did not
            // come is left out.
            bool FileCame(string? file) => string.IsNullOrWhiteSpace(file) || fileNames.ContainsKey(BoardFileNames.Normalise(file));

            string FileHere(string? file) =>
                string.IsNullOrWhiteSpace(file) ? string.Empty : fileNames[BoardFileNames.Normalise(file)];

            BoardData cloned = target.WithComponentSections(
                components: ComponentPaneEdits.ReplaceInScope(
                    target.Components,
                    row => Here(row.BoardLabel),
                    reading.OnlyRegion is { } onlyRegion
                        ? [ComponentPaneEdits.RowShown(source, sourceLabel, reading.Region)!.WithBoardLabel(own).WithRegion(onlyRegion)]
                        : source.Components.Where(row => Taken(row.BoardLabel, row.Region))
                            .Select(row => row.WithBoardLabel(own).WithRegion(reading.Filed(row.Region)))
                            .ToList()),
                componentImages: ComponentPaneEdits.ReplaceInScope(
                    target.ComponentImages,
                    row => Here(row.BoardLabel),
                    source.ComponentImages.Where(row => Taken(row.BoardLabel, row.Region) && FileCame(row.File))
                        .Select(row => row.FiledAs(own, reading.Filed(row.Region), FileHere(row.File)))
                        .ToList()),
                componentLocalFiles: ComponentPaneEdits.ReplaceInScope(
                    target.ComponentLocalFiles,
                    row => Here(row.BoardLabel),
                    source.ComponentLocalFiles.Where(row => Taken(row.BoardLabel, row.Region) && FileCame(row.File))
                        .Select(row => row.FiledAs(own, reading.Filed(row.Region), FileHere(row.File)))
                        .ToList()),
                componentLinks: ComponentPaneEdits.ReplaceInScope(
                    target.ComponentLinks,
                    row => Here(row.BoardLabel),
                    source.ComponentLinks.Where(row => Taken(row.BoardLabel, row.Region))
                        .Select(row => row.FiledAs(own, reading.Filed(row.Region)))
                        .ToList()),
                componentPins: ComponentPaneEdits.ReplaceInScope(
                    target.ComponentPins,
                    row => Here(row.BoardLabel),
                    source.ComponentPins.Where(row => Taken(row.BoardLabel, row.Region))
                        .Select(row => row.WithBoardLabel(own).WithRegion(reading.Filed(row.Region)))
                        .ToList()),
                componentReplacements: ComponentPaneEdits.ReplaceInScope(
                    target.ComponentReplacements,
                    row => Here(row.BoardLabel),
                    source.ComponentReplacements.Where(row => Taken(row.BoardLabel, row.Region))
                        .Select(row => row.FiledAs(own, reading.Filed(row.Region)))
                        .ToList()));

            return BoardTableDocument.HasSameRows(target, cloned, BoardWorkbookSchema.AllSheets) ? target : cloned;
        }
    }
}
