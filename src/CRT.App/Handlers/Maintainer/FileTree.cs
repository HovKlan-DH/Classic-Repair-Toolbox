using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using Handlers.DataHandling;

namespace Handlers.MaintainerHandling
{
    // ###########################################################################################
    // A SYSTEM'S FILES AS A FOLDER TREE (owner request, 2026-09-28: "a file-structure for all
    // existing files in BETA or PROD, and then a highlighting of files changed (added, removed,
    // changed) ... a possibility to 'show only changed files' ... including their parent folder").
    //
    // The entries are CRT.Data's SystemFileEntry - one per file, saying what the step does to it -
    // and this turns them into folders, counts the changes under each, and works out which ROWS
    // are on screen. The control (FileTreeView) only draws the rows, as a flat virtualised list:
    // a board is ~2,000 files, and a list that builds only what is visible stays quick where a
    // tree of nested controls would not.
    //
    // Pure, so what the maintainer sees is unit tested.
    // ###########################################################################################
    public sealed class FileTreeNode
    {
        internal FileTreeNode(string name, string path, SystemFileEntry? file)
        {
            this.Name = name;
            this.Path = path;
            this.File = file;
        }

        // The last segment, and the data-root-relative path ("Commodore/C128", or the file's own).
        public string Name { get; }

        public string Path { get; }

        // The file, or null for a folder.
        public SystemFileEntry? File { get; }

        public bool IsFolder => this.File is null;

        public List<FileTreeNode> Children { get; } = [];

        // Files at or under this node, by what the step does to them.
        public int Added { get; internal set; }

        public int Changed { get; internal set; }

        public int Removed { get; internal set; }

        public int Unchanged { get; internal set; }

        public int ChangeCount => this.Added + this.Changed + this.Removed;

        public bool HasChanges => this.ChangeCount > 0;
    }

    // One line on screen: the node, how deep it sits, and - for a folder - whether it is open.
    public sealed record FileTreeRow(FileTreeNode Node, int Depth, bool IsExpanded);

    public static class FileTree
    {
        // ###########################################################################################
        // The tree under an unnamed root, folders before files at every level and each in natural
        // order - "C2" before "C10", as Explorer sorts them - so a folder reads the way it does on
        // disk.
        // ###########################################################################################
        public static FileTreeNode Build(IEnumerable<SystemFileEntry>? entries)
        {
            var root = new FileTreeNode(string.Empty, string.Empty, null);
            var folders = new Dictionary<string, FileTreeNode>(StringComparer.Ordinal) { [string.Empty] = root };

            foreach (SystemFileEntry entry in entries ?? [])
            {
                string[] segments = (entry.Path ?? string.Empty).Replace('\\', '/')
                    .Split('/', StringSplitOptions.RemoveEmptyEntries);

                if (segments.Length == 0)
                    continue;

                FileTreeNode parent = root;
                string path = string.Empty;

                for (int index = 0; index < segments.Length - 1; index++)
                {
                    path = path.Length == 0 ? segments[index] : $"{path}/{segments[index]}";

                    if (!folders.TryGetValue(path, out FileTreeNode? folder))
                    {
                        folder = new FileTreeNode(segments[index], path, null);
                        folders[path] = folder;
                        parent.Children.Add(folder);
                    }

                    parent = folder;
                }

                string filePath = string.Join('/', segments);
                parent.Children.Add(new FileTreeNode(segments[^1], filePath, entry with { Path = filePath }));
            }

            FileTree.SortAndCount(root);
            return root;
        }

        private static void SortAndCount(FileTreeNode node)
        {
            if (!node.IsFolder)
            {
                switch (node.File!.Change)
                {
                    case SystemFileChange.Added: node.Added = 1; break;
                    case SystemFileChange.Changed: node.Changed = 1; break;
                    case SystemFileChange.Removed: node.Removed = 1; break;
                    default: node.Unchanged = 1; break;
                }

                return;
            }

            node.Children.Sort((left, right) =>
                left.IsFolder != right.IsFolder
                    ? (left.IsFolder ? -1 : 1)
                    : NaturalLabelComparer.Instance.Compare(left.Name, right.Name));

            foreach (FileTreeNode child in node.Children)
            {
                FileTree.SortAndCount(child);

                node.Added += child.Added;
                node.Changed += child.Changed;
                node.Removed += child.Removed;
                node.Unchanged += child.Unchanged;
            }
        }

        // ###########################################################################################
        // The folders open when the tree is first shown: every one on the way to a change, so the
        // changes are in view and a folder of 1,500 untouched images is not. For a new system that
        // is everything - which is right: the whole of it is new.
        // ###########################################################################################
        public static IReadOnlySet<string> DefaultExpanded(FileTreeNode root)
        {
            var expanded = new HashSet<string>(StringComparer.Ordinal);

            void Walk(FileTreeNode node)
            {
                foreach (FileTreeNode child in node.Children.Where(child => child.IsFolder && child.HasChanges))
                {
                    expanded.Add(child.Path);
                    Walk(child);
                }
            }

            Walk(root);
            return expanded;
        }

        // Every folder of the tree - "Expand all" (owner request, 2026-09-28).
        public static IReadOnlySet<string> AllFolders(FileTreeNode root)
        {
            ArgumentNullException.ThrowIfNull(root);

            var folders = new HashSet<string>(StringComparer.Ordinal);

            void Walk(FileTreeNode node)
            {
                foreach (FileTreeNode child in node.Children.Where(child => child.IsFolder))
                {
                    folders.Add(child.Path);
                    Walk(child);
                }
            }

            Walk(root);
            return folders;
        }

        // ###########################################################################################
        // The rows on screen, top to bottom; a folder's contents are shown when it is in `expanded`.
        //
        //   onlyChanged - just the files the step adds, changes or removes, each under ALL its
        //                 folders: "show only the changed files, including their parent folder".
        //                 The view opens with every one of those folders open (DefaultExpanded),
        //                 and they close like any other - "Collapse all" works here too, and a
        //                 closed folder still says how many changes are in it.
        //   otherwise   - everything.
        // ###########################################################################################
        public static IReadOnlyList<FileTreeRow> Rows(FileTreeNode root, IReadOnlySet<string> expanded, bool onlyChanged)
        {
            ArgumentNullException.ThrowIfNull(root);
            ArgumentNullException.ThrowIfNull(expanded);

            var rows = new List<FileTreeRow>();

            void Walk(FileTreeNode node, int depth)
            {
                foreach (FileTreeNode child in node.Children)
                {
                    if (onlyChanged && !child.HasChanges)
                        continue;

                    bool open = child.IsFolder && expanded.Contains(child.Path);
                    rows.Add(new FileTreeRow(child, depth, open));

                    if (open)
                        Walk(child, depth + 1);
                }
            }

            Walk(root, 0);
            return rows;
        }
    }

    // ###########################################################################################
    // What the tree says, in words - kept apart so a sentence is changed in one place and tested.
    // ###########################################################################################
    public static class FileTreeWording
    {
        // "3 files change: 1 new, 2 changed, 0 removed. 2,224 already the same." - or, with
        // nothing to change, that it is all the same.
        public static string Summary(FileTreeNode root)
        {
            ArgumentNullException.ThrowIfNull(root);

            string same = $"{FileTreeWording.Count(root.Unchanged, "file")} already the same";

            if (!root.HasChanges)
                return root.Unchanged == 0 ? "No files." : $"Nothing changes - {same}.";

            return $"{FileTreeWording.Count(root.ChangeCount, "file")} " + (root.ChangeCount == 1 ? "changes" : "change") +
                $": {FileTreeWording.Number(root.Added)} new, {FileTreeWording.Number(root.Changed)} changed, " +
                $"{FileTreeWording.Number(root.Removed)} removed. " + char.ToUpperInvariant(same[0]) + same[1..] + ".";
        }

        // The word on a file's row - none for a file that stays as it is.
        public static string ChangeWord(SystemFileChange change) => change switch
        {
            SystemFileChange.Added => "new",
            SystemFileChange.Changed => "changed",
            SystemFileChange.Removed => "removed",
            _ => string.Empty
        };

        // Beside a folder: how many changes are under it, so a closed folder still says so.
        public static string FolderNote(FileTreeNode folder) =>
            folder.HasChanges ? $"{FileTreeWording.Number(folder.ChangeCount)} changing" : string.Empty;

        // ###########################################################################################
        // *** THE ONE THING A FILE'S HOVER CARD SAYS BESIDE ITS PATH - and only when what opens is
        // not the file as it will be. *** The card says nothing about the change (owner request,
        // 2026-09-28: "It should not show any text, if it has changed or whatever"); the row does.
        // But the workbook and the highlight file are written FROM THE TABLE when a submission is
        // approved, so before then its tree can only open BETA's current copy - or nothing, for a
        // new system - and checking that copy for the new one would be checking the wrong file.
        // Null for every other file.
        // ###########################################################################################
        public static string? Note(SystemFileEntry file)
        {
            ArgumentNullException.ThrowIfNull(file);

            if (file.OpenFrom == SystemFileSource.NotWrittenYet)
                return "Written from the table when you approve - there is nothing to open yet.";

            if (file.WrittenOnApproval && file.Change != SystemFileChange.Unchanged && file.OpenFrom == SystemFileSource.Beta)
                return "This opens BETA's copy as it is now - approving writes a new one from the table.";

            return null;
        }

        private static string Number(int count) => count.ToString("N0", CultureInfo.InvariantCulture);

        private static string Count(int count, string noun) =>
            $"{FileTreeWording.Number(count)} {noun}{(count == 1 ? string.Empty : "s")}";
    }
}
