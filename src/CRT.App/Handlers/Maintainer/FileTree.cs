using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using Handlers.DataHandling;

namespace Handlers.MaintainerHandling
{
    // ###########################################################################################
    // A BOARD'S FILES AS A FOLDER TREE (owner request, 2026-09-28: "a file-structure for all
    // existing files in BETA or PROD, and then a highlighting of files changed (added, removed,
    // changed) ... a possibility to 'show only changed files' ... including their parent folder").
    //
    // The entries are CRT.Data's BoardFileEntry - one per file, saying what the step does to it -
    // and this turns them into folders, counts the changes under each, and works out which ROWS
    // are on screen. The control (FileTreeView) only draws the rows, as a flat virtualised list:
    // a board is ~2,000 files, and a list that builds only what is visible stays quick where a
    // tree of nested controls would not.
    //
    // Pure, so what the maintainer sees is unit tested.
    // ###########################################################################################
    public sealed class FileTreeNode
    {
        internal FileTreeNode(string name, string path, BoardFileEntry? file)
        {
            this.Name = name;
            this.Path = path;
            this.File = file;
        }

        // The last segment, and the data-root-relative path ("Commodore/C128", or the file's own).
        public string Name { get; }

        public string Path { get; }

        // The file, or null for a folder.
        public BoardFileEntry? File { get; }

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
        public static FileTreeNode Build(IEnumerable<BoardFileEntry>? entries)
        {
            var root = new FileTreeNode(string.Empty, string.Empty, null);
            var folders = new Dictionary<string, FileTreeNode>(StringComparer.Ordinal) { [string.Empty] = root };

            foreach (BoardFileEntry entry in entries ?? [])
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
                    case BoardFileChange.Added: node.Added = 1; break;
                    case BoardFileChange.Changed: node.Changed = 1; break;
                    case BoardFileChange.Removed: node.Removed = 1; break;
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
        // changes are in view and a folder of 1,500 untouched images is not. For a new board that
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

        // ###########################################################################################
        // The folders open when a LISTING is first shown (the Boards screen's Files view, 2026-10-03):
        // every folder on the way to `folder`, and `folder` itself - the board's own, where nearly
        // all of its files are - and nothing else, so the shared folders around it start closed.
        // Only folders the tree holds; none for a blank or unknown folder.
        // ###########################################################################################
        public static IReadOnlySet<string> FoldersOnTheWayTo(FileTreeNode root, string? folder)
        {
            ArgumentNullException.ThrowIfNull(root);

            IReadOnlySet<string> folders = FileTree.AllFolders(root);
            var open = new HashSet<string>(StringComparer.Ordinal);

            string[] segments = (folder ?? string.Empty).Replace('\\', '/').Split('/', StringSplitOptions.RemoveEmptyEntries);

            for (int i = 1; i <= segments.Length; i++)
            {
                string path = string.Join('/', segments.Take(i));

                if (!folders.Contains(path))
                    break;

                open.Add(path);
            }

            return open;
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
        // ###########################################################################################
        // "3 files change: 1 new, 2 changed, 0 removed. 2,224 already the same." - or, with
        // nothing to change, that it is all the same.
        //
        // *** THE WORKBOOK AND THE HIGHLIGHT FILE ARE SAID APART (2026-09-30). *** An approval writes
        // both FROM THE TABLE, and the workbook changes on every one - it carries the publish date.
        // Counted with the rest, every submission's tree said "1 file changes" at the least, and the
        // headline could not agree with the count on the Files button, which is what the
        // SUBMISSION'S OWN files do (SubmissionViews.ChangingFiles). So they get a sentence of their
        // own, and the number is the button's. Their rows still say "changed".
        // ###########################################################################################
        public static string Summary(FileTreeNode root)
        {
            ArgumentNullException.ThrowIfNull(root);

            var written = new List<BoardFileEntry>();
            int added = 0, changed = 0, removed = 0;

            void Walk(FileTreeNode node)
            {
                foreach (FileTreeNode child in node.Children)
                {
                    if (child.IsFolder)
                    {
                        Walk(child);
                        continue;
                    }

                    BoardFileEntry file = child.File!;

                    if (file.Change == BoardFileChange.Unchanged)
                        continue;

                    // *** ONLY A FILE THE APPROVAL WRITES - added or changed - IS "WRITTEN" (code review,
                    // 2026-10-01). *** One that is REMOVED is a removal whatever generated it: filed as
                    // "written" it was reported as "Approving writes the workbook" for a tree that
                    // deletes it. ForApproval never builds that today; this keeps the sentence honest
                    // if a source ever does.
                    if (file.WrittenOnApproval && file.Change != BoardFileChange.Removed)
                        written.Add(file);
                    else if (file.Change == BoardFileChange.Added)
                        added++;
                    else if (file.Change == BoardFileChange.Changed)
                        changed++;
                    else
                        removed++;
                }
            }

            Walk(root);

            int counted = added + changed + removed;
            string same = $"{FileTreeWording.Count(root.Unchanged, "file")} already the same";
            string? writes = FileTreeWording.WrittenFiles(written);

            if (counted == 0)
            {
                if (writes is not null)
                    return $"Approving writes {writes} from the table; nothing else changes. " + char.ToUpperInvariant(same[0]) + same[1..] + ".";

                return root.Unchanged == 0 ? "No files." : $"Nothing changes - {same}.";
            }

            string also = writes is null ? string.Empty : $" Approving also writes {writes} from the table.";

            return $"{FileTreeWording.Count(counted, "file")} " + (counted == 1 ? "changes" : "change") +
                $": {FileTreeWording.Number(added)} new, {FileTreeWording.Number(changed)} changed, " +
                $"{FileTreeWording.Number(removed)} removed.{also} " + char.ToUpperInvariant(same[0]) + same[1..] + ".";
        }

        // A listing's line (2026-10-03): how many files, nothing about changes - "2,224 files", "1 file",
        // "No files."
        public static string Listing(FileTreeNode root)
        {
            ArgumentNullException.ThrowIfNull(root);

            int files = root.Unchanged + root.ChangeCount;

            return files == 0 ? "No files." : $"{FileTreeWording.Count(files, "file")}.";
        }

        // "the workbook", "the highlight file" or both - named by what they are, not by path. Null
        // for neither.
        private static string? WrittenFiles(IReadOnlyList<BoardFileEntry> written)
        {
            bool workbook = written.Any(file => !BoardFileEntries.IsHighlightFile(file.Path));
            bool sidecar = written.Any(file => BoardFileEntries.IsHighlightFile(file.Path));

            return (workbook, sidecar) switch
            {
                (true, true) => "the workbook and the highlight file",
                (true, false) => "the workbook",
                (false, true) => "the highlight file",
                _ => null
            };
        }

        // A file's size on its row (owner request, 2026-10-04) - none for a folder, or for a file
        // whose size the server did not say (not written yet, or an older server).
        public static string Size(FileTreeNode node)
        {
            ArgumentNullException.ThrowIfNull(node);

            return node.File?.SizeBytes is long bytes ? FileSizeWording.Format(bytes) : string.Empty;
        }

        // The word on a file's row - none for a file that stays as it is.
        public static string ChangeWord(BoardFileChange change) => change switch
        {
            BoardFileChange.Added => "new",
            BoardFileChange.Changed => "changed",
            BoardFileChange.Removed => "removed",
            _ => string.Empty
        };

        // Beside a folder: how many files change under it, so a closed folder still says so -
        // "1 file changed", "3 files changed" (owner request, 2026-10-01; it read "3 changing").
        // "Changed" covers a new and a removed file too; each file's own row says which.
        public static string FolderNote(FileTreeNode folder) =>
            folder.HasChanges
                ? $"{FileTreeWording.Number(folder.ChangeCount)} {(folder.ChangeCount == 1 ? "file" : "files")} changed"
                : string.Empty;

        // ###########################################################################################
        // *** THE ONE THING A FILE'S HOVER CARD SAYS BESIDE ITS PATH - and only when what opens is
        // not the file as it will be. *** The card says nothing about the change (owner request,
        // 2026-09-28: "It should not show any text, if it has changed or whatever"); the row does.
        // But the workbook and the highlight file are written FROM THE TABLE when a submission is
        // approved, so before then its tree can only open BETA's current copy - or nothing, for a
        // new board - and checking that copy for the new one would be checking the wrong file.
        // Null for every other file.
        // ###########################################################################################
        public static string? Note(BoardFileEntry file)
        {
            ArgumentNullException.ThrowIfNull(file);

            if (file.OpenFrom == BoardFileSource.NotWrittenYet)
                return "Written from the table when you approve - there is nothing to open yet.";

            if (file.WrittenOnApproval && file.Change != BoardFileChange.Unchanged && file.OpenFrom == BoardFileSource.Beta)
                return "This opens BETA's copy as it is now - approving writes a new one from the table.";

            return null;
        }

        private static string Number(int count) => count.ToString("N0", CultureInfo.InvariantCulture);

        private static string Count(int count, string noun) =>
            $"{FileTreeWording.Number(count)} {noun}{(count == 1 ? string.Empty : "s")}";
    }
}
