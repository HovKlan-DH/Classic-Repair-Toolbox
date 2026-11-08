using CRT;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // ONE COMPUTER THE USER OWNS (CRT 4.0 step 4 - the plan: "Your own machines, each with its
    // repairs"; owner decision, 2026-10-24: called a "Device" wherever a person reads it, "My
    // devices" and "Add device", and OwnedDevice in code so it never mixes with minipro's "device",
    // the chip type). Its name, its hardware and board, region, serial number, its description, its
    // condition, comments, links of interest, photos, its status and its place when sorted by hand. Its repairs are workbooks naming it
    // (WorkbookRecord.OwnedDeviceId), never a list here - a workbook deleted leaves nothing behind.
    //
    // A device has no list of changes made to it any more (owner request, 2026-11-06: "remove the
    // 'Changes made to it', as this is exactly what the workbooks will cover"). The "changes" an
    // earlier CRT 4.0 pre-release wrote stay in the file, untouched: Unknown keeps them and writes
    // them back, as any field this CRT does not know - nothing a user typed is thrown away.
    //
    // Local data, like workbooks: never synced, never sent anywhere.
    // ###########################################################################################
    public sealed class OwnedDeviceRecord
    {
        [JsonPropertyName("id")] public int Id { get; set; }
        [JsonPropertyName("name")] public string Name { get; set; } = string.Empty;

        // "{hardware}|{board}", the shape a workbook's BoardKey has (Main.GetCurrentBoardKey).
        [JsonPropertyName("boardKey")] public string BoardKey { get; set; } = string.Empty;
        // PAL or NTSC - offered only for a board that comes in both (BoardRegions), and kept as it is
        // on a board that does not.
        [JsonPropertyName("region")] public string Region { get; set; } = string.Empty;
        [JsonPropertyName("serialNumber")] public string SerialNumber { get; set; } = string.Empty;

        // Its description - several lines, as typed. "Where it came from" until 2026-11-06 (owner request:
        // 'change text "Where did it come from" to "Description"'), so it is STORED as "origin": a name
        // already stored keeps its spelling.
        [JsonPropertyName("origin")] public string Description { get; set; } = string.Empty;

        // ###########################################################################################
        // "notes" - the device page's own box until 2026-11-07 (owner request: 'Remove the "Notes" as we now have the
        // "Description"'). Read only so notes a device had are FOLDED INTO ITS DESCRIPTION as it is read
        // (OwnedDeviceStore.FoldNotes), and never written: the next save writes them as the description.
        // ###########################################################################################
        [JsonPropertyName("notes")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public string? Notes { get; set; }

        // Its place in "My devices" sorted by hand (owner request, 2026-11-07: "Sort manually in list (drag up/down)";
        // OwnedDeviceSorting.Manual) - 1 for the first; 0 for a device never placed, listed after those that were.
        [JsonPropertyName("displayOrder")] public int DisplayOrder { get; set; }

        // Its condition, 0 to 5 stars (owner request, 2026-11-06: 'can we add a "Condition" and then have 0
        // to 5 stars?') - 0, no star, for a device written before it, or one nobody rated.
        [JsonPropertyName("condition")] public int Condition { get; set; }

        // One of OwnedDeviceStatus' words, stored as the word.
        [JsonPropertyName("status")] public string Status { get; set; } = OwnedDeviceStatus.Unknown;
        [JsonPropertyName("createdDate")] public DateTime CreatedDate { get; set; }

        // Its comments, each dated - a worklog's own kind (owner request, 2026-11-06: the device form "more
        // like the 'Worklog' modal", with "Comments" and "Photos").
        [JsonPropertyName("comments")] public List<WorklogCommentRecord> Comments { get; set; } = new();

        // Its links of interest - a worklog's own kind too (owner request, 2026-11-06: 'Just like the worklogs has,
        // then do add the same "Links of interest" for the device').
        [JsonPropertyName("links")] public List<WorklogLinkRecord> Links { get; set; } = new();

        // The photos lie in the device's own folder under their FileName, as a worklog's do.
        [JsonPropertyName("photos")] public List<WorklogAttachmentRecord> Photos { get; set; } = new();

        // Whatever a later CRT writes that this one does not know - kept and written back, so
        // opening a device in an older CRT loses nothing. The "changes" an earlier pre-release wrote
        // are kept here too (see the class header).
        [JsonExtensionData] public Dictionary<string, JsonElement>? Unknown { get; set; }
    }

    // ###########################################################################################
    // What the "Add device" form collects - every field of a device the form edits. Its comments and
    // photos are lists of their own (OwnedDeviceStore.AddComment, AddPhotos).
    // ###########################################################################################
    public sealed record OwnedDeviceDetails(
        string Name,
        string BoardKey,
        string Region,
        string SerialNumber,
        string Description,
        string Status,
        int Condition = 0);

    // ###########################################################################################
    // A photo to copy into a device's folder: where it lies, and the comment it is to carry.
    // ###########################################################################################
    public sealed record OwnedDevicePhotoSource(string Path, string Comment);

    // ###########################################################################################
    // A DEVICE'S CONDITION - 0 to 5 stars (owner request, 2026-11-06). Anything else read is taken to
    // the nearest end, so a hand-edited file never draws six stars or a negative count.
    // ###########################################################################################
    public static class OwnedDeviceCondition
    {
        public const int MaxStars = 5;

        public static int Within(int stars) => Math.Clamp(stars, 0, MaxStars);

        // ###########################################################################################
        // The stars after a click on star `clicked` (1..5): that many - or none, when it is the
        // condition already set, so a rating can be taken away again.
        // ###########################################################################################
        public static int AfterClick(int current, int clicked)
        {
            int star = Within(clicked);
            return star == Within(current) ? 0 : star;
        }
    }

    // ###########################################################################################
    // THE WORDS A DEVICE'S STATUS IS STORED AS, in the order they are offered: "Unknown" first (owner
    // request, 2026-11-06: 'remove the "Not mine any more" in "Status" as this is different data as I
    // see it. Instead have a status of "Unknown"', shown first). A status this CRT does not know - a
    // later one's, or "Not mine any more" an earlier pre-release stored - is kept as it is, never forced
    // into one of these: the form offers it beside them.
    // ###########################################################################################
    public static class OwnedDeviceStatus
    {
        public const string Unknown = "Unknown";
        public const string Working = "Working";
        public const string Faulty = "Faulty";
        public const string InRepair = "In repair";
        public const string ForParts = "For parts";

        public static readonly IReadOnlyList<string> All = [Unknown, Working, Faulty, InRepair, ForParts];

        // The known word in its own spelling, a blank one Unknown, an unknown one as it is.
        public static string Normalise(string? status)
        {
            string trimmed = status?.Trim() ?? string.Empty;

            if (trimmed.Length == 0)
            {
                return Unknown;
            }

            return All.FirstOrDefault(known => string.Equals(known, trimmed, StringComparison.OrdinalIgnoreCase)) ?? trimmed;
        }

        // ###########################################################################################
        // The theme colour a status is drawn in: working green, faulty red, in repair amber - the
        // colours CRT's own pass, fail and warning lines use - and the rest muted.
        // ###########################################################################################
        public static string ColourKeyOf(string? status) => Normalise(status) switch
        {
            Working => "Text_Success_Fg",
            Faulty => "Text_Fail_Fg",
            InRepair => "Text_Warning_Fg",
            _ => "Workbooks_Muted_Fg",
        };
    }

    // ###########################################################################################
    // READS AND WRITES THE USER'S DEVICES - one "device_{id}" folder each in the Workbooks folder
    // (WorklogManager.WorkbookRoot), holding its index.json and its photos, beside the
    // "workbook_{id}" folders; the plan: "stored locally beside the workbooks, with IDs never reused,
    // like workbook IDs". As with workbooks there is no file listing them: deleting the folder
    // deletes the device, and every query reads the folders on disk.
    //
    // Ids come from the root's own counter (WorklogManager.AllocateNextDeviceId), so a deleted
    // device's number never comes back - a repair still naming it reads as on no device rather than
    // on whatever device took the number next.
    // ###########################################################################################
    public static class OwnedDeviceStore
    {
        private const string DeviceFolderPrefix = "device_";

        private static readonly JsonSerializerOptions SaveJsonOptions = new() { WriteIndented = true };

        public static string BuildDeviceFolderName(int deviceId) =>
            $"{DeviceFolderPrefix}{deviceId.ToString(CultureInfo.InvariantCulture)}";

        // Reads the id out of a "device_{id}" folder name - anything else is no device.
        public static bool TryParseDeviceFolderName(string? folderName, out int deviceId)
        {
            deviceId = 0;

            return folderName != null &&
                   folderName.StartsWith(DeviceFolderPrefix, StringComparison.Ordinal) &&
                   int.TryParse(folderName[DeviceFolderPrefix.Length..], NumberStyles.None, CultureInfo.InvariantCulture, out deviceId) &&
                   deviceId > 0;
        }

        private static string Root => WorklogManager.WorkbookRoot;

        // The device's own folder, or null when it does not exist (or there is no usable root).
        public static string? GetFolder(int deviceId)
        {
            if (string.IsNullOrEmpty(Root))
            {
                return null;
            }

            string folder = Path.Combine(Root, BuildDeviceFolderName(deviceId));
            return Directory.Exists(folder) ? folder : null;
        }

        // ###########################################################################################
        // Every device on disk, by name (capitals aside), then by id. A folder that is no device or
        // whose index.json cannot be read is skipped - debris, or a device being deleted.
        // ###########################################################################################
        public static List<OwnedDeviceRecord> GetAll()
        {
            var devices = new List<OwnedDeviceRecord>();

            if (string.IsNullOrEmpty(Root) || !Directory.Exists(Root))
            {
                return devices;
            }

            foreach (string folder in Directory.GetDirectories(Root))
            {
                if (TryParseDeviceFolderName(Path.GetFileName(folder), out int id) && Read(folder) is { } device && device.Id == id)
                {
                    devices.Add(device);
                }
            }

            return devices
                .OrderBy(device => device.Name, StringComparer.OrdinalIgnoreCase)
                .ThenBy(device => device.Id)
                .ToList();
        }

        public static OwnedDeviceRecord? Get(int deviceId) =>
            GetFolder(deviceId) is { } folder && Read(folder) is { } device && device.Id == deviceId ? device : null;

        // ###########################################################################################
        // A new device, from the form - null when it has no name or nothing could be written. Its id
        // is allocated (and spent) first; its status Unknown unless the form said otherwise. `comments`
        // and `links`: those the form held for it, written with it in its first write.
        // ###########################################################################################
        public static OwnedDeviceRecord? Create(
            OwnedDeviceDetails details,
            IReadOnlyList<WorklogCommentRecord>? comments = null,
            IReadOnlyList<WorklogLinkRecord>? links = null)
        {
            if (string.IsNullOrEmpty(Root) || string.IsNullOrWhiteSpace(details.Name))
            {
                Logger.Warning("Failed to add a device: no name, or no usable workbook root folder");
                return null;
            }

            int id = WorklogManager.AllocateNextDeviceId();
            string folder = Path.Combine(Root, BuildDeviceFolderName(id));

            // Sorted by hand, a new device comes last: after the last one placed - or, with none placed yet, unplaced
            // itself, which the order of adding puts last.
            int lastPlace = GetAll().Select(each => each.DisplayOrder).DefaultIfEmpty(0).Max();

            var device = new OwnedDeviceRecord
            {
                Id = id,
                CreatedDate = DateTime.Now.Date,
                DisplayOrder = lastPlace > 0 ? lastPlace + 1 : 0,
            };

            Apply(device, details);

            foreach (WorklogCommentRecord comment in comments ?? [])
            {
                if (RichNote.TrimMarkup(comment.Text) is { Length: > 0 } text)
                {
                    device.Comments.Add(new WorklogCommentRecord { Id = NextCommentId(device), Text = text, Date = comment.Date });
                }
            }

            foreach (WorklogLinkRecord link in links ?? [])
            {
                if (link.Url?.Trim() is { Length: > 0 } url)
                {
                    device.Links.Add(new WorklogLinkRecord { Id = NextLinkId(device), Headline = link.Headline?.Trim() ?? string.Empty, Url = url });
                }
            }

            try
            {
                Directory.CreateDirectory(folder);
            }
            catch (Exception ex)
            {
                Logger.Warning($"Failed to create device folder: [{folder}] [{ex.Message}]");
                return null;
            }

            if (!Save(folder, device))
            {
                return null;
            }

            Logger.Info($"Setting changed: [Devices] added device [#{id}] [{device.Name}] on board [{device.BoardKey}]");
            return device;
        }

        // ###########################################################################################
        // The form's fields over a device, re-read first - a board swapped is simply its new board,
        // its repairs keeping the board they were made on. Null when the device cannot be read, has
        // no name, or the write failed.
        // ###########################################################################################
        //
        // A DEVICE WITH WORKBOOKS KEEPS ITS BOARD (owner decision, 2026-11-06, asked: "Locked while it has workbooks" -
        // a workbook is associated only to a device of its own board): the form locks the boxes, and a board handed in
        // all the same is not taken.
        public static OwnedDeviceRecord? Update(int deviceId, OwnedDeviceDetails details)
        {
            if (string.IsNullOrWhiteSpace(details.Name))
            {
                return null;
            }

            return Change(deviceId, device =>
            {
                string board = device.BoardKey;
                Apply(device, details);

                if (!string.Equals(board.Trim(), device.BoardKey, StringComparison.OrdinalIgnoreCase) &&
                    WorklogManager.GetAllWorkbooks().Any(workbook => workbook.OwnedDeviceId == deviceId))
                {
                    Logger.Warning($"Device [#{deviceId}] keeps its board [{board}]: it has workbooks, and [{device.BoardKey}] would take it off their board");
                    device.BoardKey = board;
                }
            });
        }

        // ###########################################################################################
        // Every device on the board `from` on the board `to` (its name changed, not the board in the
        // device - BoardKeyMove, 2026-11-01).
        // ###########################################################################################
        public static bool MoveBoard(string from, string to)
        {
            bool allMoved = true;

            foreach (OwnedDeviceRecord device in GetAll().Where(device => string.Equals(device.BoardKey, from, StringComparison.OrdinalIgnoreCase)))
            {
                allMoved &= Change(device.Id, found => found.BoardKey = to) is not null;
            }

            return allMoved;
        }

        // ###########################################################################################
        // The devices' order sorted by hand (OwnedDeviceSorting.Manual): `idsInOrder` first to last, each device's
        // place written only where it moved. False when a device could not be written - the others keep theirs.
        // ###########################################################################################
        public static bool SetOrder(IReadOnlyList<int> idsInOrder)
        {
            bool allWritten = true;

            for (int index = 0; index < idsInOrder.Count; index++)
            {
                int place = index + 1;

                if (Get(idsInOrder[index]) is { } device && device.DisplayOrder != place)
                {
                    allWritten &= Change(device.Id, found => found.DisplayOrder = place) is not null;
                }
            }

            return allWritten;
        }

        // ###########################################################################################
        // A device's description with the notes it had folded in (owner request, 2026-11-07: 'Remove the "Notes" as
        // we now have the "Description"'): the notes under the description, a blank line between, so nothing typed
        // in the box that went is lost. No notes: the description as it is.
        // ###########################################################################################
        internal static string FoldNotes(string? description, string? notes)
        {
            string kept = description ?? string.Empty;
            string folded = notes?.Trim() ?? string.Empty;

            if (folded.Length == 0)
            {
                return kept;
            }

            return kept.Trim().Length == 0 ? folded : kept.TrimEnd() + "\n\n" + folded;
        }

        // ###########################################################################################
        // A device's comments, each saved as it is added, changed or deleted (owner decision, 2026-11-06:
        // "At once, as a worklog does" - an existing device's; the form holds a new device's until it is
        // added, Create). Blank text is no comment. The saved device, or null when nothing changed on disk.
        // A comment's markup, a description's and a photo's comment are trimmed by RichNote.TrimMarkup, never Trim()
        // (code review, 2026-11-08): a first line's indent says whether it is a bullet and how deep.
        // ###########################################################################################
        public static OwnedDeviceRecord? AddComment(int deviceId, string? text, DateTime date)
        {
            if (RichNote.TrimMarkup(text) is not { Length: > 0 } trimmed)
            {
                return null;
            }

            return Change(deviceId, device => device.Comments.Add(new WorklogCommentRecord
            {
                Id = NextCommentId(device),
                Text = trimmed,
                Date = date,
            }));
        }

        public static OwnedDeviceRecord? UpdateComment(int deviceId, int commentId, string? text)
        {
            if (RichNote.TrimMarkup(text) is not { Length: > 0 } trimmed)
            {
                return null;
            }

            bool found = false;

            OwnedDeviceRecord? saved = Change(deviceId, device =>
            {
                if (device.Comments.FirstOrDefault(comment => comment.Id == commentId) is { } comment)
                {
                    comment.Text = trimmed;
                    found = true;
                }
            });

            return found ? saved : null;
        }

        public static OwnedDeviceRecord? RemoveComment(int deviceId, int commentId)
        {
            bool found = false;

            OwnedDeviceRecord? saved = Change(deviceId, device => found = device.Comments.RemoveAll(comment => comment.Id == commentId) > 0);

            return found ? saved : null;
        }

        // A comment's id: one past the highest the device has - ids are the device's own, 1 up.
        private static int NextCommentId(OwnedDeviceRecord device) =>
            RecordIds.Next(device.Comments.Select(comment => comment.Id));

        // ###########################################################################################
        // A device's links of interest, as a worklog's (owner request, 2026-11-06) - a headline and its address,
        // saved one at a time as they change, an existing device's at once. No address is no link. The saved
        // device, or null when nothing changed on disk.
        // ###########################################################################################
        public static OwnedDeviceRecord? AddLink(int deviceId, string? headline, string? url)
        {
            if (url?.Trim() is not { Length: > 0 } address)
            {
                return null;
            }

            return Change(deviceId, device => device.Links.Add(new WorklogLinkRecord
            {
                Id = NextLinkId(device),
                Headline = headline?.Trim() ?? string.Empty,
                Url = address,
            }));
        }

        public static OwnedDeviceRecord? UpdateLink(int deviceId, int linkId, string? headline, string? url)
        {
            if (url?.Trim() is not { Length: > 0 } address)
            {
                return null;
            }

            bool found = false;

            OwnedDeviceRecord? saved = Change(deviceId, device =>
            {
                if (device.Links.FirstOrDefault(link => link.Id == linkId) is { } link)
                {
                    link.Headline = headline?.Trim() ?? string.Empty;
                    link.Url = address;
                    found = true;
                }
            });

            return found ? saved : null;
        }

        public static OwnedDeviceRecord? RemoveLink(int deviceId, int linkId)
        {
            bool found = false;

            OwnedDeviceRecord? saved = Change(deviceId, device => found = device.Links.RemoveAll(link => link.Id == linkId) > 0);

            return found ? saved : null;
        }

        private static int NextLinkId(OwnedDeviceRecord device) =>
            RecordIds.Next(device.Links.Select(link => link.Id));

        // ###########################################################################################
        // The photos put in `orderedIds`' order (owner request, 2026-11-06: "I would like to be able to rearrange the
        // photos, just like it is possible to do in the worklogs") - each photo's DisplayOrder its place. An id the
        // device has not is passed over, and a photo the list leaves out keeps its place after those named. The saved
        // device, or null.
        // ###########################################################################################
        public static OwnedDeviceRecord? SetPhotoOrder(int deviceId, IReadOnlyList<int> orderedIds) =>
            Change(deviceId, device =>
            {
                List<WorklogAttachmentRecord> ordered = orderedIds
                    .Select(id => device.Photos.FirstOrDefault(photo => photo.Id == id))
                    .OfType<WorklogAttachmentRecord>()
                    .Distinct()
                    .Concat(device.Photos.OrderBy(photo => photo.DisplayOrder).ThenBy(photo => photo.Id))
                    .Distinct()
                    .ToList();

                for (int place = 0; place < ordered.Count; place++)
                {
                    ordered[place].DisplayOrder = place;
                }
            });

        // ###########################################################################################
        // Photos into the device's folder - each copied under its own file name, numbered when the
        // folder has that name already, never a hidden file (ContributionPackaging's rule) - then
        // listed on the device, with its comment, as ONE write. Returns the photos added (none when
        // nothing could be).
        // ###########################################################################################
        public static IReadOnlyList<WorklogAttachmentRecord> AddPhotos(int deviceId, IReadOnlyList<OwnedDevicePhotoSource> sources)
        {
            if (GetFolder(deviceId) is not { } folder)
            {
                return [];
            }

            var copied = new List<(string Name, string Comment)>();

            foreach (OwnedDevicePhotoSource source in sources.Where(source => ContributionPackaging.IsAddablePictureFile(source.Path)))
            {
                if (CopyIn(deviceId, folder, source.Path) is { } name)
                {
                    copied.Add((name, RichNote.TrimMarkup(source.Comment)));
                }
            }

            if (copied.Count == 0)
            {
                return [];
            }

            var added = new List<WorklogAttachmentRecord>();

            bool saved = Change(deviceId, device =>
            {
                foreach ((string name, string comment) in copied)
                {
                    var photo = new WorklogAttachmentRecord
                    {
                        Id = RecordIds.Next(device.Photos.Select(p => p.Id)),
                        FileName = name,
                        Comment = comment,
                        DisplayOrder = device.Photos.Count == 0 ? 0 : device.Photos.Max(p => p.DisplayOrder) + 1,
                    };

                    device.Photos.Add(photo);
                    added.Add(photo);
                }
            }) is not null;

            if (!saved)
            {
                // Not listed: the copies would lie in the folder citing nothing.
                foreach ((string name, _) in copied)
                {
                    TryDelete(Path.Combine(folder, name));
                }

                return [];
            }

            return added;
        }

        // ###########################################################################################
        // A photo's comment changed - and its picture too when `replacement` names one: copied in under
        // a name of its own, the photo pointed at it, and only then the old file deleted, so a failed
        // write never leaves the photo without its file. False when nothing could be saved.
        // ###########################################################################################
        public static bool UpdatePhoto(int deviceId, int photoId, string? replacement, string? comment)
        {
            if (GetFolder(deviceId) is not { } folder)
            {
                return false;
            }

            string? newName = null;

            if (!string.IsNullOrWhiteSpace(replacement))
            {
                if (!ContributionPackaging.IsAddablePictureFile(replacement) || CopyIn(deviceId, folder, replacement) is not { } copiedName)
                {
                    return false;
                }

                newName = copiedName;
            }

            string? oldName = null;
            bool found = false;

            bool saved = Change(deviceId, device =>
            {
                if (device.Photos.FirstOrDefault(photo => photo.Id == photoId) is not { } photo)
                {
                    return;
                }

                found = true;
                photo.Comment = RichNote.TrimMarkup(comment);

                if (newName is not null)
                {
                    oldName = photo.FileName;
                    photo.FileName = newName;
                }
            }) is not null && found;

            if (!saved)
            {
                if (newName is not null)
                {
                    TryDelete(Path.Combine(folder, newName));
                }

                return false;
            }

            if (oldName is not null && PhotoPath(deviceId, oldName) is { } oldPath)
            {
                TryDelete(oldPath);
            }

            return true;
        }

        // ###########################################################################################
        // One picture copied into the device's folder under a name of its own - its file name, numbered
        // when the folder has that name already. Null when it could not be. A plain name, directly in
        // the device's folder - nothing else is ever written (code review, 2026-10-28: a "\" in a name on
        // Linux took the copy out of it).
        // ###########################################################################################
        private static string? CopyIn(int deviceId, string folder, string source)
        {
            string name = ComponentPaneFileNames.ForDevicePhoto(
                source,
                Directory.GetFiles(folder).Select(Path.GetFileName).OfType<string>());
            string folderFull = Path.GetFullPath(folder);
            string destination = Path.GetFullPath(Path.Combine(folder, name));

            if (!string.Equals(Path.GetDirectoryName(destination), folderFull.TrimEnd(Path.DirectorySeparatorChar), StringComparison.Ordinal) ||
                !string.Equals(Path.GetFileName(destination), name, StringComparison.Ordinal))
            {
                Logger.Warning($"A photo for device [#{deviceId}] was not copied: its name [{name}] would leave the device's folder");
                return null;
            }

            try
            {
                File.Copy(source, destination);
                return name;
            }
            catch (Exception ex)
            {
                Logger.Warning($"Failed to copy a photo to device [#{deviceId}]: [{source}] [{ex.Message}]");
                return null;
            }
        }

        // ###########################################################################################
        // A photo taken off the device and its file deleted - the file only once the device no
        // longer lists it, so a failed write never leaves a listed photo without its file.
        // ###########################################################################################
        public static bool RemovePhoto(int deviceId, int photoId)
        {
            string? fileName = null;

            if (Change(deviceId, device =>
                {
                    fileName = device.Photos.FirstOrDefault(photo => photo.Id == photoId)?.FileName;
                    device.Photos.RemoveAll(photo => photo.Id == photoId);
                }) is null || fileName is null)
            {
                return false;
            }

            if (PhotoPath(deviceId, fileName) is { } path)
            {
                TryDelete(path);
            }

            return true;
        }

        // ###########################################################################################
        // Where a photo of the device lies - only a plain file name in the device's own folder, never
        // a path a hand-edited index.json makes climb out of it.
        // ###########################################################################################
        public static string? PhotoPath(int deviceId, string? fileName)
        {
            if (GetFolder(deviceId) is not { } folder || string.IsNullOrWhiteSpace(fileName) ||
                !string.Equals(Path.GetFileName(fileName), fileName, StringComparison.Ordinal) ||
                fileName is "." or "..")
            {
                return null;
            }

            string path = Path.Combine(folder, fileName);
            return File.Exists(path) ? path : null;
        }

        // ###########################################################################################
        // The device's folder deleted - its photos, notes and comments with it. Its repairs stay:
        // they are workbooks of their own, and read as on no device from then on.
        // ###########################################################################################
        public static bool Delete(int deviceId)
        {
            if (GetFolder(deviceId) is not { } folder)
            {
                return false;
            }

            try
            {
                Directory.Delete(folder, recursive: true);
            }
            catch (Exception ex)
            {
                Logger.Warning($"Failed to delete device [#{deviceId}]: [{ex.Message}]");
                return false;
            }

            Logger.Info($"Setting changed: [Devices] deleted device [#{deviceId}]");
            return true;
        }

        // ---- the record on disk ---------------------------------------------------------------

        private static void Apply(OwnedDeviceRecord device, OwnedDeviceDetails details)
        {
            device.Name = details.Name.Trim();
            device.BoardKey = details.BoardKey?.Trim() ?? string.Empty;
            device.Region = details.Region?.Trim() ?? string.Empty;
            device.SerialNumber = details.SerialNumber?.Trim() ?? string.Empty;
            device.Description = RichNote.TrimMarkup(details.Description);
            device.Status = OwnedDeviceStatus.Normalise(details.Status);
            device.Condition = OwnedDeviceCondition.Within(details.Condition);
        }

        // Re-reads the device, changes it and writes it back - the saved record, or null.
        private static OwnedDeviceRecord? Change(int deviceId, Action<OwnedDeviceRecord> change)
        {
            if (GetFolder(deviceId) is not { } folder || Read(folder) is not { } device || device.Id != deviceId)
            {
                Logger.Warning($"Failed to change device [#{deviceId}]: it could not be read");
                return null;
            }

            change(device);
            return Save(folder, device) ? device : null;
        }

        private static OwnedDeviceRecord? Read(string folder)
        {
            string path = Path.Combine(folder, AppConfig.WorklogIndexFileName);

            if (!File.Exists(path))
            {
                return null;
            }

            try
            {
                var device = JsonSerializer.Deserialize<OwnedDeviceRecord>(File.ReadAllText(path));

                if (device is not null)
                {
                    device.Photos ??= new List<WorklogAttachmentRecord>();
                    device.Comments ??= new List<WorklogCommentRecord>();
                    device.Comments.RemoveAll(comment => comment is null);
                    device.Links ??= new List<WorklogLinkRecord>();
                    device.Links.RemoveAll(link => link is null);

                    foreach (WorklogLinkRecord link in device.Links)
                    {
                        link.Headline ??= string.Empty;
                        link.Url ??= string.Empty;
                    }
                    device.Name ??= string.Empty;
                    device.BoardKey ??= string.Empty;
                    device.Region ??= string.Empty;
                    device.SerialNumber ??= string.Empty;
                    device.Description = FoldNotes(device.Description, device.Notes);
                    device.Notes = null;
                    device.Status = OwnedDeviceStatus.Normalise(device.Status);
                    device.Condition = OwnedDeviceCondition.Within(device.Condition);

                    foreach (WorklogCommentRecord comment in device.Comments)
                    {
                        comment.Text ??= string.Empty;
                    }
                }

                return device;
            }
            catch (Exception ex)
            {
                // An error once, not at every read of the list (UnreadableFiles).
                UnreadableFiles.Report(path, $"Failed to read device: [{folder}] [{ex.Message}] - skipped");
                return null;
            }
        }

        private static bool Save(string folder, OwnedDeviceRecord device) =>
            AtomicJsonFile.Write(Path.Combine(folder, AppConfig.WorklogIndexFileName), device, SaveJsonOptions, "device");

        private static void TryDelete(string path)
        {
            try
            {
                File.Delete(path);
            }
            catch (Exception ex)
            {
                Logger.Warning($"Failed to delete [{path}]: [{ex.Message}]");
            }
        }
    }
}
