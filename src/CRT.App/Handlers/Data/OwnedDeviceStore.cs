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
    // the chip type). Its name, its hardware and board, region, serial number, where it came from,
    // notes, photos, its status, and the changes made to it. Its repairs are workbooks naming it
    // (WorkbookRecord.OwnedDeviceId), never a list here - a workbook deleted leaves nothing behind.
    //
    // Local data, like workbooks: never synced, never sent anywhere.
    // ###########################################################################################
    public sealed class OwnedDeviceRecord
    {
        [JsonPropertyName("id")] public int Id { get; set; }
        [JsonPropertyName("name")] public string Name { get; set; } = string.Empty;

        // "{hardware}|{board}", the shape a workbook's BoardKey has (Main.GetCurrentBoardKey).
        [JsonPropertyName("boardKey")] public string BoardKey { get; set; } = string.Empty;
        [JsonPropertyName("region")] public string Region { get; set; } = string.Empty;
        [JsonPropertyName("serialNumber")] public string SerialNumber { get; set; } = string.Empty;
        [JsonPropertyName("origin")] public string Origin { get; set; } = string.Empty;
        [JsonPropertyName("notes")] public string Notes { get; set; } = string.Empty;

        // One of OwnedDeviceStatus' words, stored as the word.
        [JsonPropertyName("status")] public string Status { get; set; } = OwnedDeviceStatus.Working;
        [JsonPropertyName("createdDate")] public DateTime CreatedDate { get; set; }

        // The photos lie in the device's own folder under their FileName, as a worklog's do.
        [JsonPropertyName("photos")] public List<WorklogAttachmentRecord> Photos { get; set; } = new();

        // What was done to it outside a repair: a power supply replaced, a ROM fitted, a board swapped.
        [JsonPropertyName("changes")] public List<OwnedDeviceChangeRecord> Changes { get; set; } = new();

        // Whatever a later CRT writes that this one does not know - kept and written back, so
        // opening a device in an older CRT loses nothing.
        [JsonExtensionData] public Dictionary<string, JsonElement>? Unknown { get; set; }
    }

    // ###########################################################################################
    // One change made to a device ("2025-06 JiffyDOS ROM fitted"). Ids are unique within the device.
    // ###########################################################################################
    public sealed class OwnedDeviceChangeRecord
    {
        [JsonPropertyName("id")] public int Id { get; set; }
        [JsonPropertyName("date")] public DateTime Date { get; set; }
        [JsonPropertyName("text")] public string Text { get; set; } = string.Empty;
    }

    // ###########################################################################################
    // What the "Add device" form collects - every field of a device the form edits.
    // ###########################################################################################
    public sealed record OwnedDeviceDetails(
        string Name,
        string BoardKey,
        string Region,
        string SerialNumber,
        string Origin,
        string Status);

    // ###########################################################################################
    // THE WORDS A DEVICE'S STATUS IS STORED AS - the plan's five, in the order they are offered. A
    // status this CRT does not know (a later one's) is kept as it is, never forced into one of these.
    // ###########################################################################################
    public static class OwnedDeviceStatus
    {
        public const string Working = "Working";
        public const string Faulty = "Faulty";
        public const string InRepair = "In repair";
        public const string ForParts = "For parts";
        public const string Gone = "Not mine any more";

        public static readonly IReadOnlyList<string> All = [Working, Faulty, InRepair, ForParts, Gone];

        // The known word in its own spelling, a blank one Working, an unknown one as it is.
        public static string Normalise(string? status)
        {
            string trimmed = status?.Trim() ?? string.Empty;

            if (trimmed.Length == 0)
            {
                return Working;
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
        // is allocated (and spent) first; its status Working unless the form said otherwise.
        // ###########################################################################################
        public static OwnedDeviceRecord? Create(OwnedDeviceDetails details)
        {
            if (string.IsNullOrEmpty(Root) || string.IsNullOrWhiteSpace(details.Name))
            {
                Logger.Warning("Failed to add a device: no name, or no usable workbook root folder");
                return null;
            }

            int id = WorklogManager.AllocateNextDeviceId();
            string folder = Path.Combine(Root, BuildDeviceFolderName(id));

            var device = new OwnedDeviceRecord
            {
                Id = id,
                CreatedDate = DateTime.Now.Date,
            };

            Apply(device, details);

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
        // The form's fields over a device, re-read first. A board changed is RECORDED as a change of
        // the device (the plan: "A machine whose board is swapped records it as a change") - its
        // repairs keep the board they were made on. Null when the device cannot be read, has no name,
        // or the write failed.
        // ###########################################################################################
        public static OwnedDeviceRecord? Update(int deviceId, OwnedDeviceDetails details)
        {
            if (string.IsNullOrWhiteSpace(details.Name))
            {
                return null;
            }

            return Change(deviceId, device =>
            {
                string oldBoard = device.BoardKey;
                Apply(device, details);

                if (!string.Equals(oldBoard, device.BoardKey, StringComparison.OrdinalIgnoreCase) && oldBoard.Length > 0)
                {
                    AddChangeTo(device, DateTime.Now.Date, OwnedDeviceWording.BoardSwapped(oldBoard, device.BoardKey));
                }
            });
        }

        public static OwnedDeviceRecord? SetNotes(int deviceId, string? notes) =>
            Change(deviceId, device => device.Notes = notes?.Trim() ?? string.Empty);

        // ###########################################################################################
        // A change made to the device - blank text adds nothing (null).
        // ###########################################################################################
        public static OwnedDeviceChangeRecord? AddChange(int deviceId, DateTime date, string? text)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                return null;
            }

            OwnedDeviceChangeRecord? added = null;
            return Change(deviceId, device => added = AddChangeTo(device, date.Date, text.Trim())) is null ? null : added;
        }

        public static OwnedDeviceRecord? RemoveChange(int deviceId, int changeId) =>
            Change(deviceId, device => device.Changes.RemoveAll(change => change.Id == changeId));

        // ###########################################################################################
        // Photos into the device's folder - each copied under its own file name, numbered when the
        // folder has that name already, never a hidden file (ContributionPackaging's rule) - then
        // listed on the device as ONE write. Returns the photos added (none when nothing could be).
        // ###########################################################################################
        public static IReadOnlyList<WorklogAttachmentRecord> AddPhotos(int deviceId, IReadOnlyList<string> paths)
        {
            if (GetFolder(deviceId) is not { } folder)
            {
                return [];
            }

            var copied = new List<string>();

            string folderFull = Path.GetFullPath(folder);

            foreach (string source in paths.Where(ContributionPackaging.IsAddablePictureFile))
            {
                string name = ComponentPaneFileNames.ForDevicePhoto(
                    source,
                    Directory.GetFiles(folder).Select(Path.GetFileName).OfType<string>());
                string destination = Path.GetFullPath(Path.Combine(folder, name));

                // A plain name, directly in the device's folder - nothing else is ever written
                // (code review, 2026-10-28: a "\" in a name on Linux took the copy out of it).
                if (!string.Equals(Path.GetDirectoryName(destination), folderFull.TrimEnd(Path.DirectorySeparatorChar), StringComparison.Ordinal) ||
                    !string.Equals(Path.GetFileName(destination), name, StringComparison.Ordinal))
                {
                    Logger.Warning($"A photo for device [#{deviceId}] was not copied: its name [{name}] would leave the device's folder");
                    continue;
                }

                try
                {
                    File.Copy(source, destination);
                    copied.Add(name);
                }
                catch (Exception ex)
                {
                    Logger.Warning($"Failed to copy a photo to device [#{deviceId}]: [{source}] [{ex.Message}]");
                }
            }

            if (copied.Count == 0)
            {
                return [];
            }

            var added = new List<WorklogAttachmentRecord>();

            bool saved = Change(deviceId, device =>
            {
                foreach (string name in copied)
                {
                    var photo = new WorklogAttachmentRecord
                    {
                        Id = device.Photos.Count == 0 ? 1 : device.Photos.Max(p => p.Id) + 1,
                        FileName = name,
                        DisplayOrder = device.Photos.Count == 0 ? 0 : device.Photos.Max(p => p.DisplayOrder) + 1,
                    };

                    device.Photos.Add(photo);
                    added.Add(photo);
                }
            }) is not null;

            if (!saved)
            {
                // Not listed: the copies would lie in the folder citing nothing.
                foreach (string name in copied)
                {
                    TryDelete(Path.Combine(folder, name));
                }

                return [];
            }

            return added;
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
        // The device's folder deleted - its photos, notes and changes with it. Its repairs stay:
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
            device.Origin = details.Origin?.Trim() ?? string.Empty;
            device.Status = OwnedDeviceStatus.Normalise(details.Status);
        }

        private static OwnedDeviceChangeRecord AddChangeTo(OwnedDeviceRecord device, DateTime date, string text)
        {
            var change = new OwnedDeviceChangeRecord
            {
                Id = device.Changes.Count == 0 ? 1 : device.Changes.Max(c => c.Id) + 1,
                Date = date,
                Text = text,
            };

            device.Changes.Add(change);
            return change;
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
                    device.Changes ??= new List<OwnedDeviceChangeRecord>();
                    device.Name ??= string.Empty;
                    device.BoardKey ??= string.Empty;
                    device.Region ??= string.Empty;
                    device.SerialNumber ??= string.Empty;
                    device.Origin ??= string.Empty;
                    device.Notes ??= string.Empty;
                    device.Status = OwnedDeviceStatus.Normalise(device.Status);
                }

                return device;
            }
            catch (Exception ex)
            {
                Logger.Warning($"Failed to read device: [{folder}] [{ex.Message}] - skipped");
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
