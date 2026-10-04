using System.Collections;
using System.IO;
using System.Reflection;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// The test-only workbook cache (CachedWorkbooks) must hand out exactly the bytes a real write
// would, and never one board's bytes for another: two boards differing in ANY field must have
// different keys. Reflection walks every property of BoardData and of every entry type it holds,
// so a field added later is covered without anybody remembering this file.
// ###########################################################################################
public sealed class CachedWorkbooksTests
{
    [Fact]
    public void A_cached_write_gives_the_bytes_the_real_writer_gives()
    {
        using var workspace = new TempWorkspace();

        var board = new BoardData
        {
            HardwareName = "C64",
            BoardName = "250407",
            Components = [new ComponentEntry { BoardLabel = "U1", FriendlyName = "CPU" }]
        };

        string real = Path.Combine(workspace.Root, "real", "board.xlsx");
        string first = Path.Combine(workspace.Root, "first", "board.xlsx");
        string second = Path.Combine(workspace.Root, "second", "board.xlsx");

        BoardWorkbookWriter.Write(real, board);
        CachedWorkbooks.Write(first, board);
        CachedWorkbooks.Write(second, board);

        Assert.Equal(File.ReadAllBytes(real), File.ReadAllBytes(first));
        Assert.Equal(File.ReadAllBytes(real), File.ReadAllBytes(second));
    }

    [Fact]
    public void Two_boards_differing_in_any_field_never_share_a_key()
    {
        var unset = new List<string>();

        // BoardData's own values.
        foreach (PropertyInfo property in typeof(BoardData).GetProperties(BindingFlags.Public | BindingFlags.Instance))
        {
            if (property.PropertyType == typeof(string))
            {
                var changed = new BoardData();
                property.SetValue(changed, "changed");
                Assert.NotEqual(CachedWorkbooks.KeyOf(new BoardData()), CachedWorkbooks.KeyOf(changed));
            }
        }

        // Every field of every entry type, one at a time, on an otherwise identical board.
        foreach (PropertyInfo list in typeof(BoardData).GetProperties(BindingFlags.Public | BindingFlags.Instance)
                     .Where(property => property.PropertyType.IsGenericType && property.PropertyType.GetGenericTypeDefinition() == typeof(List<>)))
        {
            Type entryType = list.PropertyType.GetGenericArguments()[0];

            foreach (MemberInfo member in entryType.GetMembers(BindingFlags.Public | BindingFlags.Instance)
                         .Where(member => member is PropertyInfo { CanWrite: true } or FieldInfo { IsInitOnly: false }))
            {
                Type valueType = member is PropertyInfo p ? p.PropertyType : ((FieldInfo)member).FieldType;
                object? changedValue = CachedWorkbooksTests.ChangedValueFor(valueType);

                if (changedValue is null)
                {
                    unset.Add($"{entryType.Name}.{member.Name} ({valueType.Name})");
                    continue;
                }

                object plain = Activator.CreateInstance(entryType)!;
                object changed = Activator.CreateInstance(entryType)!;

                if (member is PropertyInfo property)
                    property.SetValue(changed, changedValue);
                else
                    ((FieldInfo)member).SetValue(changed, changedValue);

                Assert.NotEqual(
                    CachedWorkbooks.KeyOf(CachedWorkbooksTests.BoardWith(list, plain)),
                    CachedWorkbooks.KeyOf(CachedWorkbooksTests.BoardWith(list, changed)));
            }
        }

        // A field of a type this test cannot vary is a field it cannot vouch for - teach it.
        Assert.True(unset.Count == 0, "Add a changed value for: " + string.Join(", ", unset));
    }

    private static BoardData BoardWith(PropertyInfo list, object entry)
    {
        var board = new BoardData();
        ((IList)list.GetValue(board)!).Add(entry);
        return board;
    }

    private static object? ChangedValueFor(Type type) =>
        type == typeof(string) ? "changed" :
        type == typeof(bool) ? true :
        type == typeof(int) ? 7 :
        type == typeof(long) ? 7L :
        type == typeof(double) ? 7.5 :
        null;
}
