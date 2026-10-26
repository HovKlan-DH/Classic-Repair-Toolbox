using CRT;
using Handlers.DataHandling;
using Handlers.IcTesting;
using Xunit;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// The data folders CRT reads BY NAME, rather than through a board file, must be the ones the
// data tree keeps (DataTreeUsage, 2026-09-25; the 4.0 tree's DataTreeLayout, 2026-10-24).
//
// No board cites the MiniPro IC tests; the 4.0 tree keeps them only because it reserves their
// top folder (DataTreeLayout.ReservedTopFolders), and the converter carries them there. If the
// app moved the folder and the tree did not follow, every file in it would be missing from
// every user's download. So the folder is compared here, app side against tree side, and the
// test fails on either moving alone.
// ###########################################################################################
[Collection("DataManager")]
public sealed class DataFoldersReadByNameTests
{
    [Fact]
    public void The_MiniPro_folder_the_app_reads_is_one_the_tree_keeps()
    {
        string root = DataManager.DataRoot;
        string catalogue = string.IsNullOrEmpty(root)
            ? IcTestCatalogue.DefaultCatalogueDir
            : Path.GetRelativePath(root, IcTestCatalogue.DefaultCatalogueDir);

        catalogue = catalogue.Replace(Path.DirectorySeparatorChar, '/');

        Assert.StartsWith(DataTreeLayout.MiniProTestsFolder + "/", catalogue, StringComparison.Ordinal);
        Assert.Contains(DataTreeLayout.MiniProTestsFolder.Split('/')[0], DataTreeLayout.ReservedTopFolders);
    }

    [Fact]
    public void The_KiCad_folder_the_app_reads_is_the_one_the_rule_keeps()
    {
        Assert.Equal(AppConfig.KiCadDataFolderName, DataTreeUsage.KiCadFolderName);
    }
}
