using CRT;
using Handlers.DataHandling;
using Handlers.IcTesting;
using Xunit;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// The data folders CRT reads BY NAME, rather than through a workbook, must be the ones the
// server's "unused files" rule keeps (DataTreeUsage, 2026-09-25).
//
// No workbook cites the MiniPro IC tests or a board's KiCad files; the rule keeps them only
// because it is told CRT reads those folders. If the app moved one and the rule did not follow,
// the server would call every file in it unused - and remove it from every user's download. So
// each folder is compared here, app side against rule side, and the test fails on either moving
// alone.
// ###########################################################################################
[Collection("DataManager")]
public sealed class DataFoldersReadByNameTests
{
    [Fact]
    public void The_MiniPro_folder_the_app_reads_is_one_the_rule_keeps()
    {
        string root = DataManager.DataRoot;
        string catalogue = string.IsNullOrEmpty(root)
            ? IcTestCatalogue.DefaultCatalogueDir
            : Path.GetRelativePath(root, IcTestCatalogue.DefaultCatalogueDir);

        catalogue = catalogue.Replace(Path.DirectorySeparatorChar, '/');

        Assert.Contains(
            DataTreeUsage.FoldersReadByName,
            folder => catalogue.StartsWith(folder + "/", StringComparison.Ordinal));
    }

    [Fact]
    public void The_KiCad_folder_the_app_reads_is_the_one_the_rule_keeps()
    {
        Assert.Equal(AppConfig.KiCadDataFolderName, DataTreeUsage.KiCadFolderName);
    }
}
