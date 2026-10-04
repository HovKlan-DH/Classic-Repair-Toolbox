using Avalonia.Controls;
using CRT;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// The two Configuration tab check boxes that other texts tell a contributor to tick - CRT's "try it
// in BETA" notice and the server's "accepted into BETA" mail (2026-10-03). Their labels are
// CRT.Data's ConfigurationWording, which the tab's markup reads through x:Static; this holds the
// check boxes on screen to it, so a label typed back into the markup - and then changed there only -
// fails here instead of sending people looking for a box that is not there.
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class ConfigurationWordingTests
{
    [Fact]
    public void The_check_boxes_other_texts_name_carry_the_shared_labels()
    {
        UiTest.Run(() =>
        {
            var tab = new TabConfiguration();

            Assert.Equal(ConfigurationWording.BetaSourceCheckBox, tab.GetControl<CheckBox>("DownloadDataFromTestSourceCheckBox").Content);
            Assert.Equal(ConfigurationWording.CheckDataOnLaunchCheckBox, tab.GetControl<CheckBox>("CheckDataOnLaunchCheckBox").Content);
        });
    }
}
