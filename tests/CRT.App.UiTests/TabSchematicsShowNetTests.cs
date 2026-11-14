using CRT;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// TabSchematics.ShowOnlyNet - the Contribute tab's "Show on board" on an important signal (CRT 4.0
// step 2: "click-to-light-up"): only that net lit, found by the Important signals panel's own rule,
// and - while the board's KiCad data is still loading, whose load lets go of every lit net - lit once
// it has settled.
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class TabSchematicsShowNetTests
{
    private static KiCadProjectBundle Project() => new()
    {
        Root = new KiCadProjectRoot
        {
            Pcb =
            [
                new KiCadPcb
                {
                    Nets = new KiCadPcbNets
                    {
                        List =
                        [
                            new KiCadNetRef { Id = "1", Name = "/CPU/PHI2", NormalizedName = "PHI2" },
                            new KiCadNetRef { Id = "2", Name = "/CPU/~{RESET}", NormalizedName = "~{RESET}" },
                            new KiCadNetRef { Id = "3", Name = "/RAM/CAS", NormalizedName = "CAS" },
                        ],
                    },
                    Footprints =
                    [
                        new KiCadPcbFootprint
                        {
                            Reference = "U9",
                            Pads = [new KiCadPcbPad { Number = "1", Net = new KiCadNetRef { Id = "3", Name = "/RAM/CAS", NormalizedName = "CAS" } }],
                        },
                    ],
                },
            ],
        },
    };

    // ###########################################################################################
    // The "KiCad nets" panel's "Connected to:" line as drawn (owner request, 2026-11-13: "[<b>U17</b> | 74LS373] pin
    // <b>5</b>"): the template's runs read the line's pieces, the label and the pin bold and nothing else - and no space
    // creeps in between the runs, which markup whitespace between inlines would add.
    // ###########################################################################################
    [Fact]
    public void A_connected_to_line_draws_the_label_and_the_pin_bold()
    {
        UiTest.Run(() =>
        {
            var tab = new TabSchematics();
            var template = Avalonia.Controls.ControlExtensions.GetControl<Avalonia.Controls.ItemsControl>(tab, "KiCadNetConnectionsList").ItemTemplate!;

            var line = new KiCadNetConnectionLine("U17", " | 74LS373", "5", string.Empty);
            var block = Assert.IsType<Avalonia.Controls.TextBlock>(template.Build(line));
            block.DataContext = line;

            var window = new Avalonia.Controls.Window { Content = block };
            window.Show();

            var runs = block.Inlines!.OfType<Avalonia.Controls.Documents.Run>().ToList();

            Assert.Equal("[U17 | 74LS373] pin 5", string.Concat(runs.Select(run => run.Text)));
            Assert.Equal(["U17", "5"], runs.Where(run => run.FontWeight == Avalonia.Media.FontWeight.Bold).Select(run => run.Text));

            window.Close();
        });
    }

    // ###########################################################################################
    // ONLY that net: a component selected on the board - its highlight and its nets lit - is let go
    // of too (code review, 2026-10-26: one selected only on the schematic kept its nets lit beside the
    // net asked for, since clearing the list's selection changed nothing).
    // ###########################################################################################
    [Fact]
    public void A_component_selected_before_lights_nothing_once_a_net_is_shown()
    {
        UiTest.Run(() =>
        {
            var tab = new TabSchematics();
            tab.UseKiCadProjectForTests(Project());
            tab.KiCadSettled(hasProject: true);

            tab.UpdateHighlightsForComponents(["U9"]);
            Assert.Equal(["CAS"], tab.SelectedKiCadNetNamesForTests);

            tab.ShowOnlyNet("PHI2");

            Assert.Empty(tab.SelectedKiCadNetNamesForTests);
            Assert.Equal(["PHI2"], tab.LockedKiCadNetNamesForTests);
        });
    }

    [Fact]
    public void Only_the_net_asked_for_is_lit_found_by_its_own_or_its_normalised_name()
    {
        UiTest.Run(() =>
        {
            var tab = new TabSchematics();
            tab.UseKiCadProjectForTests(Project());
            tab.KiCadSettled(hasProject: true);

            tab.ShowOnlyNet(" /CPU/PHI2 ");
            Assert.Equal(["PHI2"], tab.LockedKiCadNetNamesForTests);

            // Another net asked for lets go of the first.
            tab.ShowOnlyNet("~{reset}");
            Assert.Equal(["~{RESET}"], tab.LockedKiCadNetNamesForTests);

            // A net the KiCad data does not have lights up nothing - and leaves what is lit.
            tab.ShowOnlyNet("NOT A NET");
            Assert.Equal(["~{RESET}"], tab.LockedKiCadNetNamesForTests);
        });
    }

    // ###########################################################################################
    // Asked for while the KiCad data is loading, the net waits for it - and a board load before it
    // settles drops the request: the net was another board's.
    // ###########################################################################################
    [Fact]
    public void A_net_asked_for_while_the_KiCad_data_loads_is_lit_once_it_has_settled()
    {
        UiTest.Run(() =>
        {
            var tab = new TabSchematics();

            tab.ShowOnlyNet("PHI2");
            Assert.Empty(tab.LockedKiCadNetNamesForTests);

            tab.UseKiCadProjectForTests(Project());
            tab.KiCadSettled(hasProject: true);
            Assert.Equal(["PHI2"], tab.LockedKiCadNetNamesForTests);

            // Asked once, lit once: a later load settling lights nothing by itself.
            tab.ShowOnlyNet("~{RESET}");
            tab.KiCadSettled(hasProject: true);
            Assert.Equal(["~{RESET}"], tab.LockedKiCadNetNamesForTests);
        });
    }
}
