using System;
using System.Linq;
using Avalonia.Controls;
using Handlers.DataHandling;
using CRT;

namespace ClassicRepairToolbox.Tests.Ui;

// The Configuration tab's "Currency used for worklog costs" drop-down.
//
// Unlike the scope radio group beside it, this list is BUILT IN CODE from WorklogCurrency.Options
// rather than declared in the markup, so nothing in the XAML says what it contains - the wiring is
// the only thing that puts rows in it, and it is what these tests pin. WorklogCurrencyTests already
// covers the table itself (its ordering, its codes, its fallbacks); what is here is that the tab
// actually shows that table, opens on the stored choice, and writes the CODE back when the user
// picks a country.
//
// COLLECTION NOTE: "HeadlessUi" rather than "UserSettings", for the same reason
// ConfigurationWorkbooksScopeTests documents - these construct a control and so need the shared
// dispatcher thread, and a class can only join one collection. They drive UserSettings' static
// state, which is safe because xunit.runner.json turns collection parallelism off. Every test
// restores WorklogCurrencyCode in a finally block.
[Collection("HeadlessUi")]
public sealed class ConfigurationCurrencyTests
{
    // Counts WorklogCurrencyChanged while body runs, and unsubscribes however it ends - a leaked
    // handler would keep counting into every later test in this collection.
    private static int CountCurrencyChanges(Action body)
    {
        int changes = 0;
        Action handler = () => changes++;

        UserSettings.WorklogCurrencyChanged += handler;
        try
        {
            body();
        }
        finally
        {
            UserSettings.WorklogCurrencyChanged -= handler;
        }

        return changes;
    }

    // The drop-down is populated at all - the failure this guards against is the whole list being
    // empty, which the markup alone cannot reveal since the rows come from code.
    [Fact]
    public void The_drop_down_offers_every_country()
    {
        UiTest.Run(() =>
        {
            var tab = new TabConfiguration();
            var combo = tab.GetControl<ComboBox>("WorklogCurrencyComboBox");

            var options = combo.ItemsSource!.Cast<WorklogCurrency.Option>().ToList();

            Assert.Equal(WorklogCurrency.Options.Count, options.Count);
            Assert.Contains(options, o => o.Country == "Denmark" && o.Code == "DKK");
            Assert.Contains(options, o => o.Country == "United States" && o.Code == "USD");
        });
    }

    // Each row reads "Country (CODE)" - the country is how the user picks, the code is what will be
    // printed on their worklogs and their exported PDF, so both have to be visible at the moment of
    // choosing. Asserted through the ComboBox's own DisplayMemberBinding rather than by calling
    // DisplayName directly, so a binding pointing at the wrong property fails here.
    [Fact]
    public void A_row_shows_the_country_and_the_code_it_will_print()
    {
        UiTest.Run(() =>
        {
            var tab = new TabConfiguration();
            var combo = tab.GetControl<ComboBox>("WorklogCurrencyComboBox");

            var binding = Assert.IsType<Avalonia.Data.Binding>(combo.DisplayMemberBinding);

            Assert.Equal(nameof(WorklogCurrency.Option.DisplayName), binding.Path);
            Assert.Equal("Denmark (DKK)", new WorklogCurrency.Option("Denmark", "DKK").DisplayName);
        });
    }

    // ###########################################################################################
    // A brand-new installation opens on the United States, which is the default that was asked for.
    //
    // Worth its own test because it is NOT simply "the first row whose code is USD": Puerto Rico
    // also uses the dollar and sorts first alphabetically, so a resolver scanning by code alone
    // opens every new user's tab on Puerto Rico. That is what this originally did, and the unit
    // test that caught it is mirrored here at the surface the user actually sees.
    // ###########################################################################################
    [Fact]
    public void A_fresh_installation_opens_on_the_United_States()
    {
        string saved = UserSettings.WorklogCurrencyCode;
        try
        {
            UserSettings.WorklogCurrencyCode = WorklogCurrency.DefaultCode;

            UiTest.Run(() =>
            {
                var tab = new TabConfiguration();
                var combo = tab.GetControl<ComboBox>("WorklogCurrencyComboBox");

                var selected = Assert.IsType<WorklogCurrency.Option>(combo.SelectedItem);

                Assert.Equal("United States", selected.Country);
                Assert.Equal("USD", selected.Code);
            });
        }
        finally
        {
            UserSettings.WorklogCurrencyCode = saved;
        }
    }

    // The tab shows the PERSISTED choice when it is built, not whatever the list happens to start
    // with - otherwise reopening Configuration reports a currency the app is not actually printing.
    [Fact]
    public void The_drop_down_opens_on_the_saved_currency()
    {
        string saved = UserSettings.WorklogCurrencyCode;
        try
        {
            UserSettings.WorklogCurrencyCode = "DKK";

            UiTest.Run(() =>
            {
                var tab = new TabConfiguration();
                var combo = tab.GetControl<ComboBox>("WorklogCurrencyComboBox");

                Assert.Equal("Denmark", ((WorklogCurrency.Option)combo.SelectedItem!).Country);
            });
        }
        finally
        {
            UserSettings.WorklogCurrencyCode = saved;
        }
    }

    // Picking a country writes its CODE, and notifies once. The notification is what makes the
    // Workbooks tab reprint its figures - Main subscribes to it and refreshes through the same
    // funnel every other worklog change uses.
    [Fact]
    public void Choosing_a_country_stores_its_currency_code_and_notifies_once()
    {
        string saved = UserSettings.WorklogCurrencyCode;
        try
        {
            UserSettings.WorklogCurrencyCode = "USD";

            UiTest.Run(() =>
            {
                var tab = new TabConfiguration();
                var combo = tab.GetControl<ComboBox>("WorklogCurrencyComboBox");

                int changes = CountCurrencyChanges(() =>
                    combo.SelectedItem = WorklogCurrency.Options.Single(o => o.Country == "Norway"));

                // The CODE, not the country - the country is only how it was picked.
                Assert.Equal("NOK", UserSettings.WorklogCurrencyCode);
                Assert.Equal(1, changes);
            });
        }
        finally
        {
            UserSettings.WorklogCurrencyCode = saved;
        }
    }

    // ###########################################################################################
    // Building the tab must not itself write the setting.
    //
    // The constructor selects the stored row, which raises SelectionChanged - so a handler
    // subscribed before that assignment would write the value straight back, and every subscriber
    // would rebuild the Workbooks tab merely because the user opened Configuration. Every other
    // control on this tab is wired the same way and for the same reason.
    // ###########################################################################################
    [Fact]
    public void Opening_the_tab_writes_nothing()
    {
        string saved = UserSettings.WorklogCurrencyCode;
        try
        {
            UserSettings.WorklogCurrencyCode = "DKK";

            UiTest.Run(() =>
            {
                int changes = CountCurrencyChanges(() => _ = new TabConfiguration());

                Assert.Equal(0, changes);
                Assert.Equal("DKK", UserSettings.WorklogCurrencyCode);
            });
        }
        finally
        {
            UserSettings.WorklogCurrencyCode = saved;
        }
    }

    // ###########################################################################################
    // The one field in the app where a cost is TYPED names the currency it will be recorded in.
    //
    // This is where the setting earns most of its keep: every other surface DISPLAYS a figure the
    // user already knows, while this one asks for a number with no indication of what it will be
    // treated as - and a rate entered under the wrong assumption is what makes an exported invoice
    // wrong. Set from code rather than the markup, because a literal in the XAML would be right for
    // one user only, which is exactly the wiring this pins.
    // ###########################################################################################
    [Fact]
    public void The_work_done_dialog_names_the_currency_on_its_cost_field()
    {
        string saved = UserSettings.WorklogCurrencyCode;
        try
        {
            UserSettings.WorklogCurrencyCode = "DKK";

            UiTest.Run(() =>
            {
                var dialog = new WorklogAddWorkDoneWindow();

                Assert.Equal("Cost (DKK)", dialog.GetControl<TextBlock>("CostLabelText").Text);
            });

            UserSettings.WorklogCurrencyCode = "JPY";

            UiTest.Run(() =>
            {
                var dialog = new WorklogAddWorkDoneWindow();

                Assert.Equal("Cost (JPY)", dialog.GetControl<TextBlock>("CostLabelText").Text);
            });
        }
        finally
        {
            UserSettings.WorklogCurrencyCode = saved;
        }
    }

    // Picking a DIFFERENT country that shares the current code is not a change worth a rebuild:
    // sixteen countries here use the euro, and switching between two of them prints the same thing.
    [Fact]
    public void Choosing_another_country_with_the_same_currency_notifies_nobody()
    {
        string saved = UserSettings.WorklogCurrencyCode;
        try
        {
            UserSettings.WorklogCurrencyCode = "EUR";

            UiTest.Run(() =>
            {
                var tab = new TabConfiguration();
                var combo = tab.GetControl<ComboBox>("WorklogCurrencyComboBox");

                // The tab opened on Andorra (the first euro country); Germany is a different row
                // carrying the same code.
                int changes = CountCurrencyChanges(() =>
                    combo.SelectedItem = WorklogCurrency.Options.Single(o => o.Country == "Germany"));

                Assert.Equal("EUR", UserSettings.WorklogCurrencyCode);
                Assert.Equal(0, changes);
            });
        }
        finally
        {
            UserSettings.WorklogCurrencyCode = saved;
        }
    }
}
