using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Threading;
using CRT;
using Handlers.DataHandling;
using Handlers.Oscilloscope;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// "My scope is not listed" on the Oscilloscope tab (TabOscilloscope.ModelEditor.cs - CRT 4.0 step
// 5, the plan: "start from a similar model and edit each command with Try it ... Your own models are
// kept outside the synced data"): a model of the user's own made from the one shown, every box
// saved as it is left into the user's own file, the drop-downs following its name, and Try it saying
// what it needs. Driven through the tab's controls on a data root of two published models, with the
// own models' file in a temporary folder - never the user's.
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class OscilloscopeModelEditorTests : IDisposable
{
    private readonly TempWorkspace thisWorkspace = new();

    public OscilloscopeModelEditorTests()
    {
        UserSettings.LoadFrom(this.thisWorkspace.WriteFile("settings.json", "{}"));

        string main = MainFiles.Write(
            this.thisWorkspace.Root,
            [new MainFiles.Board("Commodore/C64/250407", "C64", "250407")],
            [Model("Rigol", "DS1000Z"), Model("Siglent", "SDS1000X-E")]);

        DataManager.LoadFrom(this.thisWorkspace.Root, main, this.OwnModelsFile);
    }

    public void Dispose()
    {
        DataManager.LoadFrom(this.thisWorkspace.Root, "does-not-exist.json");
        UserSettings.LoadFrom(this.thisWorkspace.WriteFile("settings-after.json", "{}"));
        this.thisWorkspace.Dispose();
    }

    private string OwnModelsFile => this.thisWorkspace.Path_("own-oscilloscopes.json");

    private static MainDocument.OscilloscopeRow Model(string brand, string series) => new()
    {
        Brand = brand,
        SeriesOrModel = series,
        Port = "5555",
        Identify = "*IDN?",
        Stop = ":STOP",
        DumpImage = ":DISP:DATA?",
        TimeDivList = "1ms,2ms",
        VoltsDivList = "1V,2V",
    };

    private IReadOnlyList<MainDocument.OscilloscopeRow> ModelsInTheFile() =>
        DataTreeJson.ReadMain(File.ReadAllText(this.OwnModelsFile)).Oscilloscopes;

    private static TabOscilloscope TabShowing(string vendor, string model)
    {
        var tab = new TabOscilloscope();
        tab.VendorComboBox.SelectedItem = vendor;
        tab.SeriesOrModelComboBox.SelectedItem = model;
        return tab;
    }

    private static void Click(Button button) => button.RaiseEvent(new RoutedEventArgs(Button.ClickEvent));

    private static void Leave(TextBox box) => box.RaiseEvent(new FocusChangedEventArgs(InputElement.LostFocusEvent));

    // The model made, the tab on it - as "Create model" leaves it.
    private static async Task<TabOscilloscope> TabWithOwnModelAsync()
    {
        TabOscilloscope tab = TabShowing("Rigol", "DS1000Z");
        Click(tab.ScopeNotListedButton);
        tab.NewModelNameTextBox.Text = "DS1054Z";
        await tab.CreateModelAsync();
        return tab;
    }

    [Fact]
    public async Task My_scope_is_not_listed_makes_a_model_of_the_user_s_own_from_the_one_shown()
    {
        await UiTest.RunAsync(async () =>
        {
            TabOscilloscope tab = TabShowing("Rigol", "DS1000Z");

            Assert.False(tab.EditMyModelButton.IsVisible);
            Click(tab.ScopeNotListedButton);

            // The model shown is the one to start from, its vendor the new model's.
            Assert.True(tab.IsModelEditorOpen);
            Assert.True(tab.NewModelPanel.IsVisible);
            Assert.Equal("Rigol DS1000Z", tab.StartFromComboBox.SelectedItem);
            Assert.Equal("Rigol", tab.NewModelVendorTextBox.Text);

            tab.NewModelNameTextBox.Text = "DS1054Z";
            await tab.CreateModelAsync();

            MainDocument.OscilloscopeRow mine = Assert.Single(this.ModelsInTheFile());
            Assert.Equal("DS1054Z", mine.SeriesOrModel);
            Assert.Equal(":DISP:DATA?", mine.DumpImage);

            // Shown in the drop-downs, ready to edit.
            Assert.Equal("Rigol", tab.VendorComboBox.SelectedItem);
            Assert.Equal("DS1054Z", tab.SeriesOrModelComboBox.SelectedItem);
            Assert.True(tab.EditMyModelButton.IsVisible);
            Assert.True(tab.EditModelPanel.IsVisible);
            Assert.Equal("Rigol DS1054Z", tab.ModelEditorHeading.Text);
            Assert.Equal(("Rigol", "DS1054Z"), tab.EditedModelForTests);
        });
    }

    // Two models of one name: the second would never be reachable in the drop-downs.
    [Fact]
    public async Task A_name_another_model_has_makes_nothing()
    {
        await UiTest.RunAsync(async () =>
        {
            TabOscilloscope tab = TabShowing("Rigol", "DS1000Z");
            Click(tab.ScopeNotListedButton);
            tab.NewModelNameTextBox.Text = "ds1000z";

            await tab.CreateModelAsync();

            Assert.True(tab.NewModelProblemText.IsVisible);
            Assert.Equal(OscilloscopeModelRules.Taken("Rigol", "ds1000z"), tab.NewModelProblemText.Text);
            Assert.Empty(DataManager.OwnScopeModels);
        });
    }

    // ###########################################################################################
    // *** A FILE OF OWN MODELS THAT COULD NOT BE READ IS NEVER WRITTEN OVER (code review,
    // 2026-10-29). *** Create replaced it with the one model made; now it says why it makes nothing.
    // ###########################################################################################
    [Fact]
    public async Task Create_model_refuses_while_the_own_models_could_not_be_read()
    {
        await UiTest.RunAsync(async () =>
        {
            string broken = this.thisWorkspace.WriteFile("broken-own.json", "{ a hand edit gone wrong");
            DataManager.LoadFrom(this.thisWorkspace.Root, MainFiles.FileName, broken);

            TabOscilloscope tab = TabShowing("Rigol", "DS1000Z");
            Click(tab.ScopeNotListedButton);
            tab.NewModelNameTextBox.Text = "DS1054Z";
            await tab.CreateModelAsync();

            Assert.True(tab.NewModelProblemText.IsVisible);
            Assert.Equal(DataManager.OwnScopeModelsProblem, tab.NewModelProblemText.Text);
            Assert.Equal("{ a hand edit gone wrong", File.ReadAllText(broken));
            Assert.Empty(DataManager.OwnScopeModels);
        });
    }

    // No Save button: a box is saved as it is left, and Escape puts it back.
    [Fact]
    public async Task A_command_is_saved_as_its_box_is_left()
    {
        await UiTest.RunAsync(async () =>
        {
            TabOscilloscope tab = await TabWithOwnModelAsync();
            (TextBox stop, _) = tab.FieldForTests("stop");

            Assert.Equal(":STOP", stop.Text);
            stop.Text = ":STOP:NOW";

            // "Saved" stands there from making the model - taken away, what comes is this save's.
            tab.ModelEditorSaveStateText.Text = string.Empty;
            Leave(stop);
            await tab.LastModelSaveForTests;

            Assert.Equal(":STOP:NOW", Assert.Single(this.ModelsInTheFile()).Stop);
            Assert.Equal(":STOP:NOW", DataManager.Oscilloscopes.Single(scope => scope.IsOwn).Stop);

            // Said after the very save awaited here, queued on this thread.
            for (int attempt = 0; attempt < 3000 && tab.ModelEditorSaveStateText.Text != "Saved"; attempt++)
            {
                await Task.Delay(2);
            }

            Assert.Equal("Saved", tab.ModelEditorSaveStateText.Text);

            stop.Text = "typed by mistake";
            stop.RaiseEvent(new KeyEventArgs { RoutedEvent = InputElement.KeyDownEvent, Key = Key.Escape });
            Assert.Equal(":STOP:NOW", stop.Text);
        });
    }

    // A setting CRT could not use is held back with the reason; a list of steps is kept anyway.
    [Fact]
    public async Task A_port_CRT_cannot_use_is_held_back_and_steps_it_cannot_read_are_named()
    {
        await UiTest.RunAsync(async () =>
        {
            TabOscilloscope tab = await TabWithOwnModelAsync();

            (TextBox port, _) = tab.FieldForTests(ScopeModelFields.PortKey);
            port.Text = "port";
            Leave(port);

            Assert.Equal(ScopeModelFields.PortProblem, tab.FieldForTests(ScopeModelFields.PortKey).Problem);
            Assert.Equal("5555", DataManager.OwnScopeModels[0].Port);

            (TextBox steps, _) = tab.FieldForTests(ScopeModelFields.TimeStepsKey);
            steps.Text = "1ms,2,5ms";
            Leave(steps);
            await tab.LastModelSaveForTests;

            Assert.Equal("1ms,2,5ms", DataManager.OwnScopeModels[0].TimeDivList);
            Assert.Contains("\"2\"", tab.FieldForTests(ScopeModelFields.TimeStepsKey).Problem);
        });
    }

    // A rename keeps the editor on the model and the drop-downs on its new name.
    [Fact]
    public async Task A_renamed_model_is_followed_by_the_drop_downs()
    {
        await UiTest.RunAsync(async () =>
        {
            TabOscilloscope tab = await TabWithOwnModelAsync();

            tab.ModelNameTextBox.Text = "DS1054Z Plus";
            Leave(tab.ModelNameTextBox);
            await tab.LastModelSaveForTests;

            Assert.Equal("DS1054Z Plus", Assert.Single(this.ModelsInTheFile()).SeriesOrModel);
            Assert.Equal("DS1054Z Plus", tab.SeriesOrModelComboBox.SelectedItem);
            Assert.True(tab.IsModelEditorOpen);
            Assert.Equal(("Rigol", "DS1054Z Plus"), tab.EditedModelForTests);

            // Another model's name is refused, the boxes left as typed.
            tab.ModelVendorTextBox.Text = "Siglent";
            tab.ModelNameTextBox.Text = "SDS1000X-E";
            Leave(tab.ModelNameTextBox);

            Assert.True(tab.ModelNameProblemText.IsVisible);
            Assert.Equal("DS1054Z Plus", DataManager.OwnScopeModels[0].SeriesOrModel);
        });
    }

    [Fact]
    public async Task Delete_model_asks_once_and_takes_it_out_of_the_list()
    {
        await UiTest.RunAsync(async () =>
        {
            TabOscilloscope tab = await TabWithOwnModelAsync();

            Click(tab.DeleteModelButton);
            Assert.True(tab.DeleteModelConfirmPanel.IsVisible);
            Assert.Single(DataManager.OwnScopeModels);

            Click(tab.KeepModelButton);
            Assert.False(tab.DeleteModelConfirmPanel.IsVisible);

            Click(tab.DeleteModelButton);
            Click(tab.ConfirmDeleteModelButton);
            await tab.LastModelSaveForTests;

            Assert.Empty(this.ModelsInTheFile());

            // The delete goes on after the very save awaited here, queued on this thread - and an await
            // that finds the save already done comes back before it has run.
            for (int attempt = 0; attempt < 500 && tab.IsModelEditorOpen; attempt++)
            {
                await Task.Delay(2);
            }

            Assert.False(tab.IsModelEditorOpen);
            Assert.DoesNotContain("DS1054Z", (IEnumerable<string>)tab.SeriesOrModelComboBox.ItemsSource!);
            Assert.False(tab.EditMyModelButton.IsVisible);
        });
    }

    // The editor edits the model the drop-downs show.
    [Fact]
    public async Task Choosing_another_model_closes_the_editor_and_Edit_my_model_opens_it_again()
    {
        await UiTest.RunAsync(async () =>
        {
            TabOscilloscope tab = await TabWithOwnModelAsync();

            tab.SeriesOrModelComboBox.SelectedItem = "DS1000Z";

            Assert.False(tab.IsModelEditorOpen);
            Assert.False(tab.EditMyModelButton.IsVisible);

            tab.SeriesOrModelComboBox.SelectedItem = "DS1054Z";
            Assert.True(tab.EditMyModelButton.IsVisible);

            Click(tab.EditMyModelButton);
            Assert.True(tab.EditModelPanel.IsVisible);
            Assert.Equal(("Rigol", "DS1054Z"), tab.EditedModelForTests);
        });
    }

    // ------------------------------------------------------------------ Submit this model

    // What the send was handed, and the server's answer to give.
    private static (List<(OscilloscopeModelSubmissionRequest Request, string? Token)> Sent, Func<OscilloscopeModelSubmissionRequest, string?, Task<Handlers.Online.OscilloscopeModelSendOutcome>> Send) Server(
        Handlers.Online.OscilloscopeModelSendOutcome answer)
    {
        var sent = new List<(OscilloscopeModelSubmissionRequest, string?)>();

        return (sent, (request, token) =>
        {
            sent.Add((request, token));
            return Task.FromResult(answer);
        });
    }

    private static readonly Handlers.Online.OscilloscopeModelSendOutcome Taken =
        new(new OscilloscopeModelSubmissionAnswer(41, "Sent. The administrator looks at it."), []);

    // ###########################################################################################
    // THE SEND PANEL IS ALWAYS THERE under a model of the user's own, its "Submit" in it (owner request, 2026-11-07: "it
    // is not at all clear, that when you press the button 'Submit this model', that it will then add a new panel below
    // ... please always do show this panel and have the 'Submit' button inside of that panel") - no button opens it, and
    // sending leaves it where it is, its message emptied.
    // ###########################################################################################
    [Fact]
    public async Task Submit_this_model_sends_it_with_the_address_and_says_it_went()
    {
        await UiTest.RunAsync(async () =>
        {
            UserSettings.ContactEmail = "anna@example.com";
            TabOscilloscope tab = await TabWithOwnModelAsync();
            (var sent, var send) = Server(Taken);
            tab.SendOverrideForTests = send;

            Assert.True(tab.IsSubmitPanelShown);
            Assert.True(tab.SubmitModelPanel.IsVisible);
            Assert.Null(tab.FindControl<Button>("SubmitModelButton"));
            Assert.Equal("Submit", tab.SendModelButton.Content);
            Assert.Null(tab.FindControl<Button>("CancelSendModelButton"));
            Assert.Equal("anna@example.com", tab.SubmitModelEmailTextBox.Text);
            Assert.False(tab.SubmitModelEmailTextBox.IsReadOnly);
            Assert.Equal(TabOscilloscope.SuiteNotRunLine, tab.SubmitModelSuiteText.Text);

            tab.SubmitModelMessageTextBox.Text = "My DS1054Z, firmware 00.04.04";
            await tab.SendModelAsync();

            (OscilloscopeModelSubmissionRequest request, string? token) = Assert.Single(sent);
            Assert.Equal("Rigol", request.Vendor);
            Assert.Equal("DS1054Z", request.SeriesOrModel);
            Assert.Equal(":DISP:DATA?", request.Values!["dumpImage"]);
            Assert.Equal("anna@example.com", request.ContactEmail);
            Assert.Equal("My DS1054Z, firmware 00.04.04", request.Message);
            Assert.False(request.FullTestSuitePassed);
            Assert.Null(token);

            Assert.True(tab.IsSubmitPanelShown);
            Assert.Equal(string.Empty, tab.SubmitModelMessageTextBox.Text);
            Assert.Equal("Sent. The administrator looks at it.", tab.ModelSentText.Text);
            Assert.NotNull(UserSettings.GetOscilloscopeModelSent("rigol", "ds1054z"));
        });
    }

    // The administrator is told whether the suite went through on the model AS SENT - a change
    // since makes it not so.
    [Fact]
    public async Task The_full_test_suite_counts_only_for_the_model_as_it_was_run()
    {
        await UiTest.RunAsync(async () =>
        {
            TabOscilloscope tab = await TabWithOwnModelAsync();
            (var sent, var send) = Server(Taken);
            tab.SendOverrideForTests = send;
            UserSettings.ContactEmail = "anna@example.com";

            tab.NoteSuitePassedForTests(OscilloscopeEntry.From(DataManager.OwnScopeModels[0], isOwn: true));
            tab.ShowSubmitPanel();

            Assert.Equal(TabOscilloscope.SuitePassedLine, tab.SubmitModelSuiteText.Text);

            (TextBox stop, _) = tab.FieldForTests("stop");
            stop.Text = ":STOP:NOW";
            Leave(stop);

            Assert.Equal(TabOscilloscope.SuiteNotRunLine, tab.SubmitModelSuiteText.Text);

            await tab.SendModelAsync();
            Assert.False(Assert.Single(sent).Request.FullTestSuitePassed);
        });
    }

    // Nothing goes without an address to tell the outcome at - and the server's refusal stays said.
    [Fact]
    public async Task A_missing_address_or_a_refusal_keeps_the_panel_with_the_reason()
    {
        await UiTest.RunAsync(async () =>
        {
            TabOscilloscope tab = await TabWithOwnModelAsync();
            (var sent, var send) = Server(new Handlers.Online.OscilloscopeModelSendOutcome(null, ["This address has sent 10 oscilloscope models today."]));
            tab.SendOverrideForTests = send;

            tab.ShowSubmitPanel();
            tab.SubmitModelEmailTextBox.Text = "anna";
            await tab.SendModelAsync();

            Assert.Empty(sent);
            Assert.Equal(TabOscilloscope.BadAddress, tab.SubmitModelProblemText.Text);

            tab.SubmitModelEmailTextBox.Text = "anna@example.com";
            await tab.SendModelAsync();

            Assert.Single(sent);
            Assert.True(tab.SubmitModelPanel.IsVisible);
            Assert.Equal("This address has sent 10 oscilloscope models today.", tab.SubmitModelProblemText.Text);
            Assert.Null(UserSettings.GetOscilloscopeModelSent("Rigol", "DS1054Z"));
        });
    }

    // Signed in on the Maintainer tab: the account's address, read only, and the model goes with
    // the account - the Submit dialog's rule (ContactAddress). The address goes too (code review,
    // 2026-10-29): a session the server has ended meanwhile is no account there, and the model is then
    // taken as sent from that address - without it, it was refused for want of one nobody could type.
    [Fact]
    public async Task Signed_in_the_model_goes_with_the_account()
    {
        await UiTest.RunAsync(async () =>
        {
            TabOscilloscope tab = await TabWithOwnModelAsync();
            (var sent, var send) = Server(Taken);
            tab.SendOverrideForTests = send;
            tab.UseMaintainerAccount(new Handlers.MaintainerHandling.ReviewSession(
                "token-1", DateTimeOffset.UtcNow.AddDays(10), 7, "dennis@example.com", "Dennis"));

            tab.ShowSubmitPanel();

            Assert.Equal("dennis@example.com", tab.SubmitModelEmailTextBox.Text);
            Assert.True(tab.SubmitModelEmailTextBox.IsReadOnly);

            await tab.SendModelAsync();

            (OscilloscopeModelSubmissionRequest request, string? token) = Assert.Single(sent);
            Assert.Equal("dennis@example.com", request.ContactEmail);
            Assert.Equal("token-1", token);
        });
    }

    // ###########################################################################################
    // *** THE EDITOR'S OWN REFRESH KEEPS THE PORT TYPED (code review, 2026-10-29). *** It selects the
    // new model, whose port went into the TCP port box - another port than the one connected to, so
    // the session went, which the refresh exists to keep.
    // ###########################################################################################
    [Fact]
    public async Task Making_a_model_keeps_the_port_typed_for_the_scope_connected()
    {
        await UiTest.RunAsync(async () =>
        {
            TabOscilloscope tab = TabShowing("Rigol", "DS1000Z");
            tab.PortTextBox.Text = "5025";
            tab.MarkSessionEstablishedForTests();

            Click(tab.ScopeNotListedButton);
            tab.NewModelNameTextBox.Text = "DS1054Z";
            await tab.CreateModelAsync();

            Assert.Equal("DS1054Z", tab.SeriesOrModelComboBox.SelectedItem);
            Assert.Equal("5025", tab.PortTextBox.Text);
        });
    }

    // With nothing connected the port box takes the model's port, as for any model picked (code
    // review, 2026-10-29: kept, "Connect" tried the port of the model shown before).
    [Fact]
    public async Task Making_a_model_with_nothing_connected_takes_the_model_s_port()
    {
        await UiTest.RunAsync(async () =>
        {
            TabOscilloscope tab = TabShowing("Rigol", "DS1000Z");
            tab.PortTextBox.Text = "5025";

            Click(tab.ScopeNotListedButton);
            tab.NewModelNameTextBox.Text = "DS1054Z";
            await tab.CreateModelAsync();

            Assert.Equal("5555", tab.PortTextBox.Text);

            (TextBox port, _) = tab.FieldForTests("port");
            port.Text = "5026";
            Leave(port);

            Assert.Equal("5026", tab.PortTextBox.Text);
        });
    }

    // The suite going through while the send panel is open says so there at once (code review,
    // 2026-10-29: the panel kept "has not gone through" in the warning colour).
    [Fact]
    public async Task The_send_panel_says_the_suite_went_through_as_soon_as_it_has()
    {
        await UiTest.RunAsync(async () =>
        {
            TabOscilloscope tab = await TabWithOwnModelAsync();
            tab.ShowSubmitPanel();
            Assert.Equal(TabOscilloscope.SuiteNotRunLine, tab.SubmitModelSuiteText.Text);

            tab.NoteSuitePassedForTests(OscilloscopeEntry.From(DataManager.OwnScopeModels[0], isOwn: true));

            Assert.Equal(TabOscilloscope.SuitePassedLine, tab.SubmitModelSuiteText.Text);
            Assert.DoesNotContain("FieldProblem", tab.SubmitModelSuiteText.Classes);
        });
    }

    // ###########################################################################################
    // *** A DATA CHECK THAT RENAMES THE MODEL OPEN IN THE EDITOR IS FOLLOWED (code review, 2026-10-29). ***
    // A model published under its name with other commands names the user's own apart as "(my
    // version)" - the editor stayed on the old name, found no own model of it, and dropped every
    // change without a word. Now it follows the model, and a change is saved into it.
    // ###########################################################################################
    [Fact]
    public async Task A_data_check_that_renames_the_model_being_edited_is_followed()
    {
        await UiTest.RunAsync(async () =>
        {
            TabOscilloscope tab = await TabWithOwnModelAsync();
            MainDocument.OscilloscopeRow published = Model("Rigol", "DS1054Z");
            published.DumpImage = ":DISP:DATA? PNG";
            this.ReloadPublishing(published);

            tab.FollowScopeModelsReloaded();

            Assert.Equal(("Rigol", "DS1054Z (my version)"), tab.EditedModelForTests);
            Assert.Equal("DS1054Z (my version)", tab.SeriesOrModelComboBox.SelectedItem);
            Assert.Contains("(my version)", tab.ModelEditorSaveStateText.Text);

            (TextBox stop, _) = tab.FieldForTests("stop");
            stop.Text = ":STOP:NOW";
            Leave(stop);
            await tab.LastModelSaveForTests;

            Assert.Equal(":STOP:NOW", Assert.Single(this.ModelsInTheFile()).Stop);
        });
    }

    // Published just as it is, the own copy goes - the editor closes and the published one is shown.
    [Fact]
    public async Task A_data_check_that_publishes_the_model_being_edited_as_it_is_closes_the_editor()
    {
        await UiTest.RunAsync(async () =>
        {
            TabOscilloscope tab = await TabWithOwnModelAsync();
            this.ReloadPublishing(Assert.Single(DataManager.OwnScopeModels));

            tab.FollowScopeModelsReloaded();

            Assert.False(tab.IsModelEditorOpen);
            Assert.Equal("DS1054Z", tab.SeriesOrModelComboBox.SelectedItem);
            Assert.False(tab.EditMyModelButton.IsVisible);
            Assert.Empty(this.ModelsInTheFile());
        });
    }

    // ###########################################################################################
    // *** A CHANGE TYPED BEFORE THE TAB FOLLOWED THE RELOAD IS NOT LOST (code review, 2026-10-29). ***
    // The data check renames the model on the pool, and the tab follows a moment later: a box left in
    // between found no own model of the old name and dropped the change, "Saved" still showing. The
    // editor follows the reload itself then, and saves the change into the model as it is now named.
    // ###########################################################################################
    [Fact]
    public async Task A_change_made_before_the_tab_followed_a_rename_is_saved_into_the_model()
    {
        await UiTest.RunAsync(async () =>
        {
            TabOscilloscope tab = await TabWithOwnModelAsync();
            MainDocument.OscilloscopeRow published = Model("Rigol", "DS1054Z");
            published.DumpImage = ":DISP:DATA? PNG";
            this.ReloadPublishing(published);

            (TextBox stop, _) = tab.FieldForTests("stop");
            stop.Text = ":STOP:NOW";
            Leave(stop);
            await tab.LastModelSaveForTests;

            Assert.Equal(("Rigol", "DS1054Z (my version)"), tab.EditedModelForTests);
            Assert.Equal(":STOP:NOW", Assert.Single(this.ModelsInTheFile()).Stop);
        });
    }

    // ###########################################################################################
    // *** THE MODEL RENAMED IS THE ONE FOLLOWED - BY ITS RENAME, NEVER A GUESS (code review,
    // 2026-10-29). *** Found by its name's start, a "(my version)" the user had already would be
    // taken for the model renamed now, "DS1054Z (my version) 2".
    // ###########################################################################################
    [Fact]
    public async Task The_model_followed_is_the_one_the_reload_renamed()
    {
        await UiTest.RunAsync(async () =>
        {
            TabOscilloscope tab = await TabWithOwnModelAsync();
            Assert.Null(await DataManager.SaveOwnScopeModelsAsync([Model("Rigol", "DS1054Z (my version)"), .. DataManager.OwnScopeModels]));

            MainDocument.OscilloscopeRow published = Model("Rigol", "DS1054Z");
            published.DumpImage = ":DISP:DATA? PNG";
            this.ReloadPublishing(published);

            tab.FollowScopeModelsReloaded();

            Assert.Equal(("Rigol", "DS1054Z (my version) 2"), tab.EditedModelForTests);
            Assert.Equal("DS1054Z (my version) 2", tab.SeriesOrModelComboBox.SelectedItem);
        });
    }

    // A reload that takes away the model shown leaves another model selected - and another model's
    // commands must never go to the scope connected with this one: the session goes (code review,
    // 2026-10-29), as when another model is chosen.
    [Fact]
    public void A_reload_that_takes_away_the_model_shown_ends_the_session()
    {
        UiTest.Run(() =>
        {
            TabOscilloscope tab = TabShowing("Siglent", "SDS1000X-E");
            tab.MarkSessionEstablishedForTests();

            DataManager.LoadFrom(
                this.thisWorkspace.Root,
                MainFiles.Write(this.thisWorkspace.Root, [new MainFiles.Board("Commodore/C64/250407", "C64", "250407")], [Model("Rigol", "DS1000Z")]),
                this.OwnModelsFile);

            Assert.True(tab.IsSessionEstablishedForTests);

            tab.FollowScopeModelsReloaded();

            Assert.Equal("DS1000Z", tab.SeriesOrModelComboBox.SelectedItem);
            Assert.False(tab.IsSessionEstablishedForTests);
        });
    }

    // The main file read again with `published` beside the two models it had - as a data check does.
    private void ReloadPublishing(MainDocument.OscilloscopeRow published) =>
        DataManager.LoadFrom(
            this.thisWorkspace.Root,
            MainFiles.Write(
                this.thisWorkspace.Root,
                [new MainFiles.Board("Commodore/C64/250407", "C64", "250407")],
                [Model("Rigol", "DS1000Z"), Model("Siglent", "SDS1000X-E"), published]),
            this.OwnModelsFile);

    // The editor's "Run full test suite" follows the tab's own - whatever enables it, an
    // auto-connect included (code review, 2026-10-29).
    [Fact]
    public async Task The_editor_s_full_test_suite_button_follows_the_tab_s_own()
    {
        await UiTest.RunAsync(async () =>
        {
            TabOscilloscope tab = await TabWithOwnModelAsync();

            tab.RunFullTestSuiteButton.IsEnabled = true;
            Assert.True(tab.ModelRunFullTestSuiteButton.IsEnabled);

            tab.RunFullTestSuiteButton.IsEnabled = false;
            Assert.False(tab.ModelRunFullTestSuiteButton.IsEnabled);
        });
    }

    // A command fixed while the send panel is open leaves the address typed there (code review,
    // 2026-10-29: it was filled again from the settings - empty for a first-time sender).
    [Fact]
    public async Task A_change_while_sending_leaves_the_address_typed()
    {
        await UiTest.RunAsync(async () =>
        {
            TabOscilloscope tab = await TabWithOwnModelAsync();
            tab.ShowSubmitPanel();
            tab.SubmitModelEmailTextBox.Text = "anna@example.com";

            (TextBox stop, _) = tab.FieldForTests("stop");
            stop.Text = ":STOP:NOW";
            Leave(stop);
            await tab.LastModelSaveForTests;

            Assert.Equal("anna@example.com", tab.SubmitModelEmailTextBox.Text);
        });
    }

    // ###########################################################################################
    // *** A DELETE THAT COULD NOT BE WRITTEN IS SAID (code review, 2026-10-29). *** It closed the
    // editor and dropped the save's failure, and the model came back at the next start unannounced.
    // Now the model stays, and the editor says why.
    // ###########################################################################################
    [Fact]
    public async Task A_delete_that_could_not_be_written_keeps_the_model_and_says_why()
    {
        await UiTest.RunAsync(async () =>
        {
            TabOscilloscope tab = await TabWithOwnModelAsync();

            // The file cannot be replaced: a folder stands where it was.
            File.Delete(this.OwnModelsFile);
            Directory.CreateDirectory(this.OwnModelsFile);

            Click(tab.DeleteModelButton);
            await tab.DeleteEditedModelAsync();

            Assert.True(tab.IsModelEditorOpen);
            Assert.Equal("DS1054Z", Assert.Single(DataManager.OwnScopeModels).SeriesOrModel);
            Assert.StartsWith(TabOscilloscope.NotDeleted, tab.DeleteModelConfirmText.Text);
        });
    }

    // Try it needs the scope connected - and says so beside the command, not only in the output.
    [Fact]
    public async Task Try_it_with_no_scope_connected_says_to_connect_first()
    {
        await UiTest.RunAsync(async () =>
        {
            TabOscilloscope tab = await TabWithOwnModelAsync();

            Assert.Null(tab.ReplyShownForTests(ScopeCommand.Identify));

            await tab.TryModelCommandAsync(ScopeCommand.Identify);

            Assert.Equal((TabOscilloscope.NotConnectedReply, false), tab.ReplyShownForTests(ScopeCommand.Identify));

            // Nothing was tried, so its box is not framed as failed.
            Assert.Null(tab.TrialMarkForTests(ScopeCommand.Identify));
        });
    }

    // ###########################################################################################
    // EACH COMMAND'S BOX IS FRAMED BY HOW ITS TEST WENT (owner request, 2026-11-07: "if a test runs successfully or
    // fails, then it need to be visualized per task ... have the entire input field border in red (and green if OK)"):
    // each step of "Run full test suite" frames its command's box as it goes, a new run starts unframed, and a box
    // changed since its test loses its frame - it said nothing of the command now there.
    // ###########################################################################################
    [Fact]
    public async Task Each_command_s_box_is_framed_green_or_red_by_its_test_and_a_change_takes_the_frame_away()
    {
        await UiTest.RunAsync(async () =>
        {
            TabOscilloscope tab = await TabWithOwnModelAsync();

            Assert.Null(tab.TrialMarkForTests(ScopeCommand.Identify));

            tab.NoteSuiteCommandForTests(ScopeCommand.Identify, worked: true);
            tab.NoteSuiteCommandForTests(ScopeCommand.Stop, worked: false);

            Assert.True(tab.TrialMarkForTests(ScopeCommand.Identify));
            Assert.False(tab.TrialMarkForTests(ScopeCommand.Stop));
            Assert.Contains("TrialPassed", tab.FieldForTests("identify").Box.Classes);
            Assert.Contains("TrialFailed", tab.FieldForTests("stop").Box.Classes);

            (TextBox stop, _) = tab.FieldForTests("stop");
            stop.Text = ":STOP;:STOP";
            Leave(stop);
            await tab.LastModelSaveForTests;

            Assert.Null(tab.TrialMarkForTests(ScopeCommand.Stop));
            Assert.True(tab.TrialMarkForTests(ScopeCommand.Identify));
        });
    }

    // ###########################################################################################
    // The suite frames a command only while it runs - the keyboard's palettes send the same commands and say nothing of
    // the model's test - and, through the real palette, a command the scope's client refused is framed red.
    // ###########################################################################################
    [Fact]
    public async Task A_palette_run_outside_the_suite_frames_nothing()
    {
        await UiTest.RunAsync(async () =>
        {
            TabOscilloscope tab = await TabWithOwnModelAsync();

            // No suite running: a command sent says nothing of the model's test.
            Assert.Null(tab.TrialMarkForTests(ScopeCommand.Run));

            OscilloscopeEntry entry = OscilloscopeEntry.From(DataManager.OwnScopeModels.Single(model => model.SeriesOrModel == "DS1054Z"), isOwn: true);

            // "Stop" is sent, then the palette stops at a command this model has none for: neither is framed.
            await Assert.ThrowsAsync<InvalidOperationException>(() => tab.ExecutePaletteForTestsAsync(new TakingScopeClient(), entry, ScopeCommandPalette.Stop));
            await Assert.ThrowsAsync<InvalidOperationException>(() => tab.ExecutePaletteForTestsAsync(new TakingScopeClient(), entry, ScopeCommandPalette.Run));

            Assert.Null(tab.TrialMarkForTests(ScopeCommand.Stop));
            Assert.Null(tab.TrialMarkForTests(ScopeCommand.Run));
        });
    }

    // A tab with the user's own model and a "session" with `client` - host and port typed first, as their change ends a
    // session (posted, so run before the client goes in).
    private static async Task<TabOscilloscope> TabConnectedToAsync(IScopeClient client, bool autoConnect = false)
    {
        TabOscilloscope tab = await TabWithOwnModelAsync();
        tab.HostTextBox.Text = "192.0.2.10";
        tab.PortTextBox.Text = "5555";
        Dispatcher.UIThread.RunJobs();

        tab.UseConnectedClientForTests(client, autoConnect);
        return tab;
    }

    // "Try it" through the session the tab has: the scope's answer under the box, the box framed green.
    [Fact]
    public async Task Try_it_on_a_connected_scope_shows_its_answer_and_frames_the_box_green()
    {
        await UiTest.RunAsync(async () =>
        {
            TabOscilloscope tab = await TabConnectedToAsync(new TakingScopeClient("RIGOL TECHNOLOGIES,DS1054Z,DS1ZA000000001,00.04.04"));

            await tab.TryModelCommandAsync(ScopeCommand.Identify);

            (string text, bool worked) = Assert.NotNull(tab.ReplyShownForTests(ScopeCommand.Identify));
            Assert.True(worked);
            Assert.StartsWith("RIGOL TECHNOLOGIES,DS1054Z", text, StringComparison.Ordinal);
            Assert.True(tab.TrialMarkForTests(ScopeCommand.Identify));
            Assert.True(tab.IsSessionEstablishedForTests);
        });
    }

    // ###########################################################################################
    // A COMMAND THE SCOPE LEAVES UNANSWERED (owner report, 2026-11-07: "*IaDN?" - "some seconds before it showed anything.
    // If there is a wait time here, then that needs to be visualized. Also, it then seems to repeat forever(?) this
    // command"): while "Try it" waits its line counts down; then the box is framed red, the session is let go of, and -
    // the command being the Identify that connecting sends - automatic connect is held until "Connect to oscilloscope"
    // is pressed, said in the output.
    // ###########################################################################################
    [Fact]
    public async Task An_unanswered_Identify_counts_down_frames_red_and_holds_automatic_connect()
    {
        await UiTest.RunAsync(async () =>
        {
            var silent = new SilentScopeClient();
            TabOscilloscope tab = await TabConnectedToAsync(silent, autoConnect: true);
            tab.TrialAnswerLimitForTests = TimeSpan.FromSeconds(2);
            tab.StartRecordingOutputLinesForTests();

            Task trying = tab.TryModelCommandAsync(ScopeCommand.Identify);

            await silent.Asked.Task;
            Dispatcher.UIThread.RunJobs();

            (string waiting, _) = Assert.NotNull(tab.ReplyShownForTests(ScopeCommand.Identify));
            Assert.Matches(@"^Waiting for the scope to answer - [12] s left$", waiting);
            Assert.Null(tab.TrialMarkForTests(ScopeCommand.Identify));

            await trying;

            (string text, bool worked) = Assert.NotNull(tab.ReplyShownForTests(ScopeCommand.Identify));
            Assert.False(worked);
            Assert.StartsWith("The scope did not answer in time", text, StringComparison.Ordinal);
            Assert.False(tab.TrialMarkForTests(ScopeCommand.Identify));
            Assert.False(tab.IsSessionEstablishedForTests);
            Assert.True(tab.IsAutomaticConnectHeldForTests);
            Assert.Contains(tab.RecordedOutputLinesForTests, line => line.Contains(TabOscilloscope.AutomaticConnectHeldLine, StringComparison.Ordinal));
        });
    }

    // Another command left unanswered ends the session too, but connecting sends Identify, not it: nothing is held.
    [Fact]
    public async Task An_unanswered_command_other_than_Identify_holds_nothing()
    {
        await UiTest.RunAsync(async () =>
        {
            TabOscilloscope tab = await TabConnectedToAsync(new SilentScopeClient(), autoConnect: true);
            tab.TrialAnswerLimitForTests = TimeSpan.FromMilliseconds(300);

            await tab.TryModelCommandAsync(ScopeCommand.QueryTimeDiv);

            Assert.False(tab.TrialMarkForTests(ScopeCommand.QueryTimeDiv));
            Assert.False(tab.IsAutomaticConnectHeldForTests);
        });
    }

    // ###########################################################################################
    // *** "RUN FULL TEST SUITE" FRAMES A COMMAND THAT ANSWERS NOTHING BY THE SCOPE'S "LAST ERROR" (code review,
    // 2026-11-07). *** It was framed green as soon as its bytes went out: a command with a typo the scope silently refused
    // was green in the suite and red under "Try it". Its palette asks "Last error" after it - an error is red, none green.
    // ###########################################################################################
    [Fact]
    public async Task The_suite_frames_a_command_that_answers_nothing_by_the_scopes_last_error()
    {
        await UiTest.RunAsync(async () =>
        {
            TabOscilloscope tab = await TabWithOwnModelAsync();
            var entry = new OscilloscopeEntry
            {
                Brand = "Rigol",
                SeriesOrModel = "DS1054Z",
                Port = "5555",
                Identify = "*IDN?",
                Stop = ":STOP",
                OperationComplete = "*OPC?",
                DrainErrorQueue = ":SYST:ERR?",
            };

            // The scope kept an error: "Stop" was not taken. "Last error" itself answered.
            await tab.ExecutePaletteAsSuiteForTestsAsync(new TakingScopeClient("-113,\"Undefined header\""), entry, ScopeCommandPalette.Stop);
            Dispatcher.UIThread.RunJobs();

            Assert.False(tab.TrialMarkForTests(ScopeCommand.Stop));
            Assert.True(tab.TrialMarkForTests(ScopeCommand.DrainErrorQueue));

            await tab.ExecutePaletteAsSuiteForTestsAsync(new TakingScopeClient("0,\"No error\""), entry, ScopeCommandPalette.Stop);
            Dispatcher.UIThread.RunJobs();

            Assert.True(tab.TrialMarkForTests(ScopeCommand.Stop));
        });
    }

    // A scope that takes a command and never answers it.
    private sealed class SilentScopeClient : IScopeClient
    {
        public TaskCompletionSource Asked { get; } = new(TaskCreationOptions.RunContinuationsAsynchronously);

        public Task SendAsync(string commandText, CancellationToken cancellationToken) => Task.CompletedTask;

        public async Task<string> QueryLineAsync(string commandText, CancellationToken cancellationToken)
        {
            this.Asked.TrySetResult();
            await Task.Delay(Timeout.Infinite, cancellationToken);
            return string.Empty;
        }

        public Task<byte[]> QueryBinaryBlockAsync(string commandText, CancellationToken cancellationToken) => Task.FromResult(Array.Empty<byte>());
    }

    // A scope client that takes every command - what a palette sends goes nowhere.
    private sealed class TakingScopeClient : IScopeClient
    {
        private readonly string thisAnswer;

        public TakingScopeClient(string answer = "0,\"No error\"") => this.thisAnswer = answer;

        public Task SendAsync(string commandText, CancellationToken cancellationToken) => Task.CompletedTask;

        public Task<string> QueryLineAsync(string commandText, CancellationToken cancellationToken) => Task.FromResult(this.thisAnswer);

        public Task<byte[]> QueryBinaryBlockAsync(string commandText, CancellationToken cancellationToken) => Task.FromResult(Array.Empty<byte>());
    }
}
