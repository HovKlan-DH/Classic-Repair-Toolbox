using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Layout;
using Handlers.DataHandling;
using Handlers.Oscilloscope;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace CRT
{
    // ###########################################################################################
    // "MY SCOPE IS NOT LISTED" - A MODEL OF THE USER'S OWN, made and edited on this tab (CRT 4.0
    // step 5 - the plan: "On the Oscilloscope tab, start from a similar model and edit each command
    // with Try it, then run the full test suite. Your own models are kept outside the synced data").
    // Before 4.0 the way was a row added to the main Excel file, which the launch's data check put
    // back unless it was switched off.
    //
    // The editor sits above the output, which keeps showing every command sent:
    //
    //   - "My scope is not listed" opens a form: the model to START FROM (the one shown, preselected
    //     - often one of the same vendor) and the new model's vendor and name. "Create model" copies
    //     every command of it under that name into the user's own models (OscilloscopeModelRules)
    //     and shows it in the drop-downs, ready to edit.
    //   - A model of the user's own shows "Edit my model": its name, its two settings, its sixteen
    //     commands with "Try it" each, and its two lists of steps (ScopeModelFields). A box is SAVED
    //     AS IT IS LEFT, or Enter is pressed (Escape puts it back) - there is no Save button, as in
    //     the rest of 4.0. A setting CRT could not use is held back with the reason under it; a list
    //     of steps is saved anyway, its unreadable steps named.
    //   - "Try it" sends that one command, as the box holds it, over the session the tab has
    //     (ScopeCommandTrial), and its answer shows under the box. "Run full test suite" is the
    //     tab's own button: the model edited is the model selected.
    //   - "Delete model" asks once, inline, and takes it out of the list.
    //
    // *** A REFRESH OF THE DROP-DOWNS FOR A NEW OR RENAMED MODEL KEEPS THE CONNECTION. *** Choosing
    // another model drops the session (a different scope's commands); the editor's own refresh
    // selects the model the user is working on, on the scope already connected - so they can
    // connect with the similar model and try their own at once (thisKeepsSessionThroughListRefresh).
    // ###########################################################################################
    public partial class TabOscilloscope
    {
        internal const string NotConnectedReply = "Connect to your scope first - \"Connect to oscilloscope\", on the left.";

        internal const string TryingReply = "Trying...";

        internal const string NewModelIntro =
            "CRT talks to your scope over its network port, with SCPI commands - so the scope needs a network (LAN) port " +
            "that takes them, which not every scope has. Start from the model most like yours, often one of the same " +
            "vendor, and give yours a name of its own. Your scope's programming manual lists its commands.";

        internal const string EditModelIntro =
            "Your own model. Each change is saved as you leave its box, into a file of your own the data check never " +
            "puts back. Connect to your scope, then press Try it beside a command to see its answer; Run full test " +
            "suite tries them all in turn and changes the scope's settings while it runs.";

        internal const string CommandsIntro =
            "Try it sends the command as you typed it and shows the scope's answer under it - and every command sent " +
            "in the output below.";

        // While the editor's own refresh of the drop-downs runs: a selection change then is the model
        // being worked on, on the same scope, and keeps the session (see the header).
        private bool thisKeepsSessionThroughListRefresh;

        // The model of the user's own the editor shows, by its name - null while the editor is
        // closed or shows the new-model form.
        private (string Vendor, string Model)? thisEditedModel;

        // What the start-from list holds, in its order.
        private List<MainDocument.OscilloscopeRow> thisStartFromModels = new();

        private sealed class ModelFieldRow
        {
            public required ScopeModelField Field { get; init; }

            public required TextBox Box { get; init; }

            public required TextBlock Problem { get; init; }

            public required TextBlock Reply { get; init; }

            public Button? TryIt { get; init; }
        }

        private readonly List<ModelFieldRow> thisModelFieldRows = new();

        // The last save the editor started - a test waits for it.
        private Task thisLastModelSave = Task.CompletedTask;

        internal Task LastModelSaveForTests => this.thisLastModelSave;

        internal bool IsModelEditorOpen => this.ModelEditorBorder.IsVisible;

        internal (string Vendor, string Model)? EditedModelForTests => this.thisEditedModel;

        private void InitializeModelEditor()
        {
            this.SeriesOrModelComboBox.SelectionChanged += (_, _) => this.FollowSelectionInModelEditor();

            // The editor's "Run full test suite" is the tab's own, so it follows it everywhere - an
            // auto-connect and a session lost in the background included (code review, 2026-10-29).
            this.ModelRunFullTestSuiteButton.Bind(IsEnabledProperty, this.RunFullTestSuiteButton.GetObservable(IsEnabledProperty));

            foreach (TextBox nameBox in new[] { this.ModelVendorTextBox, this.ModelNameTextBox })
            {
                nameBox.LostFocus += (_, _) => this.CommitModelName();
                nameBox.KeyDown += this.OnModelNameKeyDown;
            }

            this.BuildModelFieldRows();
            this.FollowSelectionInModelEditor();
        }

        // ###########################################################################################
        // One row per value (ScopeModelFields), made once: label, box, Try it for a command, its
        // help, its problem and its answer - a heading where the kind of value changes.
        // ###########################################################################################
        private void BuildModelFieldRows()
        {
            ScopeModelFieldKind? kind = null;

            foreach (ScopeModelField field in ScopeModelFields.All)
            {
                if (field.Kind != kind)
                {
                    kind = field.Kind;
                    this.ModelFieldsPanel.Children.Add(new TextBlock
                    {
                        Classes = { "ModelHeading" },
                        Margin = new Thickness(0, 14, 0, 0),
                        Text = field.Kind switch
                        {
                            ScopeModelFieldKind.Setting => "Settings",
                            ScopeModelFieldKind.Command => "Commands",
                            _ => "Steps",
                        },
                    });

                    if (field.Kind == ScopeModelFieldKind.Command)
                    {
                        this.ModelFieldsPanel.Children.Add(new TextBlock { Classes = { "ModelIntro" }, Text = TabOscilloscope.CommandsIntro });
                    }
                }

                var box = new TextBox
                {
                    FontSize = 12,
                    FontFamily = field.Kind == ScopeModelFieldKind.Command ? new Avalonia.Media.FontFamily("Consolas") : Avalonia.Media.FontFamily.Default,
                    TextWrapping = field.Kind == ScopeModelFieldKind.StepList ? Avalonia.Media.TextWrapping.Wrap : Avalonia.Media.TextWrapping.NoWrap,
                };

                Button? tryIt = field.Command is null
                    ? null
                    : new Button { Content = "Try it", FontSize = 12, Padding = new Thickness(8, 4), VerticalAlignment = VerticalAlignment.Top };

                var line = new Grid { ColumnDefinitions = new ColumnDefinitions("*,6,Auto") };
                line.Children.Add(box);

                if (tryIt is not null)
                {
                    Grid.SetColumn(tryIt, 2);
                    line.Children.Add(tryIt);
                }

                var row = new ModelFieldRow
                {
                    Field = field,
                    Box = box,
                    Problem = new TextBlock { Classes = { "FieldProblem" }, IsVisible = false },
                    Reply = new TextBlock { Classes = { "TrialReply" }, IsVisible = false },
                    TryIt = tryIt,
                };

                box.LostFocus += (_, _) => this.CommitModelField(row);
                box.KeyDown += (_, e) => this.OnModelFieldKeyDown(row, e);

                if (tryIt is not null)
                {
                    tryIt.Click += (_, _) => this.OnTryItClick(row);
                }

                var panel = new StackPanel { Spacing = 2 };
                panel.Children.Add(new TextBlock { Classes = { "FieldLabel" }, Text = field.Label });
                panel.Children.Add(line);
                panel.Children.Add(new TextBlock { Classes = { "FieldHelp" }, Text = field.Help });
                panel.Children.Add(row.Problem);
                panel.Children.Add(row.Reply);

                this.ModelFieldsPanel.Children.Add(panel);
                this.thisModelFieldRows.Add(row);
            }
        }

        // ###########################################################################################
        // The editor follows the model selected: "Edit my model" only for one of the user's own, and
        // an open editor on another model closes - it edits the model the drop-downs show.
        // ###########################################################################################
        private void FollowSelectionInModelEditor()
        {
            OscilloscopeEntry? selected = this.GetSelectedOscilloscope();

            this.EditMyModelButton.IsVisible = selected?.IsOwn == true;

            if (this.thisKeepsSessionThroughListRefresh || this.thisEditedModel is not { } edited)
            {
                return;
            }

            if (selected is null || !string.Equals(selected.Brand.Trim(), edited.Vendor, StringComparison.OrdinalIgnoreCase) ||
                !string.Equals(selected.SeriesOrModel.Trim(), edited.Model, StringComparison.OrdinalIgnoreCase))
            {
                this.CloseModelEditor();
            }
        }

        // ###########################################################################################
        // THE MODEL LISTS READ AGAIN (Main, as soon as a load lists them - DataManager.ScopeModelsListed;
        // code reviews, 2026-10-29). A model published since under the name of a model of the user's
        // own takes it over (the same commands: the own copy goes) or names it apart ("<model> (my
        // version)"). The drop-downs were never made again, so the editor stayed on a name no own model
        // had, and every change was dropped without a word. Now the drop-downs are made again keeping
        // the session, an own model shown follows ITS rename (DataManager.OwnScopeModelRenamedTo - never
        // a guess from a name's start), and an editor whose model went closes and says why. A selection
        // that lands on another model than the one shown - the main file no longer lists it - ends the
        // session, as choosing another model does: its commands are not that scope's.
        // ###########################################################################################
        internal void FollowScopeModelsReloaded()
        {
            this.thisIsFollowingReload = true;

            try
            {
                this.FollowScopeModelsReloadedCore();
            }
            finally
            {
                this.thisIsFollowingReload = false;
            }
        }

        private void FollowScopeModelsReloadedCore()
        {
            string vendor = this.VendorComboBox.SelectedItem as string ?? string.Empty;
            string model = this.SeriesOrModelComboBox.SelectedItem as string ?? string.Empty;

            if (vendor.Length == 0 || model.Length == 0)
            {
                this.PopulateVendorDropDown();
                this.ShowModelsOfSelectedVendor();
                return;
            }

            bool wasOwn = this.thisEditedModel is not null || this.thisPublishedSelectionInputs?.SelectedOscilloscope?.IsOwn == true;
            bool stillOwn = TabOscilloscope.OwnModelNamed(vendor, model) is not null;

            string? followed = wasOwn && !stillOwn && DataManager.OwnScopeModelRenamedTo(vendor, model) is string renamed &&
                               TabOscilloscope.OwnModelNamed(vendor, renamed) is not null
                ? renamed
                : null;

            bool editing = this.thisEditedModel is not null;

            if (followed is not null && editing)
            {
                this.thisEditedModel = (vendor, followed);
                this.ModelEditorHeading.Text = $"{vendor} {followed}";

                // The name boxes follow - unless a new name is being typed into them.
                if (string.Equals(this.ModelNameTextBox.Text?.Trim(), model, StringComparison.OrdinalIgnoreCase))
                {
                    this.ModelVendorTextBox.Text = vendor;
                    this.ModelNameTextBox.Text = followed;
                }
            }

            string shown = followed ?? model;

            this.RefreshModelLists(vendor, shown);

            if (!string.Equals(this.VendorComboBox.SelectedItem as string, vendor, StringComparison.OrdinalIgnoreCase) ||
                !string.Equals(this.SeriesOrModelComboBox.SelectedItem as string, shown, StringComparison.OrdinalIgnoreCase))
            {
                this.InvalidateEstablishedOscilloscopeSession();
            }

            if (followed is not null)
            {
                string said = string.Format(CultureInfo.InvariantCulture, TabOscilloscope.KeptAsMyVersion, vendor, model, followed);

                if (editing)
                {
                    this.ModelEditorSaveStateText.Text = said;
                    this.ModelEditorSaveStateText.IsVisible = true;
                }

                this.AppendOutputLine("Info", said);
            }
            else if (wasOwn && !stillOwn)
            {
                this.CloseModelEditor();
                this.AppendOutputLine("Info", string.Format(CultureInfo.InvariantCulture, TabOscilloscope.TakenOverByPublished, vendor, model));
            }
        }

        internal const string KeptAsMyVersion =
            "{0} {1} was published with other commands than yours, so your own model is kept beside it as \"{2}\".";

        internal const string TakenOverByPublished =
            "{0} {1} was published just as your own model is, so CRT now uses the published one and your copy went.";

        private void OnScopeNotListedClick(object? sender, RoutedEventArgs e) => this.ShowNewModelForm();

        private void OnEditMyModelClick(object? sender, RoutedEventArgs e)
        {
            if (this.GetSelectedOscilloscope() is { IsOwn: true } selected)
            {
                this.ShowModelToEdit(selected.Brand, selected.SeriesOrModel);
            }
        }

        private void OnModelEditorCloseClick(object? sender, RoutedEventArgs e) => this.CloseModelEditor();

        // ###########################################################################################
        // The new-model form: every model listed to start from, the one shown preselected, its vendor
        // in the vendor box and the name left for the user.
        // ###########################################################################################
        internal void ShowNewModelForm()
        {
            this.thisEditedModel = null;
            this.thisStartFromModels = DataManager.PublishedScopeModels
                .Concat(DataManager.OwnScopeModels)
                .Where(model => !string.IsNullOrWhiteSpace(model.SeriesOrModel))
                .ToList();

            this.StartFromComboBox.ItemsSource = this.thisStartFromModels
                .Select(model => $"{model.Brand.Trim()} {model.SeriesOrModel.Trim()}")
                .ToList();

            OscilloscopeEntry? selected = this.GetSelectedOscilloscope();
            int shown = selected is null
                ? -1
                : this.thisStartFromModels.FindIndex(model => OscilloscopeModelRules.IsNamed(model, selected.Brand, selected.SeriesOrModel));

            this.StartFromComboBox.SelectedIndex = shown >= 0 ? shown : (this.thisStartFromModels.Count > 0 ? 0 : -1);
            this.NewModelVendorTextBox.Text = selected?.Brand.Trim() ?? string.Empty;
            this.NewModelNameTextBox.Text = string.Empty;
            this.NewModelProblemText.IsVisible = false;

            this.ModelEditorHeading.Text = "My scope is not listed";
            this.ModelEditorIntro.Text = TabOscilloscope.NewModelIntro;
            this.ModelEditorSaveStateText.IsVisible = false;
            this.NewModelPanel.IsVisible = true;
            this.EditModelPanel.IsVisible = false;
            this.OpenModelEditor();
        }

        private async void OnCreateModelClick(object? sender, RoutedEventArgs e)
        {
            // Guarded as a whole: an "async void" handler has no caller to take an exception.
            try
            {
                await this.CreateModelAsync();
            }
            catch (Exception ex)
            {
                Logger.Critical($"OnCreateModelClick failed - [{ex}]");
                this.ShowNewModelProblem($"The model could not be made: {ex.Message}");
            }
        }

        internal async Task CreateModelAsync()
        {
            // The own models' file could not be read at launch: written now, it would lose them.
            if (DataManager.OwnScopeModelsProblem is string notReadable)
            {
                this.ShowNewModelProblem(notReadable);
                return;
            }

            int startFrom = this.StartFromComboBox.SelectedIndex;

            if (startFrom < 0 || startFrom >= this.thisStartFromModels.Count)
            {
                this.ShowNewModelProblem("Pick the model to start from.");
                return;
            }

            string vendor = this.NewModelVendorTextBox.Text?.Trim() ?? string.Empty;
            string model = this.NewModelNameTextBox.Text?.Trim() ?? string.Empty;

            if (OscilloscopeModelRules.NameProblem(vendor, model, TabOscilloscope.EveryModel()) is string problem)
            {
                this.ShowNewModelProblem(problem);
                return;
            }

            MainDocument.OscilloscopeRow created = OscilloscopeModelRules.CopyOf(this.thisStartFromModels[startFrom], vendor, model);

            Task<string?> save = DataManager.SaveOwnScopeModelsAsync(DataManager.OwnScopeModels.Append(created).ToList());
            this.thisLastModelSave = save;

            this.RefreshModelLists(vendor, model);
            this.ShowModelToEdit(vendor, model);
            this.ShowSaveState(await save);
        }

        private void ShowNewModelProblem(string problem)
        {
            this.NewModelProblemText.Text = problem;
            this.NewModelProblemText.IsVisible = true;
        }

        // ###########################################################################################
        // The edit panel on the user's own model named `vendor` + `model`: every box filled from it,
        // the answers of an earlier model's tries gone.
        // ###########################################################################################
        internal void ShowModelToEdit(string vendor, string model)
        {
            if (TabOscilloscope.OwnModelNamed(vendor, model) is not MainDocument.OscilloscopeRow row)
            {
                this.CloseModelEditor();
                return;
            }

            this.thisEditedModel = (row.Brand.Trim(), row.SeriesOrModel.Trim());

            this.ModelVendorTextBox.Text = row.Brand.Trim();
            this.ModelNameTextBox.Text = row.SeriesOrModel.Trim();
            this.ModelNameProblemText.IsVisible = false;

            foreach (ModelFieldRow fieldRow in this.thisModelFieldRows)
            {
                fieldRow.Box.Text = fieldRow.Field.Get(row);
                fieldRow.Reply.IsVisible = false;
                this.ShowFieldProblem(fieldRow, fieldRow.Field.Kind == ScopeModelFieldKind.StepList ? ScopeModelFields.ProblemOf(fieldRow.Field, fieldRow.Box.Text) : null);
            }

            this.ModelEditorHeading.Text = $"{row.Brand.Trim()} {row.SeriesOrModel.Trim()}";
            this.ModelEditorIntro.Text = TabOscilloscope.EditModelIntro;
            this.SubmitModelPanel.IsVisible = false;
            this.ShowModelSent(row);
            this.DeleteModelConfirmPanel.IsVisible = false;
            this.DeleteModelButton.IsVisible = true;
            this.NewModelPanel.IsVisible = false;
            this.EditModelPanel.IsVisible = true;
            this.OpenModelEditor();
        }

        private void OpenModelEditor()
        {
            if (!this.ModelEditorBorder.IsVisible)
            {
                this.OutputColumnGrid.RowDefinitions[0].Height = new GridLength(3, GridUnitType.Star);
                this.OutputColumnGrid.RowDefinitions[1].Height = new GridLength(4);
                this.OutputColumnGrid.RowDefinitions[2].Height = new GridLength(1, GridUnitType.Star);
            }

            this.ModelEditorBorder.IsVisible = true;
            this.ModelEditorSplitter.IsVisible = true;
        }

        internal void CloseModelEditor()
        {
            this.thisEditedModel = null;
            this.ModelEditorBorder.IsVisible = false;
            this.ModelEditorSplitter.IsVisible = false;
            this.OutputColumnGrid.RowDefinitions[0].Height = new GridLength(0);
            this.OutputColumnGrid.RowDefinitions[1].Height = new GridLength(0);
            this.OutputColumnGrid.RowDefinitions[2].Height = new GridLength(1, GridUnitType.Star);
        }

        // ###########################################################################################
        // A box left, or Enter: the value saved into the model - unless it is the value it has, or a
        // setting CRT could not use (held back, the reason under it). Escape puts the box back.
        // ###########################################################################################
        private void OnModelFieldKeyDown(ModelFieldRow row, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                this.CommitModelField(row);
                e.Handled = true;
            }
            else if (e.Key == Key.Escape && this.EditedModelRow() is MainDocument.OscilloscopeRow model)
            {
                row.Box.Text = row.Field.Get(model);
                this.ShowFieldProblem(row, row.Field.Kind == ScopeModelFieldKind.StepList ? ScopeModelFields.ProblemOf(row.Field, row.Box.Text) : null);
                e.Handled = true;
            }
        }

        private void CommitModelField(ModelFieldRow row)
        {
            if (this.EditedModelRowFollowingReload() is not MainDocument.OscilloscopeRow model)
            {
                return;
            }

            string value = row.Box.Text?.Trim() ?? string.Empty;
            string? problem = ScopeModelFields.ProblemOf(row.Field, value);

            this.ShowFieldProblem(row, problem);

            if (problem is not null && ScopeModelFields.IsHeldBackWhenWrong(row.Field))
            {
                return;
            }

            if (string.Equals(row.Field.Get(model), value, StringComparison.Ordinal))
            {
                return;
            }

            this.SaveEditedModel(OscilloscopeModelRules.With(model, copy => row.Field.Set(copy, value)));

            // The model's port is the one "Connect" uses, as when a model is picked - but never under
            // a session, which the box's change would end (OnSeriesOrModelSelectionChanged's rule).
            if (row.Field.Key == ScopeModelFields.PortKey && !this.thisHasEstablishedOscilloscopeSession)
            {
                this.UseModelPort(value);
            }
        }

        private void ShowFieldProblem(ModelFieldRow row, string? problem)
        {
            row.Problem.Text = problem ?? string.Empty;
            row.Problem.IsVisible = problem is not null;
        }

        private void OnModelNameKeyDown(object? sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                this.CommitModelName();
                e.Handled = true;
            }
            else if (e.Key == Key.Escape && this.thisEditedModel is { } edited)
            {
                this.ModelVendorTextBox.Text = edited.Vendor;
                this.ModelNameTextBox.Text = edited.Model;
                this.ModelNameProblemText.IsVisible = false;
                e.Handled = true;
            }
        }

        // ###########################################################################################
        // A new name for the model - refused while another model has it (the reason stays under the
        // boxes, the boxes as typed); taken, the drop-downs show the model under it.
        // ###########################################################################################
        private void CommitModelName()
        {
            if (this.EditedModelRowFollowingReload() is not MainDocument.OscilloscopeRow model || this.thisEditedModel is not { } edited)
            {
                return;
            }

            string vendor = this.ModelVendorTextBox.Text?.Trim() ?? string.Empty;
            string name = this.ModelNameTextBox.Text?.Trim() ?? string.Empty;

            if (string.Equals(vendor, edited.Vendor, StringComparison.Ordinal) && string.Equals(name, edited.Model, StringComparison.Ordinal))
            {
                this.ModelNameProblemText.IsVisible = false;
                return;
            }

            IEnumerable<MainDocument.OscilloscopeRow> others = TabOscilloscope.EveryModel()
                .Where(other => !OscilloscopeModelRules.IsNamed(other, edited.Vendor, edited.Model));

            if (OscilloscopeModelRules.NameProblem(vendor, name, others) is string problem)
            {
                this.ModelNameProblemText.Text = problem;
                this.ModelNameProblemText.IsVisible = true;
                return;
            }

            this.ModelNameProblemText.IsVisible = false;
            this.thisEditedModel = (vendor, name);
            this.SaveEditedModel(OscilloscopeModelRules.With(model, copy =>
            {
                copy.Brand = vendor;
                copy.SeriesOrModel = name;
            }), formerName: edited);

            this.RefreshModelLists(vendor, name);
            this.ModelEditorHeading.Text = $"{vendor} {name}";
        }

        // The model edited replaced by `changed` in the user's own models, and saved.
        private void SaveEditedModel(MainDocument.OscilloscopeRow changed, (string Vendor, string Model)? formerName = null)
        {
            (string Vendor, string Model) named = formerName ?? this.thisEditedModel!.Value;

            List<MainDocument.OscilloscopeRow> own = DataManager.OwnScopeModels
                .Select(model => OscilloscopeModelRules.IsNamed(model, named.Vendor, named.Model) ? changed : model)
                .ToList();

            Task<string?> save = DataManager.SaveOwnScopeModelsAsync(own);
            this.thisLastModelSave = save;

            // The workers read the selection's model from what was last published.
            this.PublishOscilloscopeSelectionInputs();
            _ = this.ShowSaveStateWhenDoneAsync(save);

            // A change since the suite went through makes that no longer so - the line follows, and
            // nothing else in the panel: the address typed there stays (code review, 2026-10-29).
            if (this.SubmitModelPanel.IsVisible)
            {
                this.ShowSubmitSuiteLine();
            }
        }

        private async Task ShowSaveStateWhenDoneAsync(Task<string?> save)
        {
            try
            {
                string? failure = await save;

                if (ReferenceEquals(save, this.thisLastModelSave))
                {
                    this.ShowSaveState(failure);
                }
            }
            catch (Exception ex)
            {
                Logger.Critical($"Saving your own oscilloscope models failed - [{ex}]");
                this.ShowSaveState(ex.Message);
            }
        }

        private void ShowSaveState(string? failure)
        {
            this.ModelEditorSaveStateText.Text = failure is null
                ? "Saved"
                : $"Could not save - {failure}. CRT keeps your change and tries again with the next one.";
            this.ModelEditorSaveStateText.IsVisible = true;
        }

        private void OnDeleteModelClick(object? sender, RoutedEventArgs e)
        {
            if (this.thisEditedModel is not { } edited)
            {
                return;
            }

            this.DeleteModelConfirmText.Text = $"Delete {edited.Vendor} {edited.Model}? It cannot be brought back.";
            this.DeleteModelConfirmPanel.IsVisible = true;
            this.DeleteModelButton.IsVisible = false;
        }

        private void OnKeepModelClick(object? sender, RoutedEventArgs e)
        {
            this.DeleteModelConfirmPanel.IsVisible = false;
            this.DeleteModelButton.IsVisible = true;
        }

        internal const string NotDeleted = "The model could not be deleted:";

        private async void OnConfirmDeleteModelClick(object? sender, RoutedEventArgs e)
        {
            // Guarded as a whole: an "async void" handler has no caller to take an exception.
            try
            {
                await this.DeleteEditedModelAsync();
            }
            catch (Exception ex)
            {
                Logger.Critical($"OnConfirmDeleteModelClick failed - [{ex}]");
                this.DeleteModelConfirmText.Text = $"{TabOscilloscope.NotDeleted} {ex.Message}";
            }
        }

        // ###########################################################################################
        // The model taken out of the user's own - and only once the file says so (code review,
        // 2026-10-29: the save's failure was dropped, the editor closed, and the model came back at
        // the next start unannounced). A delete that could not be written puts the model back in the
        // lists - the file still holds it - and says why where Delete was asked.
        // ###########################################################################################
        internal async Task DeleteEditedModelAsync()
        {
            if (this.thisEditedModel is not { } edited)
            {
                return;
            }

            IReadOnlyList<MainDocument.OscilloscopeRow> before = DataManager.OwnScopeModels;
            List<MainDocument.OscilloscopeRow> own = before
                .Where(model => !OscilloscopeModelRules.IsNamed(model, edited.Vendor, edited.Model))
                .ToList();

            Task<string?> save = DataManager.SaveOwnScopeModelsAsync(own);
            this.thisLastModelSave = save;

            if (await save is string failure)
            {
                DataManager.PutBackOwnScopeModels(before);
                this.DeleteModelConfirmText.Text = $"{TabOscilloscope.NotDeleted} {failure}";
                return;
            }

            this.CloseModelEditor();

            // Another model is shown now - picked by the drop-downs' own rule, so the session goes.
            this.PopulateVendorDropDown();
            this.ShowModelsOfSelectedVendor();
            this.FollowSelectionInModelEditor();
        }

        // ###########################################################################################
        // The drop-downs made again from the lists, `vendor` + `model` selected - keeping the session
        // (see the header): the selection handlers pick what UserSettings remembers, so that is set
        // first, and the models are shown even when the vendor's selection did not change.
        // ###########################################################################################
        private void RefreshModelLists(string vendor, string model)
        {
            UserSettings.SetLastOscilloscopeVendor(vendor);
            UserSettings.SetLastOscilloscopeSeriesForVendor(vendor, model);

            this.thisKeepsSessionThroughListRefresh = true;

            try
            {
                this.PopulateVendorDropDown();
                this.ShowModelsOfSelectedVendor();
            }
            finally
            {
                this.thisKeepsSessionThroughListRefresh = false;
            }

            this.PublishOscilloscopeSelectionInputs();
            this.FollowSelectionInModelEditor();
        }

        // ###########################################################################################
        // Try it: the command as its box holds it (committed first), on the session the tab has.
        // ###########################################################################################
        private async void OnTryItClick(ModelFieldRow row)
        {
            // Guarded as a whole: an "async void" handler has no caller to take an exception.
            try
            {
                await this.TryModelCommandAsync(row.Field.Command!.Value);
            }
            catch (Exception ex)
            {
                Logger.Critical($"OnTryItClick failed - [{ex}]");
                this.ShowReply(row, new ScopeTrialOutcome(false, ex.Message));
            }
        }

        internal async Task TryModelCommandAsync(ScopeCommand command)
        {
            ModelFieldRow row = this.thisModelFieldRows.Single(candidate => candidate.Field.Command == command);

            this.CommitModelField(row);

            if (this.EditedModelRowFollowingReload() is not MainDocument.OscilloscopeRow model)
            {
                return;
            }

            OscilloscopeSelectionSnapshot snapshot = this.CreateOscilloscopeSelectionSnapshot();

            if (!snapshot.HasActiveEstablishedSession)
            {
                this.ShowReply(row, new ScopeTrialOutcome(false, TabOscilloscope.NotConnectedReply));
                return;
            }

            this.ShowReply(row, new ScopeTrialOutcome(true, TabOscilloscope.TryingReply));
            this.AppendOutputLine("Debug", "---");
            this.AppendOutputLine("Debug", $"Try it - {row.Field.Label}:");

            ScopeTrialOutcome? outcome = null;
            string? failure = null;

            await this.RunWithEstablishedOscilloscopeSessionAsync(
                snapshot,
                async (scopeClient, _, cancellationToken) =>
                {
                    // A command the scope does not know is often met with silence: the answer is
                    // waited for a shorter time than the session's, but a picture can take a while.
                    using var limit = new CancellationTokenSource(command == ScopeCommand.DumpImage ? TimeSpan.FromSeconds(30) : TimeSpan.FromSeconds(10));
                    using var linked = CancellationTokenSource.CreateLinkedTokenSource(limit.Token, cancellationToken);

                    try
                    {
                        outcome = await ScopeCommandTrial.TryAsync(
                            scopeClient,
                            model,
                            command,
                            line => this.AppendOutputLine("Debug", line),
                            linked.Token).ConfigureAwait(false);
                    }
                    catch (OperationCanceledException)
                    {
                        failure = "The scope did not answer in time. CRT let go of the connection, since the scope may still " +
                                  "answer later - check the command, and connect again.";
                        throw;
                    }
                    catch (Exception ex)
                    {
                        failure = $"{ex.Message} - CRT let go of the connection; connect again.";
                        throw;
                    }
                },
                CancellationToken.None,
                writeWarnings: true);

            this.ShowReply(row, outcome ?? new ScopeTrialOutcome(false, failure ?? TabOscilloscope.NotConnectedReply));
        }

        private void ShowReply(ModelFieldRow row, ScopeTrialOutcome outcome)
        {
            row.Reply.Text = outcome.Reply;
            row.Reply.Classes.Set("failed", !outcome.Worked);
            row.Reply.IsVisible = true;
        }

        // The Try it buttons follow the tab's own: off while a command runs. With nothing connected a
        // press says so under the command.
        private void SetTryItButtonsEnabled(bool isEnabled)
        {
            foreach (ModelFieldRow row in this.thisModelFieldRows)
            {
                if (row.TryIt is not null)
                {
                    row.TryIt.IsEnabled = isEnabled;
                }
            }
        }

        // The answer shown under a command's box - for a test.
        internal (string Text, bool Worked)? ReplyShownForTests(ScopeCommand command)
        {
            ModelFieldRow row = this.thisModelFieldRows.Single(candidate => candidate.Field.Command == command);

            return row.Reply.IsVisible ? (row.Reply.Text ?? string.Empty, !row.Reply.Classes.Contains("failed")) : null;
        }

        // The box of a value, its problem shown - for a test.
        internal (TextBox Box, string? Problem) FieldForTests(string key)
        {
            ModelFieldRow row = this.thisModelFieldRows.Single(candidate => candidate.Field.Key == key);

            return (row.Box, row.Problem.IsVisible ? row.Problem.Text : null);
        }

        private MainDocument.OscilloscopeRow? EditedModelRow() =>
            this.thisEditedModel is { } edited ? TabOscilloscope.OwnModelNamed(edited.Vendor, edited.Model) : null;

        // ###########################################################################################
        // The model edited, for something done to it - a box left, a name, Try it, a send. When no own
        // model has its name any more, a data check renamed or took it over and the tab has not
        // followed yet (it is told a moment after the lists change): it follows FIRST, so a change
        // typed meanwhile goes into the model as it is now named (code review, 2026-10-29: it was
        // dropped without a word, "Saved" still showing).
        // ###########################################################################################
        private MainDocument.OscilloscopeRow? EditedModelRowFollowingReload()
        {
            MainDocument.OscilloscopeRow? model = this.EditedModelRow();

            if (model is not null || this.thisEditedModel is null || this.thisIsFollowingReload)
            {
                return model;
            }

            this.FollowScopeModelsReloaded();
            return this.EditedModelRow();
        }

        // FollowScopeModelsReloaded is under way - a box left while it makes the drop-downs again does
        // not start it a second time.
        private bool thisIsFollowingReload;

        private static MainDocument.OscilloscopeRow? OwnModelNamed(string vendor, string model) =>
            DataManager.OwnScopeModels.FirstOrDefault(own => OscilloscopeModelRules.IsNamed(own, vendor, model));

        private static IEnumerable<MainDocument.OscilloscopeRow> EveryModel() =>
            DataManager.PublishedScopeModels.Concat(DataManager.OwnScopeModels);
    }
}
