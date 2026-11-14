using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using Avalonia;
using Avalonia.Media;
using Avalonia.Styling;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// EVERY COLOR CRT WRITES WITH READS ON THE GROUND IT IS WRITTEN ON, IN ALL THREE THEMES (owner request, 2026-11-12: "check
// all colors and make sure they have a nice discrete complementary color in a nice visible contrast ratio"). Text at WCAG's
// 4.5:1, an edge or mark that must be seen at its 3:1, each pair below as CRT draws it - in the light theme, the dark one,
// and "User preference": the light theme with the colors UserSettings seeds into the settings file over it, the ground
// a tinted #E0E4D9 that most of the light theme's text colors failed on until that day.
//
// The ratio is CRT.Data's RegionColors.ContrastRatio, the one rule. A new text color, or one drawn on a new ground, gets
// its pair here. Left out on purpose: Workbooks_ZeroCount_Fg (a zero meant to look switched off), and what is drawn over
// a picture (the popup's image counter, the label editor's labels) rather than a theme's ground.
// ###########################################################################################
[Collection("HeadlessUi")]
public class ThemeContrastTests
{
    private const double Text = 4.5;
    private const double Edge = 3.0;

    private static readonly string[] Themes = ["Light", "Dark", "User preference"];

    // (foreground, background, the ratio it must reach)
    private static IEnumerable<(string Fore, string Back, double Least)> Pairs()
    {
        IEnumerable<(string, string, double)> On(string fore, double least, params string[] grounds) =>
            grounds.Select(ground => (fore, ground, least));

        string[] page = ["Bg", "Card_Bg", "Form_Bg"];

        foreach (var pair in On("Fg", Text, "Bg", "Card_Bg", "Form_Bg", "Workbooks_Panel_Bg", "Schematics_Panels_Bg", "Table_Bg",
                     "Table_RowHover_Bg", "Workbooks_SelectedCard_Bg", "Draft_ReferenceBadge_Bg", "Changelog_Code_Bg", "Mode_Selected_Bg",
                     "Notice_Bg", "BoardTable_Error_Bg", "BoardTable_Warning_Bg", "BoardTable_Notice_Bg"))
            yield return pair;

        foreach (var pair in On("Form_Fg", Text, "Form_Bg", "Form_Selected_Hover_Bg")) yield return pair;
        foreach (var pair in On("Form_Selected_Fg", Text, "Form_Selected_Bg", "Form_Selected_Hover_Bg")) yield return pair;
        foreach (var pair in On("Button_Fg", Text, "Button_Bg", "Button_Hover_Bg", "Button_Pressed_Bg")) yield return pair;
        foreach (var pair in On("Button_Ok_Fg", Text, "Button_Ok_Bg", "Button_Ok_Hover_Bg", "Button_Ok_Pressed_Bg")) yield return pair;
        foreach (var pair in On("Button_Cancel_Fg", Text, "Button_Cancel_Bg", "Button_Cancel_Hover_Bg", "Button_Cancel_Pressed_Bg")) yield return pair;
        foreach (var pair in On("Link_Fg", Text, "Bg", "Card_Bg", "Form_Bg", "Workbooks_Panel_Bg")) yield return pair;
        foreach (var pair in On("LinkHover_Fg", Text, "Bg")) yield return pair;
        foreach (var pair in On("Table_Header_Fg", Text, "Table_Header_Bg")) yield return pair;
        foreach (var pair in On("Table_Fg", Text, "Table_Bg", "Table_RowHover_Bg", "BoardTable_Added_Bg", "BoardTable_Modified_Bg", "BoardTable_Deleted_Bg")) yield return pair;

        foreach (string fore in new[] { "Text_Success_Fg", "Text_Fail_Fg", "Text_Waiting_Fg", "Text_Warning_Fg", "Text_Muted_Fg", "Text_Info_Fg", "Feedback_TrashIcon_Fg" })
            foreach (var pair in On(fore, Text, page)) yield return pair;

        foreach (var pair in On("Main_TabHeader_Drafts_Fg", Text, "Bg")) yield return pair;
        foreach (var pair in On("Workbooks_Muted_Fg", Text, "Bg", "Card_Bg", "Workbooks_Panel_Bg", "Workbooks_SelectedCard_Bg")) yield return pair;
        foreach (var pair in On("Workbooks_Faint_Fg", Text, "Bg", "Card_Bg", "Workbooks_Panel_Bg", "Workbooks_SelectedCard_Bg")) yield return pair;
        foreach (var pair in On("Chart_Label", Text, "Bg", "Card_Bg")) yield return pair;
        foreach (var pair in On("Changelog_Quote_Fg", Text, "Bg")) yield return pair;

        foreach (string alert in new[] { "Note", "Tip", "Important", "Warning", "Caution" })
            foreach (var pair in On($"Changelog_Alert_{alert}", Text, "Bg")) yield return pair;

        // A worklog's category or state is the text of a pill on Form_Bg and a card, and the fill under a badge's "#N" - whose
        // words are white or black by ThemeResources.TextOn ("best text", below).
        foreach (string worklog in new[] { "Worklog_Category_Note", "Worklog_Category_Cosmetic", "Worklog_Category_Issue", "Worklog_Status_Open", "Worklog_Status_Closed" })
            foreach (var pair in On(worklog, Text, "Form_Bg", "Card_Bg")) yield return pair;

        // A word on its own fill.
        foreach (string part in new[]
                 {
                     "Main_Banner", "Main_BannerButtonsHover", "Schematics_Name", "Schematics_Region_PAL", "Schematics_Region_NTSC",
                     "Schematics_ComponentHover", "Schematics_Panels", "Schematics_TemporaryBar", "Schematics_ComponentLabel", "Thumbnail",
                     "Thumbnail_Selection", "ComponentPopup_Name", "ComponentPopup_Pin", "ComponentPopup_OscilloscopeExpected",
                     "ComponentPopup_OscilloscopeSyncData", "ComponentPopup_KeyboardTyping", "Draft_Chip", "Draft_ReferenceBadge_Todo",
                     "Draft_ReferenceBadge_Used", "Workbooks_SearchHit", "Notice", "Badge_NewBoard", "Badge_Attention", "Badge_Count",
                     "Tour_Card", "BoardTable_Header", "BoardTable_ReplacePreview", "ReturnBar",
                     "Schematics_Region_Own1", "Schematics_Region_Own2", "Schematics_Region_Own3", "Schematics_Region_Own4"
                 })
        {
            yield return ($"{part}_Fg", $"{part}_Bg", Text);
        }

        yield return ("BoardTable_Error_Fg", "BoardTable_Error_Mark", Text);
        yield return ("BoardTable_Warning_Fg", "BoardTable_Warning_Mark", Text);

        // Edges and marks that must be seen: an input's edge, the selected tab's underline, the chart's bars, a frame.
        foreach (var pair in On("Form_Border", Edge, "Bg", "Form_Bg")) yield return pair;
        foreach (var pair in On("Main_TabUnderline_Selected", Edge, "Bg")) yield return pair;
        foreach (var pair in On("Chart_Bar", Edge, "Bg", "Card_Bg")) yield return pair;
        foreach (var pair in On("Badge_Attention_Bg", Edge, "Bg")) yield return pair;
        foreach (var pair in On("Mode_Selected_Border", Edge, "Bg")) yield return pair;
        foreach (var pair in On("BoardTable_CurrentCell_Border", Edge, "Table_Bg")) yield return pair;
    }

    // The fills a word is written on in white or black, whichever reads better (ThemeResources.TextOn).
    private static readonly string[] BestTextFills =
    [
        "Worklog_Category_Note", "Worklog_Category_Cosmetic", "Worklog_Category_Issue", "Worklog_Status_Open", "Worklog_Status_Closed",
        "Text_Success_Fg", "Text_Fail_Fg", "Text_Waiting_Fg", "Text_Warning_Fg"
    ];

    [Fact]
    public void Every_text_and_edge_reads_on_its_ground_in_every_theme()
    {
        var failures = new List<string>();

        UiTest.Run(() =>
        {
            foreach (string theme in ThemeContrastTests.Themes)
            {
                foreach ((string fore, string back, double least) in ThemeContrastTests.Pairs())
                {
                    string ground = ThemeContrastTests.Hex(theme, back);
                    double ratio = RegionColors.ContrastRatio(ThemeContrastTests.Hex(theme, fore, ground), ground);

                    if (ratio < least)
                        failures.Add($"{theme}: {fore} on {back} is {ratio.ToString("0.00", CultureInfo.InvariantCulture)}:1, wants {least.ToString(CultureInfo.InvariantCulture)}:1");
                }

                foreach (string fill in ThemeContrastTests.BestTextFills)
                {
                    string colour = ThemeContrastTests.Hex(theme, fill);
                    double ratio = RegionColors.ContrastRatio(RegionColors.ContrastText(colour), colour);

                    if (ratio < Text)
                        failures.Add($"{theme}: the best text on {fill} is {ratio.ToString("0.00", CultureInfo.InvariantCulture)}:1, wants {Text.ToString(CultureInfo.InvariantCulture)}:1");
                }
            }
        });

        Assert.True(failures.Count == 0, "Below WCAG's ratio:\n" + string.Join("\n", failures));
    }

    // The "User preference" theme seeds a color for every key it lists from these defaults, so each must be a color CRT
    // can read - a misspelt name would leave its key at the light theme's color, silently.
    [Fact]
    public void Every_user_preference_default_is_a_colour()
    {
        foreach ((string key, string value) in UserSettings.UserThemeDefaultColors)
        {
            Assert.True(Color.TryParse(value, out _), $"{key}: \"{value}\" is no color");
        }
    }

    // A theme's color for `key` as "#RRGGBB" - a translucent one laid over `under` (the theme's ground when none is given),
    // as it is drawn.
    private static string Hex(string theme, string key, string? under = null)
    {
        Color colour = ThemeContrastTests.Resolve(theme, key);

        if (colour.A < 255)
        {
            Color below = Color.Parse(under ?? ThemeContrastTests.Hex(theme, "Bg"));
            double a = colour.A / 255.0;
            colour = Color.FromRgb(
                (byte)System.Math.Round((colour.R * a) + (below.R * (1 - a))),
                (byte)System.Math.Round((colour.G * a) + (below.G * (1 - a))),
                (byte)System.Math.Round((colour.B * a) + (below.B * (1 - a))));
        }

        return $"#{colour.R:X2}{colour.G:X2}{colour.B:X2}";
    }

    private static Color Resolve(string theme, string key)
    {
        if (theme == "User preference" && UserSettings.UserThemeDefaultColors.TryGetValue(key, out string? own))
            return Color.Parse(own);

        ThemeVariant variant = theme == "Dark" ? ThemeVariant.Dark : ThemeVariant.Light;

        Assert.True(Application.Current!.TryGetResource(key, variant, out object? resource), $"{theme}: no resource {key}");

        return resource switch
        {
            ISolidColorBrush brush => brush.Color,
            Color colour => colour,
            _ => throw new Xunit.Sdk.XunitException($"{theme}: {key} is a {resource?.GetType().Name}, not a color")
        };
    }
}
