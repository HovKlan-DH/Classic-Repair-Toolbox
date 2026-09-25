using Avalonia;
using Avalonia.Controls;
using Avalonia.Media;

namespace Handlers.Theming
{
    // ###########################################################################################
    // Resolving an app resource that lives in App.axaml's ResourceDictionary.ThemeDictionaries.
    //
    // WHY THIS IS NOT A ONE-LINER, and why nothing should go back to a plain TryFindResource on a
    // control: a theme-variant-keyed resource is NOT resolved by the single-step lookup a control
    // instance offers. The lookup has to go through Application.Current with an EXPLICIT
    // ActualThemeVariant. Getting it wrong is completely silent - the caller's fallback renders and
    // nothing logs - which is exactly the shipped bug that once left the worklog bar's status dot
    // green while an entry's own pill showed the themed red.
    //
    // That two-step idiom had been copied into seven places, each carrying its own long comment
    // explaining the same thing, plus two inline copies with no fallback at all (which simply left a
    // Background unset on a miss). One implementation means the reasoning is recorded once and a
    // future copy has somewhere to be pointed at instead.
    //
    // Reads Application.Current but touches no control, so it works from a static helper and from a
    // headless test alike.
    // ###########################################################################################
    public static class ThemeResources
    {
        // ###########################################################################################
        // The resource for a key as T, or the given fallback when it is missing or of another type.
        //
        // A fallback is REQUIRED rather than optional-with-a-null-default on purpose: every caller
        // renders what comes back, and the two sites that had no fallback left a control's Background
        // unset on a miss, which is invisible until someone notices the panel is the wrong colour.
        // ###########################################################################################
        public static T Resolve<T>(string resourceKey, T fallback)
        {
            if (Application.Current == null || string.IsNullOrWhiteSpace(resourceKey))
            {
                return fallback;
            }

            if (Application.Current.TryGetResource(resourceKey, Application.Current.ActualThemeVariant, out var resource) &&
                resource is T typed)
            {
                return typed;
            }

            return fallback;
        }

        // ###########################################################################################
        // Same, but consulting a CONTROL's own resource scope first, so a resource set on that control
        // or an ancestor overrides the app-level one. The app lookup is still the fallback and still
        // does the whole job for a ThemeDictionaries key, which the control lookup alone does not.
        //
        // Three callers need this shape (Main's worklog bar, TabSchematics, the worklog entry editor)
        // and it is a real behavioural difference from Resolve, not a stylistic one - hence a second
        // method rather than folding the two together and quietly changing what three surfaces
        // resolve.
        // ###########################################################################################
        public static T ResolveForControl<T>(IResourceHost? host, string resourceKey, T fallback)
        {
            if (host != null &&
                host.TryFindResource(resourceKey, out var localResource) &&
                localResource is T localTyped)
            {
                return localTyped;
            }

            return Resolve(resourceKey, fallback);
        }

        // A themed brush, falling back to IndianRed - what every worklog colour resolver in this app
        // already falls back to when a key is somehow missing.
        public static IBrush ResolveBrush(string resourceKey) =>
            Resolve<IBrush>(resourceKey, Brushes.IndianRed);

        public static IBrush ResolveBrush(string resourceKey, IBrush fallback) =>
            Resolve(resourceKey, fallback);

        // The Color behind a themed brush, for the callers that need to build their own brushes from
        // it (a badge fill plus a matching outline, say) rather than share one instance.
        public static Color ResolveColor(string resourceKey, Color fallback) =>
            Resolve<IBrush>(resourceKey, null!) is ISolidColorBrush solid ? solid.Color : fallback;

        // ###########################################################################################
        // The two Font Awesome families, resolved from the same app resources the markup uses so a
        // glyph built in code is the same font as the identical glyph declared in XAML.
        //
        // Falls back to the default family, which renders a placeholder box rather than throwing.
        // Note is the one worklog category whose icon comes from the Regular weight; everything else
        // in the worklog UI is Solid.
        // ###########################################################################################
        public static FontFamily ResolveFontAwesomeSolid() =>
            Resolve("FontAwesomeSolid", FontFamily.Default);

        public static FontFamily ResolveFontAwesomeRegular() =>
            Resolve("FontAwesomeRegular", FontFamily.Default);
    }
}
