using Standard;
using System.Windows.Appearance;
using System.Windows.Media;
using Microsoft.Win32;
using System.Collections;
using System.Collections.Generic;
using System.Windows.Interop;

namespace System.Windows;

internal static class ThemeManager
{

    #region Constructor

    static ThemeManager()
    {
        _cachedThemeDictionaryUris = new List<Uri>();

        if (Application.Current != null)
        {
            foreach (ResourceDictionary mergedDictionary in Application.Current.Resources.MergedDictionaries)
            {
                if (mergedDictionary.Source != null && mergedDictionary.Source.ToString().EndsWith("FluentWindows.xaml"))
                {
                    _isFluentWindowsThemeEnabled = true;
                    break;
                }
            }
        }

        _currentApplicationTheme = GetSystemTheme();
    }

    #endregion

    #region Internal Methods

    internal static void ApplySystemTheme(bool forceUpdate = false)
    {
        string systemTheme = GetSystemTheme();
        bool useLightMode = GetUseLightTheme();
        Color systemAccentColor = DwmColorization.GetSystemAccentColor();

        var windows = Application.Current?.Windows;
        if(windows != null)
        {
            ApplyTheme(windows , systemTheme, useLightMode, systemAccentColor, forceUpdate);
        }
    }

    internal static void ApplySystemTheme(Window window, bool forceUpdate = false)
    {
        string systemTheme = GetSystemTheme();
        bool useLightMode = GetUseLightTheme();
        Color systemAccentColor = DwmColorization.GetSystemAccentColor();
        ApplyTheme(new ArrayList { window } , systemTheme, useLightMode, systemAccentColor, forceUpdate);
    }

    internal static void ApplyTheme(
        IEnumerable windows, 
        string requestedTheme, 
        bool requestedUseLightMode,
        Color requestedAccentColor, 
        bool forceUpdate = false)
    {
        bool needsUpdate = forceUpdate;
        
        if(DwmColorization.GetSystemAccentColor() != DwmColorization.CurrentApplicationAccentColor)
        {
            DwmColorization.UpdateAccentColors();
            needsUpdate = true;
        }

        if(needsUpdate || requestedTheme != _currentApplicationTheme || requestedUseLightMode != _currentUseLightMode)
        {
            Uri dictionaryUri = GetFluentWindowThemeResourceUri(requestedTheme, requestedUseLightMode, out ApplicationTheme applicationTheme);
            var backdropType = applicationTheme == ApplicationTheme.HighContrast ? WindowBackdropType.None : WindowBackdropType.MainWindow;

            UpdateFluentWindowsThemeResources(dictionaryUri);

            foreach(Window window in windows)
            {
                SetImmersiveDarkMode(window, !requestedUseLightMode);
                window.WindowBackdropType = backdropType;
            }

            _currentApplicationTheme = requestedTheme;
            _currentUseLightMode = requestedUseLightMode;
        }
    }

    private static bool SetImmersiveDarkMode(Window window, bool useDarkMode)
    {
        if (window == null)
        {
            return false;
        }

        IntPtr handle = new WindowInteropHelper(window).Handle;

        if (handle != IntPtr.Zero)
        {
            var dwmResult = NativeMethods.DwmSetWindowAttributeUseImmersiveDarkMode(handle, useDarkMode);
            return dwmResult == HRESULT.S_OK;
        }

        return false;
    }

    #region Helper Methods

    internal static string GetSystemTheme()
    {
        string systemTheme = Registry.GetValue(_regThemeKeyPath,
            "CurrentTheme", "aero.theme") as string ?? String.Empty;

        return systemTheme;
    }
   
    internal static bool GetUseLightTheme()
    {
        var appsUseLightTheme = Registry.GetValue(_regPersonalizeKeyPath,
            "AppsUseLightTheme", null) as int?;

        if (appsUseLightTheme == null)
        {

            // Slight deviation from Harshit's code
            return Registry.GetValue(_regPersonalizeKeyPath,
                "SystemUsesLightTheme", null) as int? == 0 ? false : true;
        }

        // Slight deviation from Harshit's code
        return appsUseLightTheme != 0;
    }

    internal static void UpdateFluentWindowsThemeResources(Uri dictionaryUri)
    {
        ArgumentNullException.ThrowIfNull(dictionaryUri, nameof(dictionaryUri));

        var newDictionary = new ResourceDictionary() { Source = dictionaryUri };

        foreach (var key in newDictionary.Keys)
        {
            if (Application.Current.Resources.Contains(key))
            {
                if (!object.Equals(Application.Current.Resources[key], newDictionary[key]))
                {
                    Application.Current.Resources[key] = newDictionary[key];
                }
            }
            else
            {
                Application.Current.Resources.Add(key, newDictionary[key]);
            }
        }
    }

    internal static void AddFluentWindowsThemeDictionary(ResourceDictionary dictionary)
    {
        if(IsFluentWindowsThemeEnabled && !_cachedThemeDictionaryUris.Contains(dictionary.Source))
        {
            Application.Current.Resources.MergedDictionaries.Add(dictionary);
            _cachedThemeDictionaryUris.Add(dictionary.Source);
        }
    }

    #endregion

    #endregion

    #region Internal Properties

    internal static bool IsFluentWindowsThemeEnabled
    {
        get
        {
            return _isFluentWindowsThemeEnabled;
        }
    }

    #endregion

    #region Private Methods

    private static Uri GetFluentWindowThemeResourceUri(
        string systemTheme, bool useLightMode,
        out ApplicationTheme applicationTheme
    )
    {
        string themeColorFileName = "light.xaml";

        if(SystemParameters.HighContrast)
        {
            applicationTheme = ApplicationTheme.HighContrast;
            themeColorFileName = systemTheme switch
            {
                string s when s.Contains("hcblack") => "hcblack.xaml",
                string s when s.Contains("hcwhite") => "hcwhite.xaml",
                string s when s.Contains("hc1") => "hc1.xaml",
                _ => "hc2.xaml"
            };

        }
        else 
        {
            applicationTheme = useLightMode ? ApplicationTheme.Light : ApplicationTheme.Dark;
            themeColorFileName = useLightMode ? "light.xaml" : "dark.xaml";
        }

        return new Uri("pack://application:,,,/PresentationFramework.FluentWindows;component/Resources/Theme/" + themeColorFileName, UriKind.Absolute);
    }

    #endregion

    #region Private Members

    private static readonly string _regThemeKeyPath = "HKEY_CURRENT_USER\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Themes";

    private static readonly string _regPersonalizeKeyPath = "HKEY_CURRENT_USER\\Software\\Microsoft\\Windows\\CurrentVersion\\Themes\\Personalize";

    private static string _currentApplicationTheme;

    private static bool _currentUseLightMode = true;

    private static bool _isFluentWindowsThemeEnabled = false;

    private static ICollection<Uri> _cachedThemeDictionaryUris;

    #endregion
}