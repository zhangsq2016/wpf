using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security;
using System.Windows;
using System.Windows.Interop;
using System.Windows.Media;
using MS.Internal.PresentationCore.WindowsRuntime;

namespace MS.Internal.WindowsRuntime
{
    namespace Windows.UI.ViewManagement
    {
        internal class UISettings : IDisposable
        {
            private static readonly bool _isSupported;

            private UISettingsRCW.IUISettings3 _uisettings;

            private static Color _fallbackAccentColor = Color.FromArgb(0xff, 0x00, 0x78, 0xd4);

            private Color _accentColor, _accentLight1, _accentLight2, _accentLight3;
            private Color _accentDark1, _accentDark2, _accentDark3;

            private bool _useFallbackColor = false;

            static UISettings()
            {
                try
                {
                    _isSupported = true;
                    
                    if (GetWinRTInstance() == null)
                    {
                        _isSupported = false;
                    }
                }
                catch
                {
                    _isSupported = false;
                }
            }

            internal UISettings()
            {
                if (!_isSupported)
                {
                    throw new PlatformNotSupportedException();
                }

                try
                {
                    _uisettings = GetWinRTInstance() as UISettingsRCW.IUISettings3;
                }
                catch (COMException)
                {
                }

                if (_uisettings == null)
                {
                    throw new PlatformNotSupportedException();
                }
            }

            internal bool GetColorValue(UISettingsRCW.UIColorType desiredColor, out Color color)
            {
                bool result = false;

                try
                {
                    var uiColor = _uisettings.GetColorValue(desiredColor);
                    color = Color.FromArgb(uiColor.A, uiColor.R, uiColor.G, uiColor.B);
                    result = true;
                }
                catch (COMException)
                {
                    color = _fallbackAccentColor;
                }

                return result;
            }

            internal void TryUpdateAccentColors()
            {
                _useFallbackColor = true;
                try
                {
                    if(GetColorValue(UISettingsRCW.UIColorType.Accent, out Color systemAccent))
                    {
                        bool result = true;
                        if(_accentColor != systemAccent)
                        {
                            result &= GetColorValue(UISettingsRCW.UIColorType.AccentLight1, out _accentLight1);
                            result &= GetColorValue(UISettingsRCW.UIColorType.AccentLight2, out _accentLight2);
                            result &= GetColorValue(UISettingsRCW.UIColorType.AccentLight3, out _accentLight3);
                            result &= GetColorValue(UISettingsRCW.UIColorType.AccentDark1, out _accentDark1);
                            result &= GetColorValue(UISettingsRCW.UIColorType.AccentDark2, out _accentDark2);
                            result &= GetColorValue(UISettingsRCW.UIColorType.AccentDark3, out _accentDark3);
                            _accentColor = systemAccent;
                        }
                        _useFallbackColor = !result;
                    }
                }
                catch
                {
                }
            }

            private static object GetWinRTInstance()
            {
                object _winRtInstance = null;
                try
                {
                    _winRtInstance = UISettingsRCW.GetUISettingsInstance();
                }
                catch (Exception e) when (e is TypeLoadException || e is FileNotFoundException)
                {
                    _winRtInstance = null;
                }

                return _winRtInstance;
            }

            #region Color Properties

            internal Color AccentColor => _useFallbackColor ? _fallbackAccentColor : _accentColor;
            internal Color AccentLight1 => _useFallbackColor ? _fallbackAccentColor : _accentLight1;
            internal Color AccentLight2 => _useFallbackColor ? _fallbackAccentColor : _accentLight2;
            internal Color AccentLight3 => _useFallbackColor ? _fallbackAccentColor : _accentLight3;
            internal Color AccentDark1 => _useFallbackColor ? _fallbackAccentColor : _accentDark1;
            internal Color AccentDark2 => _useFallbackColor ? _fallbackAccentColor : _accentDark2;
            internal Color AccentDark3 => _useFallbackColor ? _fallbackAccentColor : _accentDark3;

            #endregion

            #region IDisposable

            bool _disposed = false;

            ~UISettings()
            {
                Dispose(false);
            }

            public void Dispose()
            {
                Dispose(true);
                GC.SuppressFinalize(this);
            }

            private void Dispose(bool disposing)
            {
                if (!_disposed)
                {
                    if (_uisettings != null)
                    {
                        try
                        {
                            // Release the input pane here
                            Marshal.ReleaseComObject(_uisettings);
                        }
                        catch
                        {
                            // Don't want to raise any exceptions in a finalizer, eat them here
                        }

                        _uisettings = null;
                    }

                    _disposed = true;
                }
            }

            #endregion
        }
    }
}