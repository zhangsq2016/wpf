using System.Runtime.InteropServices;
using System.Windows.Interop;
using System.Windows.Media;
using MS.Internal;
using Standard;
using HRESULT = Standard.HRESULT;

// ReSharper disable once CheckNamespace
namespace System.Windows.Appearance;

internal static class WindowBackdropManager
{
    internal static bool IsSupported(WindowBackdropType backdropType)
    {
        return backdropType switch
        {
            WindowBackdropType.Auto => Utility.IsOSWindows11Insider1OrNewer,
            WindowBackdropType.TabbedWindow => Utility.IsOSWindows11Insider1OrNewer,
            WindowBackdropType.MainWindow => Utility.IsOSWindows11OrNewer,
            WindowBackdropType.TransientWindow => Utility.IsOSWindows7OrNewer,
            WindowBackdropType.None => true,
            _ => false
        };
    }

    internal static bool SetBackdrop(Window window, WindowBackdropType backdropType)
    {
        if (window is null || 
                !IsSupported(backdropType) || 
                window.AllowsTransparency ||
                IsBackdropEnabled == false)
        {
            return false;
        }

        var handle = new WindowInteropHelper(window).Handle;
        if(handle == IntPtr.Zero)
        {
            return false;
        }

        return SetBackdropCore(handle, backdropType);
    }

    #region Private Methods

    private static bool SetBackdropCore(IntPtr hwnd, WindowBackdropType backdropType)
    {
        ArgumentNullException.ThrowIfNull(hwnd, nameof(hwnd));

        if (backdropType == WindowBackdropType.None)
        {
            RestoreBackground(hwnd);
            return RemoveBackdrop(hwnd);
        }

        RemoveBackground(hwnd);
        return ApplyBackdrop(hwnd, backdropType);
    }

    private static bool ApplyBackdrop(IntPtr hwnd, WindowBackdropType backdropType)
    {
        UpdateGlassFrame(hwnd, backdropType);
        
        var backdropPvAttribute = backdropType switch
        {
            WindowBackdropType.Auto => Standard.DWMSBT.DWMSBT_TABBEDWINDOW,
            WindowBackdropType.TabbedWindow => Standard.DWMSBT.DWMSBT_TABBEDWINDOW,
            WindowBackdropType.MainWindow => Standard.DWMSBT.DWMSBT_MAINWINDOW,
            WindowBackdropType.TransientWindow => Standard.DWMSBT.DWMSBT_TRANSIENTWINDOW,
            _ => Standard.DWMSBT.DWMSBT_NONE
        };

        var dwmResult = NativeMethods.DwmSetWindowAttributeSystemBackdropType(hwnd, backdropPvAttribute);
        return dwmResult == HRESULT.S_OK;
    }

    private static bool RemoveBackdrop(IntPtr hWnd)
    {
        UpdateGlassFrame(hWnd, WindowBackdropType.None);

        var backdropPvAttribute = Standard.DWMSBT.DWMSBT_NONE;
        var dwmResult = NativeMethods.DwmSetWindowAttributeSystemBackdropType(hWnd, backdropPvAttribute);
        return dwmResult == HRESULT.S_OK;
    }

    private static bool RemoveBackground(IntPtr hwnd)
    {
        var windowSource = HwndSource.FromHwnd(hwnd);
        if(windowSource?.Handle != IntPtr.Zero && windowSource.CompositionTarget != null)
        {
            windowSource.CompositionTarget.BackgroundColor = Colors.Transparent;
            return true;
        }
        return false;
    }

    private static bool RestoreBackground(IntPtr hwnd)
    {
        var windowSource = HwndSource.FromHwnd(hwnd);
        if(windowSource?.Handle != IntPtr.Zero && windowSource.CompositionTarget != null)
        {
            windowSource.CompositionTarget.BackgroundColor = SystemColors.WindowColor;
            return true;
        }
        return false;
    }

    private static bool UpdateGlassFrame(IntPtr hWnd, WindowBackdropType backdropType)
    {
        MARGINS margins = new MARGINS();
        if(backdropType != WindowBackdropType.None)
        {
            margins = new MARGINS { cxLeftWidth = -1, cxRightWidth = -1, cyTopHeight = -1, cyBottomHeight = -1 };                    
        }

        var dwmApiResult = NativeMethods.DwmExtendFrameIntoClientArea(hWnd, ref margins);
        return new HRESULT((uint)dwmApiResult) == HRESULT.S_OK;
    }

    #endregion

    #region Internal Properties

    internal static bool IsBackdropEnabled
    {
        get
        {
            if (_isBackdropEnabled == null)
            {
                _isBackdropEnabled = true;
                
                if (FrameworkAppContextSwitches.DisableFluentWindowsThemeWindowBackdrop || !Utility.IsOSWindows11Insider1OrNewer || !ThemeManager.IsFluentWindowsThemeEnabled)
                {
                    _isBackdropEnabled = false;
                }
            }

            return (bool)_isBackdropEnabled;
        }
    }

    private static bool? _isBackdropEnabled = null;

    #endregion

}