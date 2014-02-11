using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.Windows;
using System.Windows.Interop;

namespace PowerpointJabber
{
    class WindowsInteropFunctions
    {
        #region user32.dll imports
        private delegate bool EnumDelegate(IntPtr hWnd, int lParam);
        public const int SW_HIDE = 0,
            SW_SHOWNORMAL = 1,
            SW_NORMAL = 1,
            SW_SHOWMINIMIZED = 2,
            SW_SHOWMAXIMIZED = 3,
            SW_MAXIMIZE = 3,
            SW_SHOWNOACTIVATE = 4,
            SW_SHOW = 5,
            SW_MINIMIZE = 6,
            SW_SHOWMINNOACTIVE = 7,
            SW_SHOWNA = 8,
            SW_RESTORE = 9,
            SW_SHOWDEFAULT = 10,
            SW_MAX = 10;
        [StructLayout(LayoutKind.Sequential)]
        public struct POINT
        {
            public int X;
            public int Y;

            public POINT(int x, int y)
            {
                this.X = x;
                this.Y = y;
            }
        }
        [StructLayout(LayoutKind.Sequential)]
        public struct WINDOWPLACEMENT
        {
            public int length;
            public int flags;
            public int showCmd;
            public POINT minPosition;
            public POINT maxPosition;
            public RECT normalPosition;
        }

        [DllImport("user32.dll")]
        private static extern bool GetWindowPlacement(IntPtr hWnd, out WINDOWPLACEMENT lpwndpl);

        [StructLayout(LayoutKind.Sequential)]
        public struct RECT
        {
            public int left;
            public int top;
            public int right;
            public int bottom;
        }

        [DllImport("user32.dll", EntryPoint = "EnumDesktopWindows", ExactSpelling = false, CharSet = CharSet.Auto, SetLastError = true)]
        private static extern bool _EnumDesktopWindows(IntPtr hDesktop, EnumDelegate lpEnumCallbackFunction, IntPtr lParam);

        [DllImport("user32.dll", EntryPoint = "GetWindowText", ExactSpelling = false, CharSet = CharSet.Auto, SetLastError = true)]
        private static extern int _GetWindowText(IntPtr hWnd, StringBuilder lpWindowText, int nMaxCount);

        [DllImport("user32.dll")]
        private static extern int GetWindowRect(int hwnd, ref RECT rc);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern bool IsWindowVisible(IntPtr hWnd);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool GetWindowRect(HandleRef hwnd, out RECT lpRect);

        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll", EntryPoint = "SystemParametersInfo")]
        private static extern bool SystemParametersInfo(System.UInt32 uiAction, System.UInt32 uiParam, System.UInt32 pvParam, System.UInt32 fWinIni);

        [DllImport("user32.dll", EntryPoint = "SetForegroundWindow")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("User32.dll", EntryPoint = "SetActiveWindow")]
        private static extern void SetActiveWindow(int hWnd);

        // Find window by Caption only. Note you must pass IntPtr.Zero as the first parameter.
        [DllImport("user32.dll", EntryPoint = "FindWindow", SetLastError = true)]
        private static extern IntPtr FindWindowByCaption(IntPtr ZeroOnly, string lpWindowName);

        private static IntPtr FindWindowByCaption(string lpWindowName)
        {
            return FindWindowByCaption(IntPtr.Zero, lpWindowName);
        }
        #endregion

        public static bool presenterActive
        {
            get
            {
                if ((int)presenterWindow > 0)
                    return true;
                else return false;
            }
        }
        public static string GetWindowText(IntPtr hWnd)
        {
            StringBuilder title = new StringBuilder(255);
            int titleLength = _GetWindowText(hWnd, title, title.Capacity + 1);
            title.Length = titleLength;
            return title.ToString();
        }
        public static void switchToMeTL()
        {
            Dictionary<IntPtr, string> mTitlesList = new Dictionary<IntPtr, string>();
            EnumDelegate enumfunc = (EnumDelegate)delegate(IntPtr hWnd, int lParam)
            {
                string title = GetWindowText(hWnd);
                mTitlesList.Add(hWnd, title);
                return true;
            };
            IntPtr hDesktop = IntPtr.Zero; // current desktop
            bool success = _EnumDesktopWindows(hDesktop, enumfunc, IntPtr.Zero);
            int successFrequency = 0;
            if (success)
            {
                foreach (var KV in mTitlesList)
                {
                    if (KV.Value.StartsWith("MeTL") || KV.Value.EndsWith("- MeTL"))
                    {
                        BringWindowToFront(KV.Key);
                        successFrequency++;
                    }
                }
                if (successFrequency == 0)
                    System.Diagnostics.Process.Start("iexplore.exe", "-extoff http://metl.adm.monash.edu.au/MeTL2011/MeTL%20Presenter.application");
            }
            else
                System.Diagnostics.Process.Start("iexplore.exe", "-extoff http://metl.adm.monash.edu.au/MeTL2011/MeTL%20Presenter.application");
        }
        public static IntPtr presenterWindow
        {
            get
            {
                try
                {
                    return FindWindowByCaption("PowerPoint Presenter View - [" + ThisAddIn.instance.Application.ActivePresentation.Windows[1].Caption + "]");
                }
                catch (Exception)
                {
                    return (IntPtr)0;
                }
            }
        }
        public struct WindowStateData
        {
            public bool isVisible;
            public double X;
            public double Y;
            public double Height;
        }
        private static IntPtr currentWindow()
        {
            try
            {
                return presenterActive ? presenterWindow : (IntPtr)ThisAddIn.instance.Application.ActivePresentation.SlideShowWindow.HWND;
            }
            catch (Exception) { return (IntPtr)0; }
        }
        public static WindowStateData getAppropriateViewData()
        {
            var window = currentWindow();
           //
            var stateData = new WindowStateData();
            var placementData = new WINDOWPLACEMENT();
            GetWindowPlacement(window, out placementData);
            stateData.isVisible = (isWindowFocused(window) || (ThisAddIn.instance != null && ThisAddIn.instance.SSSW != null && ThisAddIn.instance.SSSW.HWND != null && isWindowFocused(ThisAddIn.instance.SSSW.HWND)));
            RECT rect = new RECT();
            if (placementData.showCmd != SW_SHOWMAXIMIZED)
            {
                GetWindowRect((int)window, ref rect);
                stateData.X = rect.left;
                stateData.Y = rect.top;
                stateData.Height = rect.bottom - rect.top;
            }
            else
            {
                stateData.X = 0;
                stateData.Y = 0;
                stateData.Height = System.Windows.Forms.Screen.FromHandle(window).WorkingArea.Height;  
            }
            return stateData;
        }
        public static void BringAppropriateViewToFront()
        {
            if (presenterActive)
                BringWindowToFront(presenterWindow);
            else
                BringWindowToFront((IntPtr)ThisAddIn.instance.Application.ActivePresentation.SlideShowWindow.HWND);
        }
        private static void BringWindowToFront(IntPtr windowHandle)
        {
            System.UInt32 oldUILockout = 0x0000;
            SystemParametersInfo((System.UInt32)0x2000, 0, 0, oldUILockout);
            SystemParametersInfo((System.UInt32)0x2001, 0, 0, 0x0000);
            SetForegroundWindow(windowHandle);
            SetActiveWindow((int)windowHandle);
            SystemParametersInfo((System.UInt32)0x2001, 0, 0, oldUILockout);
        }
        private static bool isWindowFocused(IntPtr windowHandle)
        {
            return (GetForegroundWindow() == windowHandle);
        }
    }
}
