using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Text;

namespace PowerPointStayOnSlidePointsAddIn
{
    internal static class NativeMethods
    {
        public const int WH_MOUSE_LL = 14;
        public const int WH_KEYBOARD_LL = 13;

        public const int WM_MOUSEWHEEL = 0x020A;
        public const int WM_LBUTTONDOWN = 0x0201;

        public const int WM_KEYDOWN = 0x0100;
        public const int WM_KEYUP = 0x0101;
        public const int WM_SYSKEYDOWN = 0x0104;
        public const int WM_SYSKEYUP = 0x0105;

        public const int VK_PRIOR = 0x21; // Page Up
        public const int VK_NEXT = 0x22;  // Page Down
        public const int VK_END = 0x23;
        public const int VK_HOME = 0x24;
        public const int VK_LEFT = 0x25;
        public const int VK_UP = 0x26;
        public const int VK_RIGHT = 0x27;
        public const int VK_DOWN = 0x28;

        public const int VK_SHIFT = 0x10;
        public const int VK_LSHIFT = 0xA0;
        public const int VK_RSHIFT = 0xA1;

        public const int VK_CONTROL = 0x11;
        public const int VK_LCONTROL = 0xA2;
        public const int VK_RCONTROL = 0xA3;

        public const int VK_MENU = 0x12; // Alt
        public const int VK_LMENU = 0xA4;
        public const int VK_RMENU = 0xA5;

        public const uint GA_ROOT = 2;

        public delegate IntPtr LowLevelMouseProc(int nCode, IntPtr wParam, IntPtr lParam);
        public delegate IntPtr LowLevelKeyboardProc(int nCode, IntPtr wParam, IntPtr lParam);

        [StructLayout(LayoutKind.Sequential)]
        public struct POINT
        {
            public int X;
            public int Y;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct RECT
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct MSLLHOOKSTRUCT
        {
            public POINT pt;
            public uint mouseData;
            public uint flags;
            public uint time;
            public IntPtr dwExtraInfo;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct KBDLLHOOKSTRUCT
        {
            public uint vkCode;
            public uint scanCode;
            public uint flags;
            public uint time;
            public IntPtr dwExtraInfo;
        }

        [DllImport("user32.dll", SetLastError = true)]
        public static extern IntPtr SetWindowsHookEx(
            int idHook,
            LowLevelMouseProc lpfn,
            IntPtr hMod,
            uint dwThreadId);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern IntPtr SetWindowsHookEx(
            int idHook,
            LowLevelKeyboardProc lpfn,
            IntPtr hMod,
            uint dwThreadId);

        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool UnhookWindowsHookEx(IntPtr hhk);

        [DllImport("user32.dll")]
        public static extern IntPtr CallNextHookEx(
            IntPtr hhk,
            int nCode,
            IntPtr wParam,
            IntPtr lParam);

        [DllImport("user32.dll")]
        public static extern short GetAsyncKeyState(int vKey);

        [DllImport("user32.dll")]
        public static extern bool GetCursorPos(out POINT lpPoint);

        [DllImport("user32.dll")]
        public static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);

        [DllImport("user32.dll")]
        public static extern IntPtr WindowFromPoint(POINT point);

        [DllImport("user32.dll")]
        public static extern IntPtr GetAncestor(IntPtr hwnd, uint gaFlags);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern IntPtr FindWindowEx(IntPtr parentHandle, IntPtr childAfter, string className, string windowTitle);

        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool GetWindowRect(IntPtr hwnd, out RECT lpRect);

        public static bool IsKeyDown(int virtualKey)
        {
            return (GetAsyncKeyState(virtualKey) & 0x8000) != 0;
        }

        public static bool IsCtrlDown()
        {
            return IsKeyDown(VK_CONTROL) || IsKeyDown(VK_LCONTROL) || IsKeyDown(VK_RCONTROL);
        }

        public static bool IsShiftDown()
        {
            return IsKeyDown(VK_SHIFT) || IsKeyDown(VK_LSHIFT) || IsKeyDown(VK_RSHIFT);
        }

        public static bool IsAltDown()
        {
            return IsKeyDown(VK_MENU) || IsKeyDown(VK_LMENU) || IsKeyDown(VK_RMENU);
        }

        public static Point GetCursorPosition()
        {
            GetCursorPos(out POINT p);
            return new Point(p.X, p.Y);
        }

        public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool EnumChildWindows(IntPtr hwnd, EnumWindowsProc func, IntPtr lParam);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        public static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

        public static bool TryGetMdiClassRect(IntPtr pptHwnd, out RECT rect)
        {
            rect = default;
            if (pptHwnd == IntPtr.Zero) return false;

            // 1. Try legacy MDIClient -> mdiClass
            IntPtr mdiClient = FindWindowEx(pptHwnd, IntPtr.Zero, "MDIClient", null);
            if (mdiClient != IntPtr.Zero)
            {
                IntPtr mdiClass = FindWindowEx(mdiClient, IntPtr.Zero, "mdiClass", null);
                if (mdiClass != IntPtr.Zero && GetWindowRect(mdiClass, out rect))
                    return true;
            }

            // 2. Try direct mdiClass
            IntPtr directMdi = FindWindowEx(pptHwnd, IntPtr.Zero, "mdiClass", null);
            if (directMdi != IntPtr.Zero && GetWindowRect(directMdi, out rect))
                return true;

            // 3. Search children for mdiClass or paneClassDC
            IntPtr found = IntPtr.Zero;
            try
            {
                EnumChildWindows(pptHwnd, (h, l) =>
                {
                    StringBuilder sb = new StringBuilder(64);
                    GetClassName(h, sb, 64);
                    string cls = sb.ToString();
                    if (cls.Equals("mdiClass", StringComparison.OrdinalIgnoreCase) ||
                        cls.Equals("paneClassDC", StringComparison.OrdinalIgnoreCase))
                    {
                        found = h;
                        return false; // Stop enumeration
                    }
                    return true;
                }, IntPtr.Zero);
            }
            catch { }

            if (found != IntPtr.Zero && GetWindowRect(found, out rect))
                return true;

            // 4. Reliable fallback: PowerPoint frame window rect with standard chrome offsets
            RECT appRect;
            if (GetWindowRect(pptHwnd, out appRect))
            {
                rect = new RECT
                {
                    Left = appRect.Left + 250,
                    Top = appRect.Top + 140,
                    Right = appRect.Right - 20,
                    Bottom = appRect.Bottom - 45
                };
                return true;
            }

            return false;
        }
    }
}
