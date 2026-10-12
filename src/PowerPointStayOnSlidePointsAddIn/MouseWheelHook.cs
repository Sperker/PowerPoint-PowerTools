using System;
using System.Drawing;
using System.Runtime.InteropServices;

namespace PowerPointStayOnSlidePointsAddIn
{
    internal sealed class MouseWheelHook : IDisposable
    {
        private IntPtr _hookHandle = IntPtr.Zero;
        private NativeMethods.LowLevelMouseProc _proc;

        public bool IsStarted
        {
            get { return _hookHandle != IntPtr.Zero; }
        }

        public int LastStartError { get; private set; }

        public event EventHandler<GlobalMouseWheelEventArgs> MouseWheel;
        public event EventHandler<GlobalMouseClickEventArgs> LeftButtonDown;

        public void Start()
        {
            if (_hookHandle != IntPtr.Zero)
                return;

            LastStartError = 0;
            _proc = HookCallback;

            _hookHandle = NativeMethods.SetWindowsHookEx(
                NativeMethods.WH_MOUSE_LL,
                _proc,
                IntPtr.Zero,
                0
            );

            if (_hookHandle == IntPtr.Zero)
            {
                LastStartError = Marshal.GetLastWin32Error();
            }
        }

        public void Dispose()
        {
            if (_hookHandle != IntPtr.Zero)
            {
                NativeMethods.UnhookWindowsHookEx(_hookHandle);
                _hookHandle = IntPtr.Zero;
            }
        }

        private IntPtr HookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            if (nCode >= 0)
            {
                int msg = wParam.ToInt32();
                if (msg == NativeMethods.WM_MOUSEWHEEL)
                {
                    var hookStruct = Marshal.PtrToStructure<NativeMethods.MSLLHOOKSTRUCT>(lParam);
                    short wheelDelta = (short)((hookStruct.mouseData >> 16) & 0xffff);

                    var args = new GlobalMouseWheelEventArgs(
                        new Point(hookStruct.pt.X, hookStruct.pt.Y),
                        wheelDelta
                    );

                    var handler = MouseWheel;
                    if (handler != null)
                    {
                        handler(this, args);
                    }

                    if (args.Handled)
                    {
                        return new IntPtr(1);
                    }
                }
                else if (msg == NativeMethods.WM_LBUTTONDOWN)
                {
                    var hookStruct = Marshal.PtrToStructure<NativeMethods.MSLLHOOKSTRUCT>(lParam);
                    var args = new GlobalMouseClickEventArgs(
                        new Point(hookStruct.pt.X, hookStruct.pt.Y)
                    );

                    var handler = LeftButtonDown;
                    if (handler != null)
                    {
                        handler(this, args);
                    }
                }
            }

            return NativeMethods.CallNextHookEx(_hookHandle, nCode, wParam, lParam);
        }
    }

    internal sealed class GlobalMouseWheelEventArgs : EventArgs
    {
        public Point ScreenPoint { get; private set; }
        public int Delta { get; private set; }
        public bool Handled { get; set; }

        public GlobalMouseWheelEventArgs(Point screenPoint, int delta)
        {
            ScreenPoint = screenPoint;
            Delta = delta;
            Handled = false;
        }
    }

    internal sealed class GlobalMouseClickEventArgs : EventArgs
    {
        public Point ScreenPoint { get; private set; }

        public GlobalMouseClickEventArgs(Point screenPoint)
        {
            ScreenPoint = screenPoint;
        }
    }
}
