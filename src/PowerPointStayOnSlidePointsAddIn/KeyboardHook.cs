using System;
using System.Runtime.InteropServices;

namespace PowerPointStayOnSlidePointsAddIn
{
    internal sealed class KeyboardHook : IDisposable
    {
        private IntPtr _hookHandle = IntPtr.Zero;
        private NativeMethods.LowLevelKeyboardProc _proc;

        public bool IsStarted
        {
            get { return _hookHandle != IntPtr.Zero; }
        }

        public int LastStartError { get; private set; }

        public event EventHandler<GlobalKeyEventArgs> KeyDown;

        public void Start()
        {
            if (_hookHandle != IntPtr.Zero)
                return;

            LastStartError = 0;
            _proc = HookCallback;

            _hookHandle = NativeMethods.SetWindowsHookEx(
                NativeMethods.WH_KEYBOARD_LL,
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
                if (msg == NativeMethods.WM_KEYDOWN || msg == NativeMethods.WM_SYSKEYDOWN)
                {
                    var hookStruct = Marshal.PtrToStructure<NativeMethods.KBDLLHOOKSTRUCT>(lParam);

                    var args = new GlobalKeyEventArgs((int)hookStruct.vkCode);
                    var handler = KeyDown;
                    if (handler != null)
                    {
                        handler(this, args);
                    }

                    if (args.Handled)
                    {
                        return new IntPtr(1);
                    }
                }
            }

            return NativeMethods.CallNextHookEx(_hookHandle, nCode, wParam, lParam);
        }
    }

    internal sealed class GlobalKeyEventArgs : EventArgs
    {
        public int VirtualKeyCode { get; private set; }
        public bool Handled { get; set; }

        public GlobalKeyEventArgs(int virtualKeyCode)
        {
            VirtualKeyCode = virtualKeyCode;
            Handled = false;
        }
    }
}
