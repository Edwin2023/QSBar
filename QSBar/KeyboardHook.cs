using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Collections.Generic;

namespace QSBar
{
    public class KeyboardHook : IDisposable
    {
        private delegate IntPtr KeyboardProc(int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr SetWindowsHookEx(int idHook, KeyboardProc lpfn, IntPtr hMod, uint dwThreadId);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool UnhookWindowsHookEx(IntPtr hhk);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("kernel32.dll")]
        private static extern uint GetCurrentThreadId();

        [DllImport("user32.dll")]
        private static extern short GetAsyncKeyState(int vKey);

        private const int WH_KEYBOARD = 2; // 本地键盘钩子

        private KeyboardProc _proc;
        private IntPtr _hookID = IntPtr.Zero;
        private Dictionary<Keys, Action> _shortcuts = new Dictionary<Keys, Action>();
        private Dictionary<Keys, bool> _ctrlShortcuts = new Dictionary<Keys, bool>();
        private Dictionary<Keys, bool> _altShortcuts = new Dictionary<Keys, bool>();

        public KeyboardHook()
        {
            _proc = HookCallback;
            _hookID = SetWindowsHookEx(WH_KEYBOARD, _proc, IntPtr.Zero, GetCurrentThreadId());
        }

        public void AddShortcut(bool ctrl, bool alt, Keys key, Action action)
        {
            _shortcuts[key] = action;
            _ctrlShortcuts[key] = ctrl;
            _altShortcuts[key] = alt;
        }

        private IntPtr HookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            // nCode == 0 表示有按键动作
            // lParam 的第 31 位为 0 表示按下，为 1 表示释放
            if (nCode == 0)
            {
                uint lp = (uint)lParam.ToInt64();
                bool isKeyDown = (lp & 0x80000000) == 0;

                if (isKeyDown)
                {
                    Keys key = (Keys)wParam.ToInt32();
                    bool ctrlPressed = (GetAsyncKeyState(0x11) & 0x8000) != 0; // VK_CONTROL
                    bool altPressed = (GetAsyncKeyState(0x12) & 0x8000) != 0;  // VK_MENU

                    if (_shortcuts.ContainsKey(key))
                    {
                        if (ctrlPressed == _ctrlShortcuts[key] && altPressed == _altShortcuts[key])
                        {
                            Action action = _shortcuts[key];
                            // 在本地钩子中，我们直接在当前 UI 线程执行
                            try { action.Invoke(); } catch { }
                            return (IntPtr)1; // 拦截按键，防止传给 Excel
                        }
                    }
                }
            }
            return CallNextHookEx(_hookID, nCode, wParam, lParam);
        }

        public void Dispose()
        {
            if (_hookID != IntPtr.Zero)
            {
                UnhookWindowsHookEx(_hookID);
                _hookID = IntPtr.Zero;
            }
        }
    }
}
