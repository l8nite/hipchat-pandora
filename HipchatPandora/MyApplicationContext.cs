using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Timers;
using System.Windows.Forms;
using Common;
using HipChat.Net;
using HipChat.Net.Http;
using HipchatPandora.Properties;
using Timer = System.Timers.Timer;

namespace HipchatPandora
{
    internal class MyApplicationContext : ApplicationContext
    {
        private String _currentSong = null;
        private readonly NotifyIcon _trayIcon;
        private Timer _timer;
        private HipChatClient client;

        public MyApplicationContext()
        {
             client = new HipChatClient(new ApiConnection(new Credentials("")));
             _trayIcon = new NotifyIcon
             {
                 Icon = Resources.AppIcon,
                 ContextMenu = new ContextMenu(new[] { new MenuItem("Exit", Exit) }),
                 Visible = true
             };
            
            // Ping(null, null);

            StartTimer();
        }

        [DllImport("user32.dll")]
        public static extern int SendMessage(IntPtr hWnd, int wMsg, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr FindWindowEx(IntPtr hWndParent, IntPtr hWndChildAfter, string lpClassName, string lpWindowName);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        private static IntPtr GetSystemTrayHandle()
        {
            IntPtr hWndTray = FindWindow("Shell_TrayWnd", null);
            if (hWndTray != IntPtr.Zero)
            {
                hWndTray = FindWindowEx(hWndTray, IntPtr.Zero, "TrayNotifyWnd", null);
                if (hWndTray != IntPtr.Zero)
                {
                    hWndTray = FindWindowEx(hWndTray, IntPtr.Zero, "SysPager", null);
                    if (hWndTray != IntPtr.Zero)
                    {
                        hWndTray = FindWindowEx(hWndTray, IntPtr.Zero, "ToolbarWindow32", null);
                        return hWndTray;
                    }
                }
            }

            return IntPtr.Zero;
        }

        private void StartTimer()
        {
            _timer = new Timer(5000)
                {
                    Enabled = true, 
                    AutoReset = true
                };
            _timer.Elapsed += Ping;
            _timer.Start();
        }

        private void Exit(object sender, EventArgs e)
        {
            // hide this so the user doesn't have to mouse over it
            _trayIcon.Visible = false;
            Application.Exit();
        }

        private void Ping(object sender, ElapsedEventArgs e)
        {
            // var process = Process.GetProcessesByName("Pandora").First();
            var _ToolbarWindowHandle = GetSystemTrayHandle();

            uint count = User32.SendMessage(_ToolbarWindowHandle, TB.BUTTONCOUNT, 0, 0);

            for (int i = 0; i < count; i++)
            {
                TBBUTTON tbButton = new TBBUTTON();
                string text = string.Empty;
                IntPtr ipWindowHandle = IntPtr.Zero;

                bool b = GetTBButton(_ToolbarWindowHandle, i, ref tbButton, ref text, ref ipWindowHandle);

                if (text.StartsWith("Pandora"))
                {
                    UpdateCurrentSong(text);
                }
            }
        }

        private void UpdateCurrentSong(string text)
        {
            if (_currentSong == null || _currentSong != text)
            {
                _currentSong = text;
                var parts = text.Split(new[] { '\n' });
                var song = parts[1];
                var artist = parts[2];
                client.Rooms.SendNotificationAsync("140763", "Now playing: " + song + " " + artist).Wait();
            }
        }

        // cobbled together from http://stackoverflow.com/questions/6366505/get-tooltip-text-from-icon-in-system-tray
        // and http://www.codeproject.com/Articles/10497/A-tool-to-order-the-window-buttons-in-your-taskbar
        private unsafe bool GetTBButton(IntPtr hToolbar, int i, ref TBBUTTON tbButton, ref string text, ref IntPtr ipWindowHandle)
        {
            // One page
            const int BUFFER_SIZE = 0x1000;

            byte[] localBuffer = new byte[BUFFER_SIZE];

            uint processId = 0;
            uint threadId = User32.GetWindowThreadProcessId(hToolbar, out processId);

            IntPtr hProcess = Kernel32.OpenProcess(ProcessRights.ALL_ACCESS, false, processId);
            if (hProcess == IntPtr.Zero)
            {
                Debug.Assert(false);
                return false;
            }

            IntPtr ipRemoteBuffer = Kernel32.VirtualAllocEx(
                hProcess, 
                IntPtr.Zero, 
                new UIntPtr(BUFFER_SIZE), 
                MemAllocationType.COMMIT, 
                MemoryProtection.PAGE_READWRITE);

            if (ipRemoteBuffer == IntPtr.Zero)
            {
                Debug.Assert(false);
                return false;
            }

            // TBButton
            fixed (TBBUTTON* pTBButton = &tbButton)
            {
                IntPtr ipTBButton = new IntPtr(pTBButton);

                int b = (int)User32.SendMessage(hToolbar, TB.GETBUTTON, (IntPtr)i, ipRemoteBuffer);
                if (b == 0)
                {
                    Debug.Assert(false);
                    return false;
                }

                // this is fixed
                int dwBytesRead = 0;
                IntPtr ipBytesRead = new IntPtr(&dwBytesRead);

                bool b2 = Kernel32.ReadProcessMemory(hProcess, ipRemoteBuffer, ipTBButton, new UIntPtr((uint)sizeof(TBBUTTON)), ipBytesRead);

                if (!b2)
                {
                    Debug.Assert(false);
                    return false;
                }
            }

            // button text
            fixed (byte* pLocalBuffer = localBuffer)
            {
                IntPtr ipLocalBuffer = new IntPtr(pLocalBuffer);

                int chars = (int)User32.SendMessage(hToolbar, TB.GETBUTTONTEXTW, (IntPtr)tbButton.idCommand, ipRemoteBuffer);
                if (chars == -1)
                {
                    Debug.Assert(false);
                    return false;
                }

                // this is fixed
                int dwBytesRead = 0;
                IntPtr ipBytesRead = new IntPtr(&dwBytesRead);

                bool b4 = Kernel32.ReadProcessMemory(hProcess, ipRemoteBuffer, ipLocalBuffer, new UIntPtr(BUFFER_SIZE), ipBytesRead);

                if (!b4)
                {
                    Debug.Assert(false);
                    return false;
                }

                text = Marshal.PtrToStringUni(ipLocalBuffer, chars);

                if (text == " ")
                {
                    text = string.Empty;
                }
            }

            // window handle
            fixed (byte* pLocalBuffer = localBuffer)
            {
                IntPtr ipLocalBuffer = new IntPtr(pLocalBuffer);

                // this is in the remote virtual memory space
                IntPtr ipRemoteData = new IntPtr(tbButton.dwData);

                // this is fixed
                int dwBytesRead = 0;
                IntPtr ipBytesRead = new IntPtr(&dwBytesRead);

                bool b4 = Kernel32.ReadProcessMemory(hProcess, ipRemoteData, ipLocalBuffer, new UIntPtr(4), ipBytesRead);

                if (!b4)
                {
                    return false;
                }

                // Debug.Assert(false); return false; }
                if (dwBytesRead != 4)
                {
                    Debug.Assert(false);
                    return false;
                }

                int iWindowHandle = BitConverter.ToInt32(localBuffer, 0);
                if (iWindowHandle == -1)
                {
                    Debug.Assert(false);
                }
                // return false; }

                ipWindowHandle = new IntPtr(iWindowHandle);
            }

            Kernel32.VirtualFreeEx(hProcess, ipRemoteBuffer, UIntPtr.Zero, MemAllocationType.RELEASE);

            Kernel32.CloseHandle(hProcess);

            return true;
        }
    }
}