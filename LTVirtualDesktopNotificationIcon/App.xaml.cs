using System.Configuration;
using System.Data;
using System.Windows;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Timers;
using System.Windows.Interop;
using System.Windows.Threading;

namespace LTVirtualDesktopNotificationIcon
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    /// 

    public partial class App : Application
    {
        NotifyIconData data = new NotifyIconData();
        List<nint> icons = new List<nint>();
        HwndSource source;
        uint newMenuItemID = 2;

        DispatcherTimer timer = new DispatcherTimer();

        const int WM_COMMAND = 0x0111;
        const int WM_RBUTTONUP = 0x0205;

        long closeMenuItemID = 0;

        const int callbackMessageId = 0x8001;

        protected override void OnStartup(StartupEventArgs e)
        {
            for (int i = 1; i < 10; i++)
            {
                string path = Environment.CurrentDirectory + "\\icons\\num_" + i.ToString() + ".ico";
                icons.Add(LoadImageW(0, path, 1, 128, 128, 0x00000010));
            }

            VirtualDesktopManager manager = new VirtualDesktopManager();

            data.cbSize = (uint)Marshal.SizeOf(data);

            MainWindow window = new MainWindow();
            var helper = new WindowInteropHelper(window);
            var handle = helper.EnsureHandle();

            source = HwndSource.FromHwnd(handle);
            source.AddHook(WndProc);

            data.WindowHandle = handle;
            data.IconID = 0;
            data.CallbackMessageId = callbackMessageId;
            data.VersionOrTimeout = 0x4;

            data.IconHandle = icons[manager.GetCurrentDesktopindex()];

            data.IconState = IconState.Visible;
            data.StateMask = IconState.Hidden;

            data.Flags = IconDataFlags.Message | IconDataFlags.Icon | IconDataFlags.Tip;

            data.BalloonFlags = 0;
            data.ToolTipText = "LT Virtual Desktop Notification";
            data.BalloonText = data.BalloonTitle = string.Empty;

            if (!Shell_NotifyIconW(NotifyIconMsgType.Add, ref data))
            {
                Console.WriteLine("ERROR: Cant add system tray icon");
                return;
            }

            timer.Interval = new TimeSpan(0, 0, 0, 0, 33);
            timer.Tick += (object? sender, EventArgs e) =>
            {
                int index = manager.GetCurrentDesktopindex();
                data.IconHandle = icons[index];

                if (!Shell_NotifyIconW(NotifyIconMsgType.Modify, ref data))
                {
                    Console.WriteLine("ERROR: Cant modify system tray icon");
                    return;
                }
            };
            timer.Start();
        }

        protected override void OnExit(ExitEventArgs e)
        {
            timer.Stop();
            source.RemoveHook(WndProc);
            if (!Shell_NotifyIconW(NotifyIconMsgType.Delete, ref data))
            {
                Console.WriteLine("ERROR: Cant close system tray icon");
                return;
            }
        }

        private IntPtr WndProc(nint hwnd, int msg, IntPtr wParam, IntPtr lParam, ref bool handled)
        {
            switch (msg)
            {
                case WM_COMMAND:
                    {
                        long itemID = (long)wParam;
                        if (itemID == closeMenuItemID)
                        {
                            handled = true;
                            Current.Shutdown();
                        }
                    }
                    break;

                case callbackMessageId:
                    switch (lParam)
                    {
                        case WM_RBUTTONUP:
                            handled = true;
                            WinPoint point = new WinPoint();
                            if (!GetCursorPos(ref point))
                            {
                                Console.WriteLine("ERROR: Cant get cursor position");
                            }

                            var menu = CreatePopupMenu();

                            if (!InsertMenuW(menu, 0, 0x00000400, ref newMenuItemID, "Close"))
                            {
                                Console.WriteLine("ERROR: Cant insert menu item");
                            }

                            closeMenuItemID = GetMenuItemID(menu, 0);

                            if (!SetForegroundWindow(source.Handle))
                            {
                                Console.WriteLine("ERROR: Cant set foregorund window");
                            }

                            Rect rect;
                            if (!TrackPopupMenu(menu, 0x0002, point.x, point.y, 0, hwnd, ref rect))
                            {
                                Console.WriteLine("ERROR: Cant track popup menu");
                            }

                            if (!DestroyMenu(menu))
                            {
                                Console.WriteLine("ERROR: Cant destroy menu");
                            }
                            break;
                    }
                    break;
                default:
                    return DefWindowProcW(hwnd, msg, wParam, lParam); 
            }

            return IntPtr.Zero;
        }

        [DllImport("user32.dll", EntryPoint = "DefWindowProcW")]
        public static extern IntPtr DefWindowProcW(IntPtr hWnd, int Msg, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll", EntryPoint = "GetMenuItemID")]
        public static extern long GetMenuItemID(IntPtr hMenu, int nPos);

        [DllImport("user32.dll", EntryPoint = "SetForegroundWindow")]
        public static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll", EntryPoint = "CreatePopupMenu")]
        public static extern IntPtr CreatePopupMenu();

        [DllImport("user32.dll", EntryPoint = "TrackPopupMenu")]
        public static extern bool TrackPopupMenu(IntPtr hMenu, uint uFlags, int x, int y, int nReserved, IntPtr hWnd, ref Rect prcRect);

        [DllImport("user32.dll", EntryPoint = "InsertMenuW", CharSet = CharSet.Unicode)]
        public static extern bool InsertMenuW(IntPtr hmenu, uint uPosition, uint flags, ref uint uIDNewItem, string lpNewItem);

        [DllImport("user32.dll", EntryPoint = "DestroyMenu")]
        public static extern bool DestroyMenu(IntPtr hMenu);


        [StructLayout(LayoutKind.Sequential)]
        public struct WinPoint
        {
            public int x;
            public int y;
        }


        [DllImport("user32.dll", EntryPoint = "GetCursorPos")]
        public static extern bool GetCursorPos(ref WinPoint point);

        [DllImport("user32.dll", EntryPoint = "LoadImageW", CharSet = CharSet.Unicode)]
        public static extern IntPtr LoadImageW(IntPtr hInst, string name, uint type, int cx, int cy, uint fuLoad);

        [DllImport("shell32.dll", EntryPoint = "Shell_NotifyIconW", CharSet = CharSet.Unicode)]
        public static extern bool Shell_NotifyIconW(NotifyIconMsgType dwMessage, ref NotifyIconData lpData);

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        public struct NotifyIconData
        {
            public uint cbSize;
            public IntPtr WindowHandle;
            public uint IconID;
            public IconDataFlags Flags;
            public uint CallbackMessageId;
            public IntPtr IconHandle;

            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 128)]
            public string ToolTipText;
            public IconState IconState;
            public IconState StateMask;

            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 256)]
            public string BalloonText;
            public uint VersionOrTimeout;

            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 64)]
            public string BalloonTitle;
            public int BalloonFlags;
            public Guid TaskbarIconGuid;
            public IntPtr BalloonIconHandle;
        }

        public enum NotifyIconMsgType
        {
            Add = 0x00,
            Modify = 0x01,
            Delete = 0x02,
            SetFocus = 0x03,
            SetVersion = 0x04
        }

        [Flags]
        public enum IconDataFlags
        {
            Message = 0x01,
            Icon = 0x02,
            Tip = 0x04,
            State = 0x08,
            Info = 0x10,
            Realtime = 0x40,
            UseLegacyToolTips = 0x80
        }

        public enum IconState
        {
            Visible = 0x00,
            Hidden = 0x01,
            // unuseed
            //Shared = 0x02
        }
    }
}
