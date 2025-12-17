using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Forms;

namespace KakaotalkBot
{
    public class WindowsMacro
    {
        private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);
        [DllImport("user32.dll")]
        private static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);
        [DllImport("user32.dll")]
        private static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);
        [DllImport("user32.dll")]
        private static extern int GetWindowTextLength(IntPtr hWnd);
        [DllImport("user32.dll")]
        private static extern bool IsWindowVisible(IntPtr hWnd);
        [DllImport("user32.dll")]
        static extern bool SetCursorPos(int X, int Y);
        [DllImport("user32.dll")]
        static extern void mouse_event(uint dwFlags, uint dx, uint dy, uint dwData, UIntPtr dwExtraInfo);
        [DllImport("user32.dll")]
        static extern void keybd_event(
    byte bVk,
    byte bScan,
    uint dwFlags,
    UIntPtr dwExtraInfo


);

        [DllImport("user32.dll")]
        static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

        [StructLayout(LayoutKind.Sequential)]
        struct RECT
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }

        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);
        [DllImport("user32.dll")]
        public static extern bool PostMessage(IntPtr hWnd, uint Msg, int wParam, int lParam);

        [DllImport("user32.dll")]
        private static extern IntPtr FindWindow(string lpClassName, string lpWindowName);
        [DllImport("user32.dll")]
        private static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter, string lpszClass, string lpszWindow);
        [DllImport("user32.dll")]
        public static extern bool SendMessage(IntPtr hWnd, uint Msg, int wParam, string lParam);
        [DllImport("user32.dll")]
        static extern uint SendInput(uint nInputs, INPUT[] pInputs, int cbSize);
        [StructLayout(LayoutKind.Sequential)]
        struct INPUT
        {
            public int type;
            public KEYBDINPUT ki;
        }

        [StructLayout(LayoutKind.Sequential)]
        struct KEYBDINPUT
        {
            public ushort wVk;
            public ushort wScan;
            public uint dwFlags;
            public uint time;
            public IntPtr dwExtraInfo;
        }


        private static WindowsMacro instance;
        public static WindowsMacro Instance
        {
            get
            {
                if (instance == null)
                {
                    instance = new WindowsMacro();
                }
                return instance;
            }
        }

        public Form Form { get; set; }

        private const uint GMEM_MOVEABLE = 0x0002;

        const uint MOUSEEVENTF_LEFTDOWN = 0x02;
        const uint MOUSEEVENTF_LEFTUP = 0x04;
        const uint MOUSEEVENTF_RIGHTDOWN = 0x0008;
        const uint MOUSEEVENTF_RIGHTUP = 0x0010;

        // 핫키 관련 상수
        private const int WM_HOTKEY = 0x0312;
        private const int MOD_ALT = 0x0001;
        private const int MOD_CONTROL = 0x0002;
        private const int MOD_SHIFT = 0x0004;

        const int WM_CLOSE = 0x0010;
        const int WM_SETTEXT = 0x000C;
        const int WM_KEYDOWN = 0x0100;
        const int WM_KEYUP = 0x0101;
        const int VK_RETURN = 0x0D;
        const int VK_SPACE = 0x20;

        private const int WM_USER = 0x0400;
        private const int EM_SETTEXTEX = WM_USER + 97;

        private const int ST_DEFAULT = 0x0000;
        private const int ST_KEEPUNDO = 0x0001;

        private const int INPUT_KEYBOARD = 1;
        private const uint KEYEVENTF_KEYUP = 0x0002;
        private const uint KEYEVENTF_UNICODE = 0x0004;

        private const ushort VK_CONTROL = 0x11;
        const byte VK_A = 0x41;
        private const ushort VK_V = 0x56;
        private const byte VK_ESCAPE = 0x1B;
        const byte VK_BACK = 0x08;


        private WindowsMacro() { }

        public void ClickLeft()
        {
            mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, UIntPtr.Zero);
            mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, UIntPtr.Zero);
        }
        public void ClickRight()
        {
            mouse_event(MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, UIntPtr.Zero);
            mouse_event(MOUSEEVENTF_RIGHTUP, 0, 0, 0, UIntPtr.Zero);
        }
        public void SendReturn(IntPtr hwnd)
        {
            SendKeys.SendWait("~"); // ^ == Ctrl
        }

        public void SendCtrlKey(IntPtr hwnd, char key)
        {
            SetForegroundWindow(hwnd);
            SendKeys.SendWait("^" + key); // ^ == Ctrl
        }

        public void SendInput(ushort virtualKey)
        {
            INPUT[] inputs = new INPUT[2];

            inputs[0] = new INPUT
            {
                type = INPUT_KEYBOARD,
                ki = new KEYBDINPUT { wVk = virtualKey }
            };

            inputs[1] = new INPUT
            {
                type = INPUT_KEYBOARD,
                ki = new KEYBDINPUT
                {
                    wVk = virtualKey,
                    dwFlags = KEYEVENTF_KEYUP
                }
            };

            SendInput((uint)inputs.Length, inputs, Marshal.SizeOf(typeof(INPUT)));
        }

        public static void Paste(string text)
        {
            Clipboard.SetText(text);

            var inputs = new[]
            {
            new INPUT{ type=INPUT_KEYBOARD, ki=new KEYBDINPUT{ wVk=VK_CONTROL } },
            new INPUT{ type=INPUT_KEYBOARD, ki=new KEYBDINPUT{ wVk=VK_V } },
            new INPUT{ type=INPUT_KEYBOARD, ki=new KEYBDINPUT{ wVk=VK_V, dwFlags=KEYEVENTF_KEYUP } },
            new INPUT{ type=INPUT_KEYBOARD, ki=new KEYBDINPUT{ wVk=VK_CONTROL, dwFlags=KEYEVENTF_KEYUP } },
            };

            SendInput((uint)inputs.Length, inputs, Marshal.SizeOf(typeof(INPUT)));
        }

        static void KeyDown(byte vk)
        {
            keybd_event(vk, 0, 0, UIntPtr.Zero);
        }

        static void KeyUp(byte vk)
        {
            keybd_event(vk, 0, KEYEVENTF_KEYUP, UIntPtr.Zero);
        }

        static void Press(byte vk)
        {
            KeyDown(vk);
            KeyUp(vk);
        }

        public static void SendText(string text)
        {
            Clipboard.SetText(text);

            KeyDown((byte)VK_CONTROL);
            Press((byte)VK_V);
            KeyUp((byte)VK_CONTROL);

            Press(VK_RETURN);
        }

        public static void CloseWindow(IntPtr hwnd)
        {
            if (hwnd != IntPtr.Zero)
            {
                PostMessage(hwnd, WM_CLOSE, 0, 0);
            }
        }

        public static void LaunchKakaoTalk()
        {
            string kakaoPath = @"C:\Program Files (x86)\Kakao\KakaoTalk\KakaoTalk.exe";

            try
            {
                Process.Start(kakaoPath);
            }
            catch (Exception ex)
            {
                MessageBox.Show("카카오톡 실행 실패: " + ex.Message);
            }
        }

        public List<WindowInfo> GetWindowList()
        {
            List<WindowInfo> windowTitles = new List<WindowInfo>();

            EnumWindows((hWnd, lParam) =>
            {
                if (IsWindowVisible(hWnd))
                {
                    int length = GetWindowTextLength(hWnd);
                    if (length > 0)
                    {
                        StringBuilder builder = new StringBuilder(length + 1);
                        GetWindowText(hWnd, builder, builder.Capacity);
                        string title = builder.ToString();

                        WindowInfo windowInfo = new WindowInfo();
                        windowInfo.Title = title;
                        windowInfo.Handle = hWnd;
                        windowTitles.Add(windowInfo);
                    }
                }
                return true; // 계속 열거
            }, IntPtr.Zero);

            //Console.WriteLine("현재 열려 있는 윈도우 목록:");
            //foreach (string title in windowTitles)
            //{
            //    Console.WriteLine($"- {title}");
            //}

            return windowTitles;
        }

        public IntPtr FindVoiceRoomWindow()
        {
            IntPtr handle = IntPtr.Zero;
            StringBuilder sb = new StringBuilder();

            EnumWindows((hWnd, lParam) =>
            {
                if (IsWindowVisible(hWnd))
                {
                    int length = GetWindowTextLength(hWnd);
                    if (length > 0)
                    {
                        StringBuilder builder = new StringBuilder(length + 1);
                        GetWindowText(hWnd, builder, builder.Capacity);
                        string title = builder.ToString();
                        if (title.StartsWith("보이스룸: "))
                        {
                            handle = hWnd;
                        }
                    }
                }
                return true; // 계속 열거
            }, IntPtr.Zero);

            return handle;
        }

        public void OpenChatRoom(string roomName)
        {
            // 1. 카카오톡 메인 창 찾기
            bool alreadyOpen = false;
            IntPtr hwndKakao = FindWindow(null, "카카오톡");
            if (hwndKakao == IntPtr.Zero)
            {
                alreadyOpen = false;
                LaunchKakaoTalk();
            }
            else
            {
                alreadyOpen = true;
            }

            while(FindWindow(null, "카카오톡")  == IntPtr.Zero)
            {
                Thread.Sleep(1000);
            }

            if(alreadyOpen == false)
            {
                Thread.Sleep(20000);
                hwndKakao = FindWindow(null, "카카오톡");
            }

            Point p = GetWindowPos(hwndKakao);
            SetCursorPos(p.X + 32, p.Y + 120);
            ClickLeft();
            Thread.Sleep(1000);

            // 2. 검색 Edit 컨트롤 찾아 들어가기
            IntPtr hwndEdit1 = FindWindowEx(hwndKakao, IntPtr.Zero, "EVA_ChildWindow", null);
            IntPtr hwndEdit2_1 = FindWindowEx(hwndEdit1, IntPtr.Zero, "EVA_Window", null);
            IntPtr hwndEdit2_2 = FindWindowEx(hwndEdit1, hwndEdit2_1, "EVA_Window", null);
            IntPtr hwndEdit3 = FindWindowEx(hwndEdit2_2, IntPtr.Zero, "Edit", null); // 최종 Edit 컨트롤

            if (hwndEdit3 == IntPtr.Zero) return;

            SetForegroundWindow(hwndEdit3);
            Thread.Sleep(50);
            // TODO: 글자 수만큼 해야함.
            Press(VK_BACK);
            Thread.Sleep(100);
            Press(VK_BACK);
            Thread.Sleep(100);
            Press(VK_BACK);

            SendText(roomName);
            Thread.Sleep(1000);

            SetCursorPos(p.X + 195, p.Y + 120);
            ClickLeft();
            Thread.Sleep(50);
            ClickLeft();

            Thread.Sleep(1000);

            CloseChatRoom(roomName);

            SetCursorPos(p.X + 195, p.Y + 120);
            ClickLeft();
            Thread.Sleep(50);
            ClickLeft();

        }

        public void CloseWindow(string windowName)
        {
            IntPtr hWnd = FindWindow(null, windowName);

            if (hWnd != IntPtr.Zero)
            {
                PostMessage(hWnd, WM_CLOSE, 0, 0);
            }
        }

        public void CloseChatRoom(string roomName)
        {
            IntPtr hwndMain = FindWindow(null, roomName);
            if (hwndMain != IntPtr.Zero)
            {
                Point roomPos = GetWindowPos(hwndMain);
                Point roomSize = GetWindowSize(hwndMain);
                SetCursorPos(roomPos.X + roomSize.X - 12, roomPos.Y + 12);
                ClickLeft();
            }
            
        }

        public string CopyChatroomText(IntPtr hwndMain)
        {
            IntPtr hwndList = FindWindowEx(hwndMain, IntPtr.Zero, "EVA_VH_ListControl_Dblclk", null);

            if (hwndList == IntPtr.Zero)
            {
                return "";
            }

            // 채팅 전체 선택 후 복사 (Ctrl+A → Ctrl+C)
            SendCtrlKey(hwndList, 'A');
            Thread.Sleep(100);
            SendCtrlKey(hwndList, 'c');
            Thread.Sleep(200);

            string text = string.Empty;
            try
            {
                text = Clipboard.GetText();

            }
            catch (Exception e)
            {

            }

            return text;
        }

        public void SendTextToChatroom(string chatroomName, string message)
        {
            //soliloquyTimer.Reset();

            IntPtr hwndMain = FindWindow(null, chatroomName);
            IntPtr hwndEdit = FindWindowEx(hwndMain, IntPtr.Zero, "RichEdit50W", null);

            try
            {
                Clipboard.SetText(message);

                SetForegroundWindow(hwndMain);
                Thread.Sleep(100);
                SendCtrlKey(hwndEdit, 'v');

                SendReturn(hwndEdit);
                Thread.Sleep(300);
            }
            catch (Exception e)
            {
            }

        }

        public Point GetWindowPos(IntPtr hwnd)
        {
            RECT r;
            GetWindowRect(hwnd, out r);
            Point point = new Point();
            point.X = r.Left;
            point.Y = r.Top;
            return point;
        }

        public Point GetWindowSize(IntPtr hwnd)
        {
            RECT r;
            GetWindowRect(hwnd, out r);
            Point point = new Point();
            point.X = r.Right - r.Left;
            point.Y = r.Bottom - r.Top;
            return point;
        }
    }
}
