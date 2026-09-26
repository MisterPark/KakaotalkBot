using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Automation;
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
        static extern bool GetCursorPos(out POINT lpPoint);
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
        [DllImport("user32.dll")]
        static extern IntPtr GetDC(IntPtr hwnd);

        [DllImport("user32.dll")]
        static extern int ReleaseDC(IntPtr hwnd, IntPtr hdc);

        [DllImport("gdi32.dll")]
        static extern uint GetPixel(IntPtr hdc, int x, int y);


        struct POINT
        {
            public int X;
            public int Y;
        }

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
            if (!StaInputWorker.Instance.IsCurrent) { StaInputWorker.Instance.Invoke(() => ClickLeft()); return; }
            mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, UIntPtr.Zero);
            mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, UIntPtr.Zero);
        }
        public void ClickRight()
        {
            if (!StaInputWorker.Instance.IsCurrent) { StaInputWorker.Instance.Invoke(() => ClickRight()); return; }
            mouse_event(MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, UIntPtr.Zero);
            mouse_event(MOUSEEVENTF_RIGHTUP, 0, 0, 0, UIntPtr.Zero);
        }
        public void SendReturn()
        {
            if (!StaInputWorker.Instance.IsCurrent) { StaInputWorker.Instance.Invoke(() => SendReturn()); return; }
            SendKeys.SendWait("~"); // ^ == Ctrl
        }

        public void SendCtrlKey(IntPtr hwnd, char key)
        {
            if (!StaInputWorker.Instance.IsCurrent) { StaInputWorker.Instance.Invoke(() => SendCtrlKey(hwnd, key)); return; }
            SetForegroundWindow(hwnd);
            SendKeys.SendWait("^" + key); // ^ == Ctrl
        }

        public void SendInput(ushort virtualKey)
        {
            if (!StaInputWorker.Instance.IsCurrent) { StaInputWorker.Instance.Invoke(() => SendInput(virtualKey)); return; }
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
            if (!StaInputWorker.Instance.IsCurrent) { StaInputWorker.Instance.Invoke(() => Paste(text)); return; }
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
            if (!StaInputWorker.Instance.IsCurrent) { StaInputWorker.Instance.Invoke(() => SendText(text)); return; }
            Clipboard.SetText(text);

            KeyDown((byte)VK_CONTROL);
            Press((byte)VK_V);
            KeyUp((byte)VK_CONTROL);

            Press(VK_RETURN);
        }

        public static void CloseWindow(IntPtr hwnd)
        {
            if (!StaInputWorker.Instance.IsCurrent) { StaInputWorker.Instance.Invoke(() => CloseWindow(hwnd)); return; }
            if (hwnd != IntPtr.Zero)
            {
                PostMessage(hwnd, WM_CLOSE, 0, 0);
            }
        }

        public static void LaunchKakaoTalk()
        {
            string kakaoPath = @"C:\Program Files\Kakao\KakaoTalk\KakaoTalk.exe";

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

        public IntPtr FindTargetWindow(string target)
        {
            if (string.IsNullOrEmpty(target)) return IntPtr.Zero;

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
                        if (title.StartsWith(target))
                        {
                            handle = hWnd;
                        }
                    }
                }
                return true; // 계속 열거
            }, IntPtr.Zero);

            return handle;
        }

        public bool IsKakaoTalkOpen()
        {
            IntPtr hwndKakao = FindWindow(null, "카카오톡");

            return hwndKakao != IntPtr.Zero;
        }

        public bool IsChatRoomOpen(string roomName)
        {
            List<WindowInfo> list = GetWindowList();
            foreach (WindowInfo windowInfo in list)
            {
                if (windowInfo.Title == roomName)
                {
                    return true;
                }
            }

            return false;
        }

        [DllImport("user32.dll")]
        private static extern IntPtr GetAncestor(IntPtr window, uint flags);
        [DllImport("user32.dll")]
        private static extern bool IsWindow(IntPtr window);

        internal string LastRecycleError { get; private set; }
        private bool DeferRecycle(string reason) { LastRecycleError = reason; return false; }

        internal bool RecycleChatRoom(string roomName, int expectedPid)
        {
            LastRecycleError = null;
            bool recycled;
            string error;
            bool completed = StaInputWorker.Instance.TryInvoke(() =>
            {
                if (IsChatRoomOpen(roomName))
                {
                    IntPtr input = KakaoInputMentions.EmptyInputWindow(roomName, expectedPid);
                    if (input == IntPtr.Zero) return DeferRecycle("작성 중인 내용이 있어 재열기를 보류했습니다.");
                    IntPtr room = GetAncestor(input, 2);
                    if (room == IntPtr.Zero) return DeferRecycle("채팅창을 확인하지 못했습니다.");
                    CloseWindow(room);
                    var closing = Stopwatch.StartNew();
                    while (IsWindow(room))
                    {
                        if (closing.ElapsedMilliseconds >= 3000) return DeferRecycle("채팅창 닫기 시간 초과");
                        Thread.Sleep(50);
                    }
                }
                OpenChatRoom(roomName);
                var opening = Stopwatch.StartNew();
                while (!IsChatRoomOpen(roomName))
                {
                    if (opening.ElapsedMilliseconds >= 3000) return DeferRecycle("채팅창 재열기 시간 초과");
                    Thread.Sleep(50);
                }
                // 이름이 같은 다른 계정 창으로 연결되지 않았는지도 다시 검사합니다.
                KakaoInputMentions.EmptyInputWindow(roomName, expectedPid);
                return true;
            }, out recycled, out error);
            if (!completed) LastRecycleError = error;
            return completed && recycled;
        }

        private DateTime nextKakaoLaunch;

        public void OpenChatRoom(string roomName)
        {
            if (!StaInputWorker.Instance.IsCurrent) { StaInputWorker.Instance.Invoke(() => OpenChatRoom(roomName)); return; }
            // 실행·로그인 대기를 STA 안에서 하지 않습니다. 봇의 다음 열기 시도에서 다시 확인합니다.
            IntPtr hwndKakao = FindWindow(null, "카카오톡");
            if (hwndKakao == IntPtr.Zero)
            {
                if (DateTime.UtcNow >= nextKakaoLaunch)
                {
                    nextKakaoLaunch = DateTime.UtcNow.AddSeconds(30);
                    LaunchKakaoTalk();
                }
                throw new InvalidOperationException("카카오톡 실행·로그인 대기 중입니다. 다음 주기에 다시 확인합니다.");
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

            var opened = Stopwatch.StartNew();
            while (!IsChatRoomOpen(roomName) && opened.ElapsedMilliseconds < 1000) Thread.Sleep(50);
        }

        public void CloseWindow(string windowName)
        {
            if (!StaInputWorker.Instance.IsCurrent) { StaInputWorker.Instance.Invoke(() => CloseWindow(windowName)); return; }
            IntPtr hWnd = FindWindow(null, windowName);

            if (hWnd != IntPtr.Zero)
            {
                PostMessage(hWnd, WM_CLOSE, 0, 0);
            }
        }

        public void CloseChatRoom(string roomName)
        {
            if (!StaInputWorker.Instance.IsCurrent) { StaInputWorker.Instance.Invoke(() => CloseChatRoom(roomName)); return; }
            IntPtr hwndMain = FindWindow(null, roomName);
            if (hwndMain != IntPtr.Zero)
            {
                Point roomPos = GetWindowPos(hwndMain);
                Point roomSize = GetWindowSize(hwndMain);
                SetCursorPos(roomPos.X + roomSize.X - 12, roomPos.Y + 12);
                ClickLeft();
            }

        }

        public void CloseChatRoom(IntPtr room)
        {
            if (!StaInputWorker.Instance.IsCurrent) { StaInputWorker.Instance.Invoke(() => CloseChatRoom(room)); return; }
            if (room != IntPtr.Zero)
            {
                Point roomPos = GetWindowPos(room);
                Point roomSize = GetWindowSize(room);
                SetCursorPos(roomPos.X + roomSize.X - 12, roomPos.Y + 12);
                ClickLeft();
            }
        }

        public string CopyChatroomText(IntPtr hwndMain)
        {
            if (!StaInputWorker.Instance.IsCurrent) return StaInputWorker.Instance.Invoke(() => CopyChatroomText(hwndMain));
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

        private void RunChatSend(Action action)
        { StaInputWorker.Instance.Invoke(action); }

        internal void SendOperatorNotice(string room, int pid, string text)
        {
            string body = Database.Instance.ForChat(text);
            RunChatSend(() => new NativeMentionInput(room, pid).SendPlain(body));
        }

        /// <summary>닉네임으로 개인 멘션 객체를 만들고 ID로 대상을 지정하여 본문과 한 번에 전송한다.</summary>
        public void SendMentionToChatroom(string chatroomName, long userId, string candidateNickname, string message, int expectedProcessId = 0, IEnumerable<string> fallbackNicknames = null)
        {
            string body = Database.Instance.ForChat(message);
            var candidates = fallbackNicknames == null ? null : new List<string>(fallbackNicknames);
            RunChatSend(() => new KakaoMentionSender(chatroomName, expectedProcessId).Send(userId, candidateNickname, body, candidates));
        }

        public void SendTextToChatroom(string chatroomName, string message)
        {
            message = Database.Instance.ForChat(message);
            RunChatSend(() =>
            {
                IntPtr hwndMain = FindWindow(null, chatroomName);
                if (hwndMain == IntPtr.Zero) return;
                IntPtr hwndEdit = FindWindowEx(hwndMain, IntPtr.Zero, "RichEdit50W", null);
                if (hwndEdit == IntPtr.Zero) return;
                try
                {
                    Clipboard.SetText(message);
                    SetForegroundWindow(hwndMain);
                    Thread.Sleep(100);
                    SendCtrlKey(hwndEdit, 'v');
                    SendReturn();
                    Thread.Sleep(300);
                }
                catch (Exception)
                {
                }
            });
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

        public static Color GetScreenPixelColor(int x, int y)
        {
            IntPtr hdc = GetDC(IntPtr.Zero);
            uint pixel = GetPixel(hdc, x, y);
            ReleaseDC(IntPtr.Zero, hdc);

            return Color.FromArgb(
                (int)(pixel & 0x000000FF),
                (int)(pixel & 0x0000FF00) >> 8,
                (int)(pixel & 0x00FF0000) >> 16
            );
        }

        public Point GetCursorPos()
        {
            GetCursorPos(out POINT p);
            return new Point(p.X, p.Y);
        }

        public void SetCursor(int x, int y)
        {
            if (!StaInputWorker.Instance.IsCurrent) { StaInputWorker.Instance.Invoke(() => SetCursor(x, y)); return; }
            SetCursorPos(x, y);
        }

        public bool IsKakaoTalkChatRoom(IntPtr hwnd)
        {
            IntPtr hwnd1 = FindWindowEx(hwnd, IntPtr.Zero, "RICHEDIT50W", null);
            if (hwnd1 == IntPtr.Zero) return false;
            IntPtr hwnd2 = FindWindowEx(hwnd, IntPtr.Zero, "EVA_VH_ListControl_Dblclk", null);
            if (hwnd2 == IntPtr.Zero) return false;
            IntPtr hwnd3 = FindWindowEx(hwnd, IntPtr.Zero, "Edit", null);
            if (hwnd3 == IntPtr.Zero) return false;

            return true;
        }

        public List<WindowInfo> FindAllKakaoTalkChatRoom()
        {
            List<WindowInfo> result = new List<WindowInfo>();

            var list = WindowsMacro.Instance.GetWindowList();
            foreach (var window in list)
            {
                if (WindowsMacro.Instance.IsKakaoTalkChatRoom(window.Handle))
                {
                    result.Add(window);
                }
            }

            return result;
        }

        public void SetForeground(IntPtr hwnd)
        {
            if (!StaInputWorker.Instance.IsCurrent) { StaInputWorker.Instance.Invoke(() => SetForeground(hwnd)); return; }
            SetForegroundWindow(hwnd);
        }
    }
}
