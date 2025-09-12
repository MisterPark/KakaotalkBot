using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
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
        static extern void mouse_event(uint dwFlags, uint dx, uint dy, uint dwData, UIntPtr dwExtraInfo);


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

        

        private static WindowsMacro instance;
        public static WindowsMacro Instance
        {
            get
            {
                if(instance == null)
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
            IntPtr hwndKakao = FindWindow(null, "카카오톡");
            if (hwndKakao == IntPtr.Zero)
            {
                LaunchKakaoTalk();
                return;
            }

            // 2. 검색 Edit 컨트롤 찾아 들어가기
            IntPtr hwndEdit1 = FindWindowEx(hwndKakao, IntPtr.Zero, "EVA_ChildWindow", null);
            IntPtr hwndEdit2_1 = FindWindowEx(hwndEdit1, IntPtr.Zero, "EVA_Window", null);
            IntPtr hwndEdit2_2 = FindWindowEx(hwndEdit1, hwndEdit2_1, "EVA_Window", null);
            IntPtr hwndEdit3 = FindWindowEx(hwndEdit2_2, IntPtr.Zero, "Edit", null); // 최종 Edit 컨트롤

            if (hwndEdit3 == IntPtr.Zero) return;


            SendMessage(hwndEdit3, WM_SETTEXT, 0, " ");
            //Thread.Sleep(300);
            SendReturn(hwndEdit3);
            //Thread.Sleep(300);

            // 3. 검색어 입력
            SendMessage(hwndEdit3, WM_SETTEXT, 0, roomName);
            //Thread.Sleep(500); // 안정성 확보
            // 4. 엔터키 전송 (채팅방 열기)
            SendReturn(hwndEdit3);
            //Thread.Sleep(100);
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
                Form.Invoke((MethodInvoker)delegate { text = Clipboard.GetText(); });

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
                Form.Invoke((MethodInvoker)delegate { Clipboard.SetText(message); });

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
    }
}
