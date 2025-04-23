using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Forms;

namespace KakaotalkBot
{
    public partial class Form1 : Form
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
        static extern bool SetCursorPos(int x, int y);
        [DllImport("user32.dll")]
        static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);
        [DllImport("user32.dll")]
        static extern void mouse_event(uint dwFlags, uint dx, uint dy, uint dwData, UIntPtr dwExtraInfo);
        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern IntPtr GlobalAlloc(uint uFlags, UIntPtr dwBytes);
        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern IntPtr GlobalLock(IntPtr hMem);
        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool GlobalUnlock(IntPtr hMem);
        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern IntPtr GlobalFree(IntPtr hMem);

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

        public const int WM_CLOSE = 0x0010;
        public const int WM_SETTEXT = 0x000C;
        public const int WM_KEYDOWN = 0x0100;
        public const int WM_KEYUP = 0x0101;
        public const int VK_RETURN = 0x0D;

        private const int WM_USER = 0x0400;
        private const int EM_SETTEXTEX = WM_USER + 97;

        private const int ST_DEFAULT = 0x0000;
        private const int ST_KEEPUNDO = 0x0001;

        [DllImport("user32.dll")]
        private static extern bool RegisterHotKey(IntPtr hWnd, int id, int fsModifiers, int vk);

        [DllImport("user32.dll")]
        private static extern bool UnregisterHotKey(IntPtr hWnd, int id);

        [DllImport("user32.dll")]
        private static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        [DllImport("user32.dll")]
        private static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter, string lpszClass, string lpszWindow);
        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);
        [DllImport("user32.dll")]
        public static extern bool PostMessage(IntPtr hWnd, uint Msg, int wParam, int lParam);

        [DllImport("user32.dll")]
        public static extern bool SendMessage(IntPtr hWnd, uint Msg, int wParam, string lParam);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern IntPtr SendMessage(IntPtr hWnd, int Msg, int wParam, int lParam);
        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);

        [StructLayout(LayoutKind.Sequential)]
        private struct SETTEXTEX
        {
            public uint flags;
            public uint codepage;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct RECT
        {
            public int Left;   // 창의 왼쪽 위치 (X)
            public int Top;    // 창의 위쪽 위치 (Y)
            public int Right;  // 오른쪽 (X + Width)
            public int Bottom; // 아래쪽 (Y + Height)
        }

        public struct WindowInfo
        {
            public string Title;
            public IntPtr Handle;
        }

        public struct Command
        {
            public string Nickname;
            public string Keyword;
        }

        private System.Windows.Forms.Timer timer;
        private System.Windows.Forms.Timer timer2;
        private System.Windows.Forms.Timer timer3;
        private Bitmap yesButton;
        private Bitmap yesButton2;
        private Bitmap noButton;
        private Bitmap okButton;
        private Bitmap okButton2;
        private Bitmap hostButton;
        private Bitmap headImage;
        private Bitmap viceHeadImage;
        private Bitmap listenerImage;

        private GoogleSheetHelper keywordSheet;
        private Settings settings;

        private List<string> keywords = new List<string>();
        private List<List<string>> keywordAnswer = new List<List<string>>();

        private Queue<Command> commands = new Queue<Command>();

        private List<string> chatLog = new List<string>();
        private int chatIndex = 0;
        private Random random = new Random();
        public Form1()
        {
            InitializeComponent();

            settings = Settings.Load();
            Settings.Save(settings);
            textBox2.Text = settings.ApplicationName;
            textBox3.Text = settings.SpreadsheetId;

            keywordSheet = new GoogleSheetHelper(settings.ApplicationName, settings.SpreadsheetId);
            keywords = GetKeywords();
            keywordAnswer = GetKeywordList();

            yesButton = new Bitmap("수락.png");
            yesButton2 = new Bitmap("수락2.png");
            noButton = new Bitmap("거절.png");
            okButton = new Bitmap("확인.png");
            okButton2 = new Bitmap("확인2.png");
            hostButton = new Bitmap("진행자.png");
            headImage = new Bitmap("방장.png");
            viceHeadImage = new Bitmap("부방장.png");
            listenerImage = new Bitmap("리스너경계선.png");

            timer = new System.Windows.Forms.Timer();
            timer.Interval = 500;
            timer.Tick += Timer_Tick;

            timer2 = new System.Windows.Forms.Timer();
            timer2.Interval = 3000;
            timer2.Tick += Timer_Tick2;

            timer3 = new System.Windows.Forms.Timer();
            timer3.Interval = 6000;
            timer3.Tick += Timer_Tick3;
            timer3.Start();

            UpdateWindowList();
        }

        private List<WindowInfo> GetWindowList()
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

        private void UpdateWindowList()
        {
            listView1.Items.Clear();

            List<WindowInfo> windowList = GetWindowList();
            foreach (WindowInfo window in windowList)
            {
                string[] strings = { window.Title, window.Handle.ToString() };
                ListViewItem item = new ListViewItem(strings);
                item.Tag = window.Handle;
                listView1.Items.Add(item);
            }
        }

        private List<List<string>> GetDB()
        {
            var db = keywordSheet.ReadAllFromSheet("DB");
            return db;
        }

        private List<List<string>> GetKeywordList()
        {
            var keywords = keywordSheet.ReadAllFromSheet("키워드");
            return keywords;
        }

        private List<string> GetKeywords()
        {
            List<string> keywordList = new List<string>();
            var keywords = keywordSheet.ReadAllFromSheet("키워드");
            foreach (var row in keywords)
            {
                keywordList.Add(row[0]);
            }

            return keywordList;
        }

        private string GetAnswer(string keyword)
        {
            var keywords = keywordAnswer;
            foreach (var row in keywords)
            {
                if (row[0] == keyword)
                {
                    return row[1];
                }
            }

            return null;
        }

        private string[] GetAnswers(string keyword)
        {
            var keywords = keywordAnswer;
            foreach (var row in keywords)
            {
                if (row[0] == keyword)
                {
                    return row[1].Split(new string[] { "\n" }, StringSplitOptions.RemoveEmptyEntries);
                }
            }

            return null;
        }

        private void Timer_Tick(object sender, EventArgs e)
        {
            DetectScreen();
            DetectScreen2();
        }

        private void Timer_Tick2(object sender, EventArgs e)
        {
            UpdateWindowList();
            ProcessCopyChat();
        }

        private void Timer_Tick3(object sender, EventArgs e)
        {
            keywordSheet = new GoogleSheetHelper(settings.ApplicationName, settings.SpreadsheetId);
            keywords = GetKeywords();
            keywordAnswer = GetKeywordList();
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

        private void OpenChatRoom(string roomName)
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
            Thread.Sleep(300);
            SendReturn(hwndEdit3);
            Thread.Sleep(300);

            // 3. 검색어 입력
            SendMessage(hwndEdit3, WM_SETTEXT, 0, roomName);
            Thread.Sleep(500); // 안정성 확보
            // 4. 엔터키 전송 (채팅방 열기)
            SendReturn(hwndEdit3);
            Thread.Sleep(100);
        }

        public static void CloseWindow(IntPtr hwnd)
        {
            if (hwnd != IntPtr.Zero)
            {
                PostMessage(hwnd, WM_CLOSE, 0, 0);
            }
        }
        private string CopyChatroomText(string windowName)
        {
            IntPtr hwndMain = FindWindow(null, windowName);
            IntPtr hwndList = FindWindowEx(hwndMain, IntPtr.Zero, "EVA_VH_ListControl_Dblclk", null);

            if (hwndList == IntPtr.Zero) return "";

            // 채팅 전체 선택 후 복사 (Ctrl+A → Ctrl+C)
            SendCtrlKey(hwndList, 'A');
            Thread.Sleep(100);
            SendCtrlKey(hwndList, 'C');
            Thread.Sleep(100);

            try
            {
                IDataObject iData = Clipboard.GetDataObject();
                if (iData.GetDataPresent(DataFormats.Text))
                {
                    return (string)iData.GetData(DataFormats.Text);
                }
            }
            catch (Exception ex)
            {
                return $"클립보드 오류: {ex.Message}";
            }

            return "";
        }
        private string CopyChatroomText(IntPtr hwndMain)
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
            Thread.Sleep(300);

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
        private void SendCtrlKey(IntPtr hwnd, char key)
        {
            SetForegroundWindow(hwnd);
            SendKeys.SendWait("^" + key); // ^ == Ctrl
        }

        private void SetRichEditText(IntPtr richEditHandle, string text)
        {
            // SETTEXTEX 구조체 생성
            SETTEXTEX setTextEx = new SETTEXTEX
            {
                flags = ST_DEFAULT,
                codepage = 1200 // UTF-16 (Unicode)
            };

            IntPtr ptrStruct = Marshal.AllocHGlobal(Marshal.SizeOf(setTextEx));
            Marshal.StructureToPtr(setTextEx, ptrStruct, false);

            // 문자열 포인터 확보
            IntPtr hGlobal = GlobalAlloc(GMEM_MOVEABLE, (UIntPtr)((text.Length + 1) * 2));
            IntPtr ptrText = GlobalLock(hGlobal);
            Marshal.Copy(text.ToCharArray(), 0, ptrText, text.Length);
            Marshal.WriteInt16(ptrText, text.Length * 2, 0); // null terminator
            GlobalUnlock(hGlobal);

            // 메시지 전송
            SendMessage(richEditHandle, EM_SETTEXTEX, ptrStruct, hGlobal);

            // 메모리 해제
            Marshal.FreeHGlobal(ptrStruct);
            GlobalFree(hGlobal);
        }

        public void SendTextToChatroom(string chatroomName, string message)
        {
            IntPtr hwndMain = FindWindow(null, chatroomName);
            IntPtr hwndEdit = FindWindowEx(hwndMain, IntPtr.Zero, "RichEdit50W", null);

            try
            {
                Clipboard.SetText(message);

                SetForegroundWindow(hwndMain);
                Thread.Sleep(300);
                SendCtrlKey(hwndEdit, 'v');

                //SendMessage(hwndEdit, WM_SETTEXT, 0, message);
                //SetRichEditText(hwndEdit, message);
                Thread.Sleep(300);
                SendReturn(hwndEdit);
            }
            catch (Exception e)
            {
            }

        }
        public void SendReturn(IntPtr hwnd)
        {
            PostMessage(hwnd, WM_KEYDOWN, VK_RETURN, 0);
            Thread.Sleep(100);
            PostMessage(hwnd, WM_KEYUP, VK_RETURN, 0);
            Thread.Sleep(100);
        }

        private void ProcessCopyChat()
        {
            if (string.IsNullOrEmpty(textBox1.Text)) return;

            ListViewItem item = listView1.FindItemWithText(textBox1.Text);
            if (item == null)
            {
                OpenChatRoom(textBox1.Text);
                return;
            }
            IntPtr handle = (IntPtr)item.Tag;

            string chat = CopyChatroomText(handle);
            string[] lines = chat.Split(new[] { "\r\n", "\n" }, StringSplitOptions.None);

            for (int i = 0; i < lines.Length; i++)
            {
                string line = lines[i];

                if (chatLog.Contains(line) == false)
                {
                    chatLog.Add(line);
                }
            }

            for (int i = chatIndex; i < chatLog.Count; i++)
            {
                string line = chatLog[i];
                if (!line.StartsWith("[")) continue;

                int firstClose = line.IndexOf(']');
                int secondClose = line.IndexOf(']', firstClose + 1);

                if (firstClose != -1 && secondClose != -1 && secondClose + 2 <= line.Length)
                {
                    string nickname = line.Substring(1, firstClose - 1).Trim();
                    string message = line.Substring(secondClose + 2).Trim();

                    // 키워드 포함 여부 검사
                    if (keywords.Any(k => message == k))
                    {
                        if (string.IsNullOrEmpty(nickname) == false)
                        {
                            ProcessKeyword(nickname, message);
                        }
                    }
                }
                chatIndex++;
            }

            ProcessCommand();

            richTextBox1.Text = string.Join("\n", chatLog);
            richTextBox1.SelectionStart = richTextBox1.TextLength;
            richTextBox1.ScrollToCaret();
        }

        private void ProcessKeyword(string nickname, string message)
        {
            //SendTextToChatroom(textBox1.Text, message);

            Command command = new Command();
            command.Nickname = nickname;
            command.Keyword = message;
            commands.Enqueue(command);
        }

        private void ProcessCommand()
        {
            if (commands.Count == 0) return;


            Command command = commands.Dequeue();

            if (command.Keyword == "/명령어")
            {
                string answer = GetAnswer(command.Keyword);

                if (string.IsNullOrEmpty(answer) == false)
                {
                    SendTextToChatroom(textBox1.Text, $"[{command.Nickname}]님의 명령\n{answer}");
                }
            }
            else if (command.Keyword == "/출첵")
            {
                string answer = GetAnswer(command.Keyword);

                if (string.IsNullOrEmpty(answer) == false)
                {
                    SendTextToChatroom(textBox1.Text, $"[{command.Nickname}]님의 명령\n{answer}");
                }
            }
            else
            {
                string[] answers = GetAnswers(command.Keyword);

                int rand = random.Next(0, answers.Length);
                string answer = answers[rand];

                if (string.IsNullOrEmpty(answer) == false)
                {
                    SendTextToChatroom(textBox1.Text, $"{answer}");
                }
            }

        }

        //================================================


        private IntPtr FindVoiceRoomWindow()
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

            //richTextBox1.Text = sb.ToString();

            return handle;
        }

        private void DetectScreen()
        {
            IntPtr handle = FindVoiceRoomWindow();
            if (handle != IntPtr.Zero)
            {
                if (GetWindowRect(handle, out RECT rect))
                {
                    int x = rect.Left;
                    int y = rect.Top;
                    int width = rect.Right - rect.Left;
                    int height = rect.Bottom - rect.Top;

                    //SetCursorPos(x, y);
                    Rectangle voiceRoomArea = new Rectangle(x, y, width, height);

                    Rectangle totalBounds = Screen.AllScreens.Select(s => s.Bounds).Aggregate(Rectangle.Union);
                    using (Bitmap bitmap = new Bitmap(totalBounds.Width, totalBounds.Height))
                    {
                        using (Graphics g = Graphics.FromImage(bitmap))
                        {
                            g.CopyFromScreen(totalBounds.Location, Point.Empty, totalBounds.Size);
                        }

                        // 이미지 자르기
                        Bitmap voiceRoom = bitmap.Clone(voiceRoomArea, bitmap.PixelFormat);

                        // 리스너 경계선 찾기
                        Point listnerPoint = new Point(0, 0);
                        if (FindImage(voiceRoom, listenerImage, ref listnerPoint, 0.05))
                        {
                            Rectangle speakerArea = new Rectangle(x, y, width, listnerPoint.Y);
                            Bitmap speaker = bitmap.Clone(speakerArea, bitmap.PixelFormat);

                            Rectangle listnerArea = new Rectangle(x, listnerPoint.Y, width, height - listnerPoint.Y);
                            Bitmap listner = bitmap.Clone(listnerArea, bitmap.PixelFormat);

                            // 자동수락
                            Point targetPoint = new Point(0, 0);
                            if (FindImage(voiceRoom, yesButton, ref targetPoint, 0.05))
                            {
                                SetCursorPos(x + targetPoint.X, y + targetPoint.Y);
                                ClickLeft();
                            }

                            // 자동수락2
                            Point targetPoint0 = new Point(0, 0);
                            if (FindImage(voiceRoom, yesButton2, ref targetPoint0, 0.05))
                            {
                                SetCursorPos(x + targetPoint0.X, y + targetPoint0.Y);
                                ClickLeft();
                            }

                            // 자동확인
                            Point targetPoint2 = new Point(0, 0);
                            if (FindImage(voiceRoom, okButton, ref targetPoint2, 0.05))
                            {
                                SetCursorPos(x + targetPoint2.X, y + targetPoint2.Y);
                                ClickLeft();
                            }

                            // 자동확인2
                            Point targetPoint3 = new Point(0, 0);
                            if (FindImage(voiceRoom, okButton2, ref targetPoint3, 0.05))
                            {
                                SetCursorPos(x + targetPoint3.X, y + targetPoint3.Y);
                                ClickLeft();
                            }

                            // 부방장 자동진행자
                            Point targetPoint4 = new Point(0, 0);
                            if (FindImage(speaker, viceHeadImage, ref targetPoint4, 0.05))
                            {
                                SetCursorPos(x + targetPoint4.X, y + targetPoint4.Y);
                                ClickRight();
                            }

                            // 방장 자동진행자
                            Point targetPoint5 = new Point(0, 0);
                            if (FindImage(speaker, headImage, ref targetPoint5, 0.05))
                            {
                                SetCursorPos(x + targetPoint5.X, y + targetPoint5.Y);
                                ClickRight();
                            }

                            // 자동 진행자로 초대
                            Point targetPoint6 = new Point(0, 0);
                            if (FindImage(voiceRoom, hostButton, ref targetPoint6, 0.05))
                            {
                                SetCursorPos(x + targetPoint6.X, y + targetPoint6.Y);
                                ClickLeft();
                            }

                            speaker.Dispose();
                            listner.Dispose();
                        }

                        voiceRoom.Dispose();
                    }
                }
            }
        }
        private void DetectScreen2()
        {
            IntPtr handle = FindVoiceRoomWindow();
            if (handle != IntPtr.Zero)
            {
                if (GetWindowRect(handle, out RECT rect))
                {
                    int x = rect.Left;
                    int y = rect.Top;
                    int width = rect.Right - rect.Left;
                    int height = rect.Bottom - rect.Top;

                    //SetCursorPos(x, y);
                    Rectangle voiceRoomArea = new Rectangle(x, y, width, height);

                    Rectangle totalBounds = Screen.AllScreens.Select(s => s.Bounds).Aggregate(Rectangle.Union);
                    using (Bitmap bitmap = new Bitmap(totalBounds.Width, totalBounds.Height))
                    {
                        using (Graphics g = Graphics.FromImage(bitmap))
                        {
                            g.CopyFromScreen(totalBounds.Location, Point.Empty, totalBounds.Size);
                        }

                        // 이미지 자르기
                        Bitmap voiceRoom = bitmap.Clone(voiceRoomArea, bitmap.PixelFormat);

                        // 자동 진행자로 초대
                        Point targetPoint6 = new Point(0, 0);
                        if (FindImage(voiceRoom, hostButton, ref targetPoint6, 0.05))
                        {
                            SetCursorPos(x + targetPoint6.X, y + targetPoint6.Y);
                            ClickLeft();
                        }

                        voiceRoom.Dispose();
                    }
                }
            }
        }

        private byte[] GetPixelData(Bitmap bmp, out int stride)
        {
            Rectangle rect = new Rectangle(0, 0, bmp.Width, bmp.Height);
            BitmapData bmpData = bmp.LockBits(rect, ImageLockMode.ReadOnly, PixelFormat.Format24bppRgb);
            IntPtr ptr = bmpData.Scan0;
            stride = bmpData.Stride;
            int bytes = Math.Abs(stride) * bmp.Height;
            byte[] rgbValues = new byte[bytes];

            Marshal.Copy(ptr, rgbValues, 0, bytes);
            bmp.UnlockBits(bmpData);

            return rgbValues;
        }

        public bool FindImage(Bitmap source, Bitmap template, ref Point point, double error = 0.01)
        {
            // 이미지 LockBits
            BitmapData sourceData = source.LockBits(
                new Rectangle(0, 0, source.Width, source.Height),
                ImageLockMode.ReadOnly,
                PixelFormat.Format24bppRgb);

            BitmapData templateData = template.LockBits(
                new Rectangle(0, 0, template.Width, template.Height),
                ImageLockMode.ReadOnly,
                PixelFormat.Format24bppRgb);

            int srcStride = sourceData.Stride;
            int tplStride = templateData.Stride;

            int tplW = template.Width;
            int tplH = template.Height;

            byte[] srcBytes = new byte[srcStride * source.Height];
            byte[] tplBytes = new byte[tplStride * template.Height];

            Marshal.Copy(sourceData.Scan0, srcBytes, 0, srcBytes.Length);
            Marshal.Copy(templateData.Scan0, tplBytes, 0, tplBytes.Length);

            source.UnlockBits(sourceData);
            template.UnlockBits(templateData);


            // 슬라이딩 윈도우 매칭
            for (int y = 0; y <= source.Height - tplH; y++)
            {
                for (int x = 0; x <= source.Width - tplW; x++)
                {
                    if (MatchImageInner(srcBytes, tplBytes, tplW, tplH, srcStride, tplStride, x, y, error))
                    {
                        point = new Point(x, y);
                        return true;
                    }

                }
            }

            return false;
        }

        private bool MatchImageInner(byte[] srcBytes, byte[] tplBytes, int tplW, int tplH, int srcStride, int tplStride, int srcX, int srcY, double error = 0.01)
        {
            int matchPixelCount = 0;
            int totalPixelCount = tplW * tplH;

            for (int j = 0; j < tplH; j++)
            {
                for (int i = 0; i < tplW; i++)
                {
                    int srcIndex = (srcY + j) * srcStride + (srcX + i) * 3;
                    int tplIndex = j * tplStride + i * 3;

                    // BGR 순서 비교
                    int sb = srcBytes[srcIndex];
                    int sg = srcBytes[srcIndex + 1];
                    int sr = srcBytes[srcIndex + 2];

                    int tb = tplBytes[tplIndex];
                    int tg = tplBytes[tplIndex + 1];
                    int tr = tplBytes[tplIndex + 2];

                    byte sbRate = (byte)(sb * error);
                    byte sgRate = (byte)(sg * error);
                    byte srRate = (byte)(sr * error);

                    int maxB = sb + sbRate;
                    int maxG = sg + sgRate;
                    int maxR = sr + srRate;

                    int minB = sb - sbRate;
                    int minG = sg - sgRate;
                    int minR = sr - srRate;

                    // 여기 수정해야함.

                    if (sb != tb) return false;
                    if (sg != tg) return false;
                    if (sr != tr) return false;

                    //if (tb < minB) return false;
                    //if (tb > maxB) return false;
                    //if (tg < minG) return false;
                    //if (tg > maxG) return false;
                    //if (tr < minR) return false;
                    //if (tr > maxR) return false;


                    //if (tb >= minB && tb <= maxB && tg >= minG && tg <= maxG && tr >= minR && tr <= maxR)
                    //{
                    //    matchPixelCount++;
                    //}
                }
            }

            //double threshold = 1 - error;
            //double similarity = (double)matchPixelCount / totalPixelCount;
            //if (similarity >= threshold)
            //{
            //    return true;
            //}

            return true;
        }


        private void ClickLeft()
        {
            mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, UIntPtr.Zero);
            mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, UIntPtr.Zero);
        }

        private void ClickRight()
        {
            mouse_event(MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, UIntPtr.Zero);
            mouse_event(MOUSEEVENTF_RIGHTUP, 0, 0, 0, UIntPtr.Zero);
        }

        protected override void WndProc(ref Message m)
        {
            if (m.Msg == WM_HOTKEY)
            {
                int id = m.WParam.ToInt32();
                if (id == 1)
                {
                    // F5 누르면 실행할 동작
                    if (timer.Enabled)
                    {
                        timer.Stop();
                        button1.BackColor = Color.Red;
                    }
                    else
                    {
                        timer.Start();
                        button1.BackColor = Color.Green;
                    }
                }

                if (id == 2)
                {
                    if (timer2.Enabled)
                    {
                        timer2.Stop();
                        button2.BackColor = Color.Red;
                    }
                    else
                    {
                        timer2.Start();
                        button2.BackColor = Color.Green;
                    }
                }
            }
            base.WndProc(ref m);
        }

        private void button1_Click(object sender, System.EventArgs e)
        {
            if (timer.Enabled)
            {
                timer.Stop();
                button1.BackColor = Color.Red;
            }
            else
            {
                timer.Start();
                button1.BackColor = Color.Green;
            }

        }

        private void button3_Click(object sender, EventArgs e)
        {
            //pictureBox1.Image.Save("Screenshot.png");
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            RegisterHotKey(this.Handle, 1, 0, (int)Keys.F5);
            RegisterHotKey(this.Handle, 2, 0, (int)Keys.F6);
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            UnregisterHotKey(this.Handle, 2);
            UnregisterHotKey(this.Handle, 1);
            Settings.Save(settings);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (timer2.Enabled)
            {
                timer2.Stop();
                button2.BackColor = Color.Red;
            }
            else
            {
                timer2.Start();
                button2.BackColor = Color.Green;
            }
        }

        private void listView1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (listView1.SelectedItems.Count == 0) return;

            string roomName = listView1.SelectedItems[0].SubItems[0].Text;
            textBox1.Text = roomName;
        }

        private void button3_Click_1(object sender, EventArgs e)
        {
            SendTextToChatroom(textBox1.Text, $"앙 기모띠");
        }
    }
}
