using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
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
        static extern void mouse_event(uint dwFlags, uint dx, uint dy, uint dwData, UIntPtr dwExtraInfo);

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


        private Thread thread;
        private bool isRunning = false;

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


        private Settings settings;
        private Database db;

        private Queue<Command> commands = new Queue<Command>();

        private List<string> chatLog = new List<string>();
        string lastChat = "xx";
        private Random random = new Random();

        private DateTime lastUpdate = DateTime.MinValue;

        CustomTimer soliloquyTimer = new CustomTimer(300000);
        CustomTimer newsTimer = new CustomTimer(1800000);

        public Form1()
        {
            InitializeComponent();

            lastUpdate = DateTime.Now;

            settings = Settings.Load();
            Settings.Save(settings);
            textBox2.Text = settings.ApplicationName;
            textBox3.Text = settings.SpreadsheetId;

            db = new Database(settings.ApplicationName, settings.SpreadsheetId);

            yesButton = new Bitmap("수락.bmp");
            yesButton2 = new Bitmap("수락2.bmp");
            noButton = new Bitmap("거절.bmp");
            okButton = new Bitmap("확인.bmp");
            okButton2 = new Bitmap("확인2.png");
            hostButton = new Bitmap("진행자.png");
            headImage = new Bitmap("방장.bmp");
            viceHeadImage = new Bitmap("부방장.bmp");
            listenerImage = new Bitmap("리스너경계선.bmp");

            timer = new System.Windows.Forms.Timer();
            timer.Interval = 500;
            timer.Tick += Timer_Tick;

            timer2 = new System.Windows.Forms.Timer();
            timer2.Interval = 1000;
            timer2.Tick += Timer_Tick2;

            timer3 = new System.Windows.Forms.Timer();
            timer3.Interval = 6000;
            timer3.Tick += Timer_Tick3;
            timer3.Start();

            UpdateWindowList();

            thread = new Thread(WorkerThread);
            thread.Start();
        }

        private void WorkerThread()
        {
            isRunning = true;

            Stopwatch stopwatch = new Stopwatch();
            stopwatch.Start();
            long lastTick = stopwatch.ElapsedMilliseconds;
            long nowTick = stopwatch.ElapsedMilliseconds;
            long deltaTime = nowTick - lastTick;



            while (isRunning)
            {
                nowTick = stopwatch.ElapsedMilliseconds;
                deltaTime = nowTick - lastTick;
                lastTick = nowTick;
                Time.DeltaTime = deltaTime;

                ProcessCopyChat();
                ProcessReset();

                if (soliloquyTimer.Check(deltaTime))
                {
                    ProcessComonBot();
                }

                if (newsTimer.Check(deltaTime))
                {
                    ProcessNews();
                }
            }
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
                //this.Invoke((MethodInvoker)delegate {  });
            }
        }



        private void Timer_Tick(object sender, EventArgs e)
        {

        }

        private void Timer_Tick2(object sender, EventArgs e)
        {
            UpdateWindowList();
            ProcessCopyChat();
        }

        private void Timer_Tick3(object sender, EventArgs e)
        {
            db.UpdateCommands();
            db.UpdateUserTable();
            News.Update();
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

        public static void CloseWindow(IntPtr hwnd)
        {
            if (hwnd != IntPtr.Zero)
            {
                PostMessage(hwnd, WM_CLOSE, 0, 0);
            }
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
            Thread.Sleep(200);

            string text = string.Empty;
            try
            {
                Invoke((MethodInvoker)delegate { text = Clipboard.GetText(); });

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

        public void SendTextToChatroom(string chatroomName, string message)
        {
            soliloquyTimer.Reset();

            IntPtr hwndMain = FindWindow(null, chatroomName);
            IntPtr hwndEdit = FindWindowEx(hwndMain, IntPtr.Zero, "RichEdit50W", null);

            try
            {
                Invoke((MethodInvoker)delegate { Clipboard.SetText(message); });

                SetForegroundWindow(hwndMain);
                Thread.Sleep(100);
                SendCtrlKey(hwndEdit, 'v');

                //SendMessage(hwndEdit, WM_SETTEXT, 0, message);
                //SetRichEditText(hwndEdit, message);
                //Thread.Sleep(300);
                SendReturn(hwndEdit);
                Thread.Sleep(300);
                //SendReturn(hwndEdit);
                //SendReturn(hwndEdit);
            }
            catch (Exception e)
            {
            }

        }
        public void SendReturn(IntPtr hwnd)
        {
            //    PostMessage(hwnd, WM_KEYDOWN, VK_RETURN, 0);
            //    //Thread.Sleep(100);
            //    PostMessage(hwnd, WM_KEYUP, VK_RETURN, 0);
            //Thread.Sleep(100);

            SendKeys.SendWait("~"); // ^ == Ctrl
        }

        private void ProcessCopyChat()
        {
            if (string.IsNullOrEmpty(textBox1.Text)) return;
            ListViewItem item = null;
            this.Invoke((MethodInvoker)delegate { item = listView1.FindItemWithText(textBox1.Text); });
            if (item == null)
            {
                OpenChatRoom(textBox1.Text);
                return;
            }
            IntPtr handle = (IntPtr)item.Tag;

            string chat = CopyChatroomText(handle);
            string[] lines = chat.Split(new[] { "\r\n", "\n" }, StringSplitOptions.None);
            bool isFirstOpen = chatLog.Count == 0;

            chatLog = lines.Where(x => !string.IsNullOrEmpty(x)).ToList();

            if (isFirstOpen && chatLog.Count > 0)
            {
                lastChat = chatLog.Last();
                return;
            }

            int idx = chatLog.Count - 1;
            for (int i = chatLog.Count - 1; i >= 0; i--)
            {
                if (chatLog[i] == lastChat)
                {
                    idx = i;
                    break;
                }
            }

            for (int i = idx + 1; i < chatLog.Count; i++)
            {
                string line = chatLog[i];
                if (line.Contains("님이 들어왔습니다."))
                {
                    int index = line.LastIndexOf("님");
                    string nick = line.Substring(0, index);
                    ProcessKeyword(nick, "/입장");
                    continue;
                }
                else if (line.Contains("님이 나갔습니다."))
                {
                    int index = line.LastIndexOf("님");
                    string nick = line.Substring(0, index);
                    ProcessKeyword(nick, "/퇴장");
                    continue;
                }

                if (!line.StartsWith("[")) continue;

                int firstClose = line.IndexOf(']');
                int secondClose = line.IndexOf(']', firstClose + 1);

                if (firstClose != -1 && secondClose != -1 && secondClose + 2 <= line.Length)
                {
                    string nickname = line.Substring(1, firstClose - 1).Trim();
                    string message = line.Substring(secondClose + 2).Trim();

                    // 키워드 포함 여부 검사
                    if (db.Keywords.Any(k => message.StartsWith(k)))
                    {
                        if (string.IsNullOrEmpty(nickname) == false)
                        {
                            ProcessKeyword(nickname, message);
                        }
                    }
                }
            }

            if (chatLog.Count > 0)
            {
                lastChat = chatLog[chatLog.Count - 1];
            }

            ProcessCommand();

            this.Invoke((MethodInvoker)delegate
            {
                richTextBox1.Text = string.Join("\n", chatLog);

            });
            this.Invoke((MethodInvoker)delegate
            {
                richTextBox1.SelectionStart = richTextBox1.TextLength;

            });
            this.Invoke((MethodInvoker)delegate
            {
                richTextBox1.ScrollToCaret();

            });
        }

        private void ProcessReset()
        {
            if (lastUpdate.Day != DateTime.Now.Day)
            {
                lastUpdate = DateTime.Now;
                db.ResetAttendance();
            }
        }

        private void ProcessComonBot()
        {
            string[] answers = db.GetAnswers("/코몽봇");

            int rand = random.Next(0, answers.Length);
            string answer = answers[rand];

            if (string.IsNullOrEmpty(answer) == false)
            {
                SendTextToChatroom(textBox1.Text, $"{answer}");
            }
        }

        private void ProcessNews()
        {
            SendTextToChatroom(textBox1.Text, $"{News.PoliticsTop6}");
        }

        void RemoveUntilAndIncludingTarget(List<string> list, string target)
        {
            int idx = list.IndexOf(target);
            if (idx >= 0)
            {
                // 0 ~ idx까지 삭제 (타겟 포함)
                list.RemoveRange(0, idx);
            }
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

            //string answer = GetAnswer(command.Keyword);

            //if (string.IsNullOrEmpty(answer) == false)
            //{
            //    SendTextToChatroom(textBox1.Text, $"{answer}");
            //}

            if (command.Keyword == "/?" || command.Keyword == "/명령어" || command.Keyword == "/훈장" || command.Keyword == "/공지사항" || command.Keyword == "/패치노트")
            {
                string answer = db.GetAnswer(command.Keyword);

                if (string.IsNullOrEmpty(answer) == false)
                {
                    SendTextToChatroom(textBox1.Text, $"{answer}");
                }
            }
            else if (command.Keyword == "/입장" || command.Keyword == "/퇴장")
            {
                string answer = db.GetAnswer(command.Keyword);

                if (string.IsNullOrEmpty(answer) == false)
                {
                    SmartString.CurrentNickname = command.Nickname;
                    string parsedAnswer = SmartString.Parse(answer);
                    SendTextToChatroom(textBox1.Text, parsedAnswer);
                }
            }
            else if (command.Keyword == "/출첵")
            {
                string answer = db.GetAnswer(command.Keyword);

                if (string.IsNullOrEmpty(answer) == false)
                {
                    if (db.CheckAttendance(command.Nickname))
                    {
                        SendTextToChatroom(textBox1.Text, $"이미 출석한 유저입니다.");
                    }
                    else
                    {
                        if (db.FindUser(command.Nickname, out User user))
                        {
                            SendTextToChatroom(textBox1.Text, $"[{command.Nickname}]님이 {answer}\n+10포인트\n(현재 포인트: {user.Point})");
                        }
                        else
                        {
                            SendTextToChatroom(textBox1.Text, $"[{command.Nickname}]님이 {answer}\n+10포인트");
                        }


                    }
                }
            }
            else if (command.Keyword.StartsWith("/조회"))
            {
                if (command.Keyword == "/조회")
                {
                    if (db.FindUser(command.Nickname, out User user))
                    {
                        SendTextToChatroom(textBox1.Text, $"=====[유저조회]=====\n닉네임: {user.Nickname}\n포인트: {user.Point}\n인기도: {user.Popularity}\n=================");
                    }
                    else
                    {
                        SendTextToChatroom(textBox1.Text, $"정보가 없는 유저입니다.");
                    }
                }
                else if (command.Keyword.Length > 4)
                {
                    string param = command.Keyword.Substring(4);
                    param = param.Replace("@", "");
                    if (db.FindUser(param, out User user))
                    {
                        SendTextToChatroom(textBox1.Text, $"=====[유저조회]=====\n닉네임: {user.Nickname}\n포인트: {user.Point}\n인기도: {user.Popularity}\n=================");
                    }
                    else
                    {
                        SendTextToChatroom(textBox1.Text, $"정보가 없는 유저입니다.");
                    }
                }
            }
            else if (command.Keyword.StartsWith("/랭킹"))
            {
                if (command.Keyword == "/랭킹")
                {
                    string answer = db.GetAnswer(command.Keyword);

                    if (string.IsNullOrEmpty(answer)) return;

                    StringBuilder sb = new StringBuilder();
                    int beforeRank = 1;
                    int beforePop = 0;
                    List<User> rank = db.GetPopularityRank();

                    sb.AppendLine(answer);
                    for (int i = 0; i < rank.Count; i++)
                    {
                        int currentPop = rank[i].Popularity;
                        if (currentPop != beforePop)
                        {
                            beforeRank = i + 1;
                        }

                        string emoji = string.Empty;
                        if (beforeRank == 1)
                        {
                            emoji = "🥇";
                        }
                        else if (beforeRank == 2)
                        {
                            emoji = "🥈";
                        }
                        else if (beforeRank == 3)
                        {
                            emoji = "🥉";
                        }

                        sb.AppendLine($"{emoji}{beforeRank}위 {rank[i].Nickname} {rank[i].Popularity}");

                        beforePop = currentPop;
                    }

                    SendTextToChatroom(textBox1.Text, sb.ToString());
                }
            }
            else if (command.Keyword.StartsWith("/좋아"))
            {
                var match = Regex.Match(command.Keyword, @"^\/(\S+)\s+(.+?)\s+(\d+)$");

                if (match.Success)
                {
                    int point = 10;
                    string keyword = match.Groups[1].Value;
                    string nickname = match.Groups[2].Value;
                    string number = match.Groups[3].Value;

                    if(keyword != "/좋아")
                    {
                        return;
                    }

                    if (command.Nickname == nickname)
                    {
                        SendTextToChatroom(textBox1.Text, $"자신에게 할 수 없는 명령입니다.");
                        return;
                    }

                    if (db.FindUser(command.Nickname, out User a))
                    {
                        if (db.FindUser(nickname, out User b))
                        {
                            if (a.Point < point)
                            {
                                SendTextToChatroom(textBox1.Text, $"포인트가 부족합니다.\n남은 포인트: {a.Point}");
                            }
                            else
                            {
                                a.Point -= point;
                                b.Popularity += point;
                                SendTextToChatroom(textBox1.Text, $"[{b.Nickname}]님에게 👍좋아요.\n인기도 {point}점 상승");
                            }
                        }
                        else
                        {
                            SendTextToChatroom(textBox1.Text, $"정보가 없는 유저입니다.");
                        }
                    }
                    else
                    {
                        SendTextToChatroom(textBox1.Text, $"정보가 없는 유저입니다.");
                    }

                }
                else
                {
                    SendTextToChatroom(textBox1.Text, $"잘못된 형식입니다.\n형식1: /좋아 [닉네임]\n형식2: /좋아 [닉네임] [숫자]");
                }
                
            }
            else if (command.Keyword.StartsWith("/싫어"))
            {
                var match = Regex.Match(command.Keyword, @"^\/(\S+)\s+(.+?)\s+(\d+)$");

                if (match.Success)
                {
                    int point = 10;
                    string keyword = match.Groups[1].Value;
                    string nickname = match.Groups[2].Value;
                    string number = match.Groups[3].Value;

                    if (keyword != "/싫어")
                    {
                        return;
                    }

                    if (command.Nickname == nickname)
                    {
                        SendTextToChatroom(textBox1.Text, $"자신에게 할 수 없는 명령입니다.");
                        return;
                    }

                    if (db.FindUser(command.Nickname, out User a))
                    {
                        if (db.FindUser(nickname, out User b))
                        {
                            if (a.Point < point)
                            {
                                SendTextToChatroom(textBox1.Text, $"포인트가 부족합니다.\n남은 포인트: {a.Point}");
                            }
                            else
                            {
                                a.Point -= point;
                                b.Popularity -= point;
                                SendTextToChatroom(textBox1.Text, $"[{b.Nickname}]님에게 👎싫어요.\n인기도 {point}점 하락");
                            }
                        }
                        else
                        {
                            SendTextToChatroom(textBox1.Text, $"정보가 없는 유저입니다.");
                        }
                    }
                    else
                    {
                        SendTextToChatroom(textBox1.Text, $"정보가 없는 유저입니다.");
                    }

                }
                else
                {
                    SendTextToChatroom(textBox1.Text, $"잘못된 형식입니다.\n형식1: /싫어 [닉네임]\n형식2: /싫어 [닉네임] [숫자]");
                }
            }
            else if (command.Keyword.StartsWith("/정치뉴스"))
            {
                SendTextToChatroom(textBox1.Text, $"{News.PoliticsTop6}");
            }
            else
            {
                string[] answers = db.GetAnswers(command.Keyword);
                if (answers == null || answers.Length == 0)
                {
                    return;
                }

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

            return handle;
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
                button2.Text = "시작";
            }
            else
            {
                timer.Start();
                button1.BackColor = Color.Green;
                button2.Text = "실행중...";
            }

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            RegisterHotKey(this.Handle, 1, 0, (int)Keys.F5);
            RegisterHotKey(this.Handle, 2, 0, (int)Keys.F6);
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            timer.Stop();
            timer2.Stop();
            timer3.Stop();

            isRunning = false;
            thread.Join();

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
                button2.Text = "시작";
            }
            else
            {
                timer2.Start();
                button2.BackColor = Color.Green;
                button2.Text = "실행중...";
            }
        }

        private void listView1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (listView1.SelectedItems.Count == 0) return;

            string roomName = listView1.SelectedItems[0].SubItems[0].Text;
            textBox1.Text = roomName;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            //SendTextToChatroom(textBox1.Text, $"앙 기모띠");
            UpdateWindowList();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                string fullPath = openFileDialog1.FileName;
                textBox4.Text = fullPath;
                string path = Path.GetDirectoryName(fullPath);
                string fileName = Path.GetFileNameWithoutExtension(fullPath) + "_decrypted.db";


                textBox5.Text = Path.Combine(path, fileName);
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            SendTextToChatroom(textBox1.Text, $"{News.PoliticsTop6}");
        }
    }
}
