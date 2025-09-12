using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Forms;

namespace KakaotalkBot
{
    public partial class Form1 : Form
    {

        [DllImport("user32.dll")]
        private static extern bool RegisterHotKey(IntPtr hWnd, int id, int fsModifiers, int vk);

        [DllImport("user32.dll")]
        private static extern bool UnregisterHotKey(IntPtr hWnd, int id);

        private const int WM_HOTKEY = 0x0312;


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



        public struct QuizAnswer
        {
            public string Nickname;
            public string Answer;
        }

        private Thread thread;
        private bool isThreadRunning = false;
        private bool isBotRunning = false;

        private Settings settings;
        private Database db;

        private Queue<Command> commands = new Queue<Command>();

        private List<string> chatLog = new List<string>();
        string lastChat = "xx";
        private Random random;

        private DateTime lastUpdate = DateTime.MinValue;

        CustomTimer soliloquyTimer = new CustomTimer(300000);
        CustomTimer newsTimer = new CustomTimer(3600000);
        CustomTimer dbTimer = new CustomTimer(60000);

        private Queue<QuizAnswer> quizAnswers = new Queue<QuizAnswer>();
        private bool isCorrect = false;

        private System.Windows.Forms.Timer timer;

        public Form1()
        {

            InitializeComponent();

            timer = new System.Windows.Forms.Timer();
            timer.Interval = 20;
            timer.Tick += Timer_Tick;
            timer.Start();

            WindowsMacro.Instance.Form = this;

            random = new Random(DateTime.Now.Millisecond);

            lastUpdate = DateTime.Now;

            settings = Settings.Load();
            Settings.Save(settings);
            textBox2.Text = settings.ApplicationName;
            textBox3.Text = settings.SpreadsheetId;

            db = new Database(settings.ApplicationName, settings.SpreadsheetId);


            Application.ApplicationExit += new EventHandler(OnApplicationExit);

            UpdateWindowList();

            thread = new Thread(WorkerThread);
            thread.Start();
        }

        private void Timer_Tick(object sender, EventArgs e)
        {
            if(isBotRunning)
            {
                button2.BackColor = Color.Green;
            }
            else
            {
                button2.BackColor = Color.Red;
            }
        }

        private void OnApplicationExit(object sender, EventArgs e)
        {

        }

        private void WorkerThread()
        {
            isThreadRunning = true;

            Stopwatch stopwatch = new Stopwatch();
            stopwatch.Start();
            long lastTick = stopwatch.ElapsedMilliseconds;
            long nowTick = stopwatch.ElapsedMilliseconds;
            long deltaTime = nowTick - lastTick;

            while (isThreadRunning)
            {
                nowTick = stopwatch.ElapsedMilliseconds;
                deltaTime = nowTick - lastTick;
                lastTick = nowTick;
                Time.DeltaTime = deltaTime;

                if(isBotRunning)
                {
                    ProcessCopyChat();
                    ProcessReset();

                    if (soliloquyTimer.Check(deltaTime))
                    {
                        //ProcessComonBot();

                        ProcessNextCommonSense();
                    }

                    if (newsTimer.Check(deltaTime))
                    {
                        ProcessNews();
                    }

                    if (dbTimer.Check(deltaTime))
                    {
                        ProcessUpdateDB();
                    }

                    ProcessQuiz();
                }
                
            }
        }

        private void UpdateWindowList()
        {
            listView1.Items.Clear();

            List<WindowInfo> windowList = WindowsMacro.Instance.GetWindowList();
            foreach (WindowInfo window in windowList)
            {
                string[] strings = { window.Title, window.Handle.ToString() };
                ListViewItem item = new ListViewItem(strings);
                item.Tag = window.Handle;
                listView1.Items.Add(item);
            }
        }


        private void ProcessCopyChat()
        {
            if (string.IsNullOrEmpty(textBox1.Text)) return;
            ListViewItem item = null;
            this.Invoke((MethodInvoker)delegate { item = listView1.FindItemWithText(textBox1.Text); });
            if (item == null)
            {
                WindowsMacro.Instance.OpenChatRoom(textBox1.Text);
                return;
            }
            IntPtr handle = (IntPtr)item.Tag;

            string chat = WindowsMacro.Instance.CopyChatroomText(handle);
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

                    ProcessQuizAnswer(nickname, message);
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

        private void ProcessNews()
        {
            WindowsMacro.Instance.SendTextToChatroom(textBox1.Text, $"{News.PoliticsTop6}");
        }

        private void ProcessUpdateDB()
        {
            db.UpdateCommands();
            db.UpdateUserTable();
            db.UpdateCommonSenses();
            db.UpdateTopic();
            News.Update();
        }

        private void ProcessCommonSense()
        {
            string answer = db.GetCommonSenseText();
            if (string.IsNullOrEmpty(answer))
            {
                double left = soliloquyTimer.TimeLeft * 0.001;

                WindowsMacro.Instance.SendTextToChatroom(textBox1.Text, $"다음 퀴즈를 준비하고 있습니다.\n남은 시간: {left}초");
            }
            else
            {
                WindowsMacro.Instance.SendTextToChatroom(textBox1.Text, $"{answer}");
            }
        }

        private void ProcessNextCommonSense()
        {
            Quiz quiz = db.GetCurrentQuiz();
            if (quiz != null)
            {
                if (isCorrect == false)
                {
                    WindowsMacro.Instance.SendTextToChatroom(textBox1.Text, $"정답자가 없습니다.\n정답: {quiz.Answer}\n해설: {quiz.Explanation}");
                }
            }

            db.SetNextCommonSense();
            ProcessCommonSense();
        }

        private void ProcessKeyword(string nickname, string message)
        {
            //SendTextToChatroom(textBox1.Text, message);

            Command command = new Command();
            command.Nickname = nickname;
            command.Keyword = message;
            commands.Enqueue(command);
        }

        private void ProcessQuizAnswer(string nickname, string answer)
        {
            QuizAnswer quizAnswer = new QuizAnswer();
            quizAnswer.Nickname = nickname;
            quizAnswer.Answer = answer;
            quizAnswers.Enqueue(quizAnswer);
        }

        private void ProcessQuiz()
        {
            Quiz quiz = db.GetCurrentQuiz();
            if (quiz == null)
            {
                return;
            }

            while (quizAnswers.Count != 0)
            {
                QuizAnswer quizAnswer = quizAnswers.Dequeue();


                if (quizAnswer.Answer == quiz.Answer)
                {
                    if (db.FindUser(quizAnswer.Nickname, out User a))
                    {
                        int point = 0;
                        if (quiz.Difficulty == "최상")
                        {
                            point = 5;
                        }
                        else if (quiz.Difficulty == "상")
                        {
                            point = 4;
                        }
                        else if (quiz.Difficulty == "중")
                        {
                            point = 3;
                        }
                        else if (quiz.Difficulty == "하")
                        {
                            point = 2;
                        }
                        else
                        {
                            point = 1;
                        }

                        a.Point += point;
                        WindowsMacro.Instance.SendTextToChatroom(textBox1.Text, $"정답자: {quizAnswer.Nickname}\n정답: {quiz.Answer}\n해설: {quiz.Explanation}\n+{point} 포인트 득점!!👍\n 현재 포인트: {a.Point}");
                    }
                    else
                    {
                        WindowsMacro.Instance.SendTextToChatroom(textBox1.Text, $"정답: {quiz.Answer}\n해설: {quiz.Explanation}");
                    }

                    db.CurrentAnswerIndex = -1;
                    break;
                }
            }

            quizAnswers.Clear();

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
                    WindowsMacro.Instance.SendTextToChatroom(textBox1.Text, $"{answer}");
                }
            }
            else if (command.Keyword == "/입장" || command.Keyword == "/퇴장")
            {
                string answer = db.GetAnswer(command.Keyword);

                if (string.IsNullOrEmpty(answer) == false)
                {
                    SmartString.CurrentNickname = command.Nickname;
                    string parsedAnswer = SmartString.Parse(answer);
                    WindowsMacro.Instance.SendTextToChatroom(textBox1.Text, parsedAnswer);
                }
            }
            else if (command.Keyword == "/출첵")
            {
                string answer = db.GetAnswer(command.Keyword);

                if (string.IsNullOrEmpty(answer) == false)
                {
                    if (db.CheckAttendance(command.Nickname))
                    {
                        WindowsMacro.Instance.SendTextToChatroom(textBox1.Text, $"이미 출석한 유저입니다.");
                    }
                    else
                    {
                        if (db.FindUser(command.Nickname, out User user))
                        {
                            WindowsMacro.Instance.SendTextToChatroom(textBox1.Text, $"[{command.Nickname}]님이 {answer}\n+10포인트\n(현재 포인트: {user.Point})");
                        }
                        else
                        {
                            WindowsMacro.Instance.SendTextToChatroom(textBox1.Text, $"[{command.Nickname}]님이 {answer}\n+10포인트");
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
                        WindowsMacro.Instance.SendTextToChatroom(textBox1.Text, $"=====[유저조회]=====\n닉네임: {user.Nickname}\n포인트: {user.Point}\n인기도: {user.Popularity}\n=================");
                    }
                    else
                    {
                        WindowsMacro.Instance.SendTextToChatroom(textBox1.Text, $"정보가 없는 유저입니다.");
                    }
                }
                else if (command.Keyword.Length > 4)
                {
                    string param = command.Keyword.Substring(4);
                    param = param.Replace("@", "");
                    if (db.FindUser(param, out User user))
                    {
                        WindowsMacro.Instance.SendTextToChatroom(textBox1.Text, $"=====[유저조회]=====\n닉네임: {user.Nickname}\n포인트: {user.Point}\n인기도: {user.Popularity}\n=================");
                    }
                    else
                    {
                        WindowsMacro.Instance.SendTextToChatroom(textBox1.Text, $"정보가 없는 유저입니다.");
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

                    WindowsMacro.Instance.SendTextToChatroom(textBox1.Text, sb.ToString());
                }
            }
            else if (command.Keyword.StartsWith("/토론"))
            {
                if (command.Keyword == "/토론")
                {
                    string answer = db.GetAnswer(command.Keyword);

                    if (string.IsNullOrEmpty(answer)) return;

                    StringBuilder sb = new StringBuilder();

                    List<Topic> topics = db.GetOrderedTopics();
                    string currentCategory = string.Empty;
                    int count = 0;

                    sb.AppendLine(answer);
                    for (int i = 0; i < topics.Count; i++)
                    {
                        if (currentCategory != topics[i].Category)
                        {
                            sb.AppendLine();
                            sb.AppendLine($"{topics[i].Category}");
                            currentCategory = topics[i].Category;
                            count = 0;
                        }
                        sb.AppendLine($"{count + 1}. {topics[i].Title}({topics[i].CreatedAt})");
                        sb.AppendLine($"(토론왕👑: {topics[i].Winner})");
                        count++;
                    }

                    WindowsMacro.Instance.SendTextToChatroom(textBox1.Text, sb.ToString());
                }
            }
            else if (command.Keyword.StartsWith("/좋아"))
            {
                string[] format = new string[3];
                if (SmartString.ParseCommand(command.Keyword, out format) == ParseResult.Success)
                {
                    string keyword = format[0];
                    string nickname = format[1];
                    string number = format[2];

                    if (nickname.StartsWith("@"))
                    {
                        nickname = nickname.Substring(1);
                    }

                    int point = 10;
                    int.TryParse(number, out point);

                    if (keyword != "/좋아")
                    {
                        return;
                    }

                    if (command.Nickname == nickname)
                    {
                        WindowsMacro.Instance.SendTextToChatroom(textBox1.Text, $"자신에게 할 수 없는 명령입니다.");
                        return;
                    }

                    if (db.FindUser(command.Nickname, out User a))
                    {
                        if (db.FindUser(nickname, out User b))
                        {
                            if (a.Point < Math.Abs(point))
                            {
                                WindowsMacro.Instance.SendTextToChatroom(textBox1.Text, $"포인트가 부족합니다.\n남은 포인트: {a.Point}");
                            }
                            else
                            {
                                a.Point -= Math.Abs(point);
                                b.Popularity += point;
                                WindowsMacro.Instance.SendTextToChatroom(textBox1.Text, $"[{a.Nickname}]님이 [{b.Nickname}]님에게 👍좋아요.\n인기도 {point}점 상승\n현재 인기도:{b.Popularity}");
                            }
                        }
                        else
                        {
                            WindowsMacro.Instance.SendTextToChatroom(textBox1.Text, $"정보가 없는 유저입니다.");
                        }
                    }
                    else
                    {
                        WindowsMacro.Instance.SendTextToChatroom(textBox1.Text, $"정보가 없는 유저입니다.");
                    }
                }
                else
                {
                    WindowsMacro.Instance.SendTextToChatroom(textBox1.Text, $"잘못된 형식입니다.\n형식1: /좋아 [닉네임]\n형식2: /좋아 [닉네임] [숫자]");
                }

            }
            else if (command.Keyword.StartsWith("/싫어"))
            {
                string[] format = new string[3];
                if (SmartString.ParseCommand(command.Keyword, out format) == ParseResult.Success)
                {
                    string keyword = format[0];
                    string nickname = format[1];
                    string number = format[2];

                    if (nickname.StartsWith("@"))
                    {
                        nickname = nickname.Substring(1);
                    }

                    int point = 10;
                    int.TryParse(number, out point);

                    if (keyword != "/싫어")
                    {
                        return;
                    }

                    if (command.Nickname == nickname)
                    {
                        WindowsMacro.Instance.SendTextToChatroom(textBox1.Text, $"자신에게 할 수 없는 명령입니다.");
                        return;
                    }

                    if (db.FindUser(command.Nickname, out User a))
                    {
                        if (db.FindUser(nickname, out User b))
                        {
                            if (a.Point < Math.Abs(point))
                            {
                                WindowsMacro.Instance.SendTextToChatroom(textBox1.Text, $"포인트가 부족합니다.\n남은 포인트: {a.Point}");
                            }
                            else
                            {
                                a.Point -= Math.Abs(point);
                                b.Popularity -= point;
                                WindowsMacro.Instance.SendTextToChatroom(textBox1.Text, $"[{a.Nickname}]님이 [{b.Nickname}]님에게 👎싫어요.\n인기도 {point}점 하락\n현재 인기도:{b.Popularity}");
                            }
                        }
                        else
                        {
                            WindowsMacro.Instance.SendTextToChatroom(textBox1.Text, $"정보가 없는 유저입니다.");
                        }
                    }
                    else
                    {
                        WindowsMacro.Instance.SendTextToChatroom(textBox1.Text, $"정보가 없는 유저입니다.");
                    }
                }
                else
                {
                    WindowsMacro.Instance.SendTextToChatroom(textBox1.Text, $"잘못된 형식입니다.\n형식1: /싫어 [닉네임]\n형식2: /싫어 [닉네임] [숫자]");
                }
            }
            else if (command.Keyword.StartsWith("/정치뉴스"))
            {
                WindowsMacro.Instance.SendTextToChatroom(textBox1.Text, $"{News.PoliticsTop6}");
            }
            else if (command.Keyword == "/상식퀴즈")
            {
                ProcessCommonSense();
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
                    WindowsMacro.Instance.SendTextToChatroom(textBox1.Text, $"{answer}");
                }
            }

        }

        //================================================






        protected override void WndProc(ref Message m)
        {
            if (m.Msg == WM_HOTKEY)
            {
                int id = m.WParam.ToInt32();
                if (id == 1)
                {
                    // F5 누르면 실행할 동작
                }

                if (id == 2)
                {

                }
            }
            base.WndProc(ref m);
        }

        private void button1_Click(object sender, System.EventArgs e)
        {

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            RegisterHotKey(this.Handle, 1, 0, (int)Keys.F5);
            RegisterHotKey(this.Handle, 2, 0, (int)Keys.F6);
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            timer.Stop();

            isThreadRunning = false;
            thread.Join();

            UnregisterHotKey(this.Handle, 2);
            UnregisterHotKey(this.Handle, 1);
            Settings.Save(settings);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            isBotRunning = !isBotRunning;
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
            //WindowsMacro.Instance.SendTextToChatroom(textBox1.Text, $"{News.PoliticsTop6}");
            
        }
    }
}
